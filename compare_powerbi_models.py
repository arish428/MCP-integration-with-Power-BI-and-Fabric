from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from collections import defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parent
ADOMD_DIR = ROOT / "adomd_extracted" / "lib" / "net45"
NATIVE_DIR = ROOT / "adomd_extracted" / "runtimes" / "win-x64" / "native"


def configure_adomd() -> None:
    os.environ["PATH"] = f"{ADOMD_DIR};{NATIVE_DIR};{os.environ.get('PATH', '')}"
    if hasattr(os, "add_dll_directory"):
        os.add_dll_directory(str(ADOMD_DIR))
        os.add_dll_directory(str(NATIVE_DIR))
    sys.path.insert(0, str(ADOMD_DIR))

    import clr  # type: ignore

    clr.AddReference("Microsoft.AnalysisServices.AdomdClient")


def run_powershell_json(script: str) -> list[dict[str, object]]:
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", script],
        check=True,
        capture_output=True,
        text=True,
    )
    output = result.stdout.strip()
    if not output:
        return []
    data = json.loads(output)
    if isinstance(data, dict):
        return [data]
    return data


def discover_open_models() -> list[dict[str, str]]:
    script = r'''
    $pbis = Get-CimInstance Win32_Process | Where-Object {
      $_.Name -eq 'PBIDesktop.exe' -and $_.CommandLine -match '\.pbix"?$'
    }
    $ssas = Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'msmdsrv.exe' }

    $rows = foreach ($p in $pbis) {
      $children = $ssas | Where-Object { $_.ParentProcessId -eq $p.ProcessId }
      foreach ($m in $children) {
        $workspace = $null
        if ($m.CommandLine -match '-s "([^"]+)"') {
          $workspace = $matches[1]
        }

        $pbix = $null
        if ($p.CommandLine -match '"([^"]+\.pbix)"\s*$') {
          $pbix = $matches[1]
        }

        $port = $null
        if ($workspace) {
          $portFile = Join-Path $workspace 'msmdsrv.port.txt'
          if (Test-Path $portFile) {
            $raw = Get-Content $portFile -Encoding Unicode -Raw
            $port = (-join ($raw.ToCharArray() | Where-Object { $_ -ne [char]0 })).Trim()
          }
        }

        if ($pbix -and $workspace -and $port) {
          [PSCustomObject]@{
            pbix_path = $pbix
            pbix_process_id = [string]$p.ProcessId
            msmdsrv_process_id = [string]$m.ProcessId
            workspace = $workspace
            port = $port
          }
        }
      }
    }

    $rows | ConvertTo-Json -Compress
    '''
    return [
        {key: str(value) for key, value in row.items()}
        for row in run_powershell_json(script)
    ]


def normalize_value(value: object) -> object:
    if value is None:
        return None
    value_type = type(value).__name__
    if value_type in {"Int16", "Int32", "Int64", "UInt16", "UInt32", "UInt64"}:
        return int(value)
    if value_type in {"Boolean", "bool"}:
        return bool(value)
    return str(value)


def execute_query(connection_string: str, query: str) -> list[dict[str, object]]:
    from Microsoft.AnalysisServices.AdomdClient import AdomdCommand, AdomdConnection  # type: ignore

    connection = AdomdConnection(connection_string)
    connection.Open()
    try:
        command = AdomdCommand(query, connection)
        reader = command.ExecuteReader()
        try:
            columns = [reader.GetName(index) for index in range(reader.FieldCount)]
            rows: list[dict[str, object]] = []
            while reader.Read():
                row: dict[str, object] = {}
                for index, column in enumerate(columns):
                    if reader.IsDBNull(index):
                        row[column] = None
                    else:
                        row[column] = normalize_value(reader.GetValue(index))
                rows.append(row)
            return rows
        finally:
            reader.Close()
    finally:
        connection.Close()


def get_catalog(port: str) -> str:
    rows = execute_query(
        f"Data Source=localhost:{port}",
        "SELECT [CATALOG_NAME] FROM $SYSTEM.DBSCHEMA_CATALOGS",
    )
    if not rows:
        raise RuntimeError(f"No catalog returned for port {port}")
    return str(rows[0]["CATALOG_NAME"])


def fetch_model_metadata(port: str) -> dict[str, object]:
    catalog = get_catalog(port)
    connection_string = f"Data Source=localhost:{port};Catalog={catalog}"

    tables = execute_query(
        connection_string,
        "SELECT [ID], [Name], [IsHidden] FROM $SYSTEM.TMSCHEMA_TABLES",
    )
    columns = execute_query(
        connection_string,
        "SELECT [TableID], [ExplicitName], [IsHidden], [Type] FROM $SYSTEM.TMSCHEMA_COLUMNS",
    )

    table_lookup = {int(row["ID"]): row for row in tables if row.get("ID") is not None}
    column_details: list[dict[str, object]] = []
    for row in columns:
        table_id = row.get("TableID")
        column_name = row.get("ExplicitName")
        if table_id is None or column_name is None:
            continue
        table = table_lookup.get(int(table_id))
        if not table:
            continue
        table_name = str(table["Name"])
        column_details.append(
            {
                "table": table_name,
                "column": str(column_name),
                "table_hidden": bool(table.get("IsHidden", False)),
                "column_hidden": bool(row.get("IsHidden", False)),
                "type": row.get("Type"),
            }
        )

    return {
        "catalog": catalog,
        "tables": [
            {
                "id": int(row["ID"]),
                "name": str(row["Name"]),
                "is_hidden": bool(row.get("IsHidden", False)),
            }
            for row in tables
            if row.get("ID") is not None and row.get("Name") is not None
        ],
        "columns": column_details,
    }


def summarize_columns(columns: list[dict[str, object]]) -> dict[str, object]:
    total = len(columns)
    visible = sum(1 for row in columns if not bool(row["column_hidden"]))
    non_auto_date = [row for row in columns if not str(row["table"]).startswith("LocalDateTable_")]
    visible_non_auto_date = [row for row in non_auto_date if not bool(row["column_hidden"]) and not bool(row["table_hidden"])]

    return {
        "total_columns": total,
        "visible_columns": visible,
        "non_auto_date_columns": len(non_auto_date),
        "visible_non_auto_date_columns": len(visible_non_auto_date),
        "business_column_keys": sorted({f"{row['table']}.{row['column']}" for row in visible_non_auto_date}),
    }


def compare_pbix_models(left_match: str | None = None, right_match: str | None = None) -> dict[str, object]:
    discovered = discover_open_models()
    if not discovered:
        raise RuntimeError("No open PBIX models were discovered.")

    grouped: dict[str, list[dict[str, object]]] = defaultdict(list)
    for row in discovered:
        metadata = fetch_model_metadata(row["port"])
        summary = summarize_columns(metadata["columns"])
        grouped[row["pbix_path"]].append(
            {
                "port": row["port"],
                "workspace": row["workspace"],
                "catalog": metadata["catalog"],
                "table_count": len(metadata["tables"]),
                "summary": summary,
            }
        )

    pbix_paths = sorted(grouped)
    if len(pbix_paths) < 2:
        raise RuntimeError("Fewer than two PBIX files are currently open.")

    def resolve_match(match_value: str, paths: list[str]) -> str:
        normalized = match_value.lower()
        exact = [path for path in paths if path.lower() == normalized]
        if exact:
            return exact[0]
        basename = [path for path in paths if Path(path).name.lower() == normalized]
        if basename:
            return basename[0]
        contains = [path for path in paths if normalized in path.lower()]
        if len(contains) == 1:
            return contains[0]
        if len(contains) > 1:
            raise RuntimeError(f"Match '{match_value}' is ambiguous: {contains}")
        raise RuntimeError(f"Could not find an open PBIX matching '{match_value}'.")

    normalized_models: dict[str, dict[str, object]] = {}
    duplicate_consistency: dict[str, bool] = {}
    for pbix_path, instances in grouped.items():
        signatures = {
            json.dumps(
                {
                    "table_count": instance["table_count"],
                    "summary": instance["summary"],
                },
                sort_keys=True,
            )
            for instance in instances
        }
        duplicate_consistency[pbix_path] = len(signatures) == 1
        normalized_models[pbix_path] = instances[0]

    if left_match and right_match:
        left_path = resolve_match(left_match, pbix_paths)
        right_path = resolve_match(right_match, pbix_paths)
        if left_path == right_path:
            raise RuntimeError("Left and right PBIX resolved to the same file.")
    else:
        left_path, right_path = pbix_paths[:2]

    left_keys = set(normalized_models[left_path]["summary"]["business_column_keys"])
    right_keys = set(normalized_models[right_path]["summary"]["business_column_keys"])

    return {
        "files": {
            left_path: normalized_models[left_path],
            right_path: normalized_models[right_path],
        },
        "duplicate_instances_consistent": duplicate_consistency,
        "differences": {
            "only_in_left": sorted(left_keys - right_keys),
            "only_in_right": sorted(right_keys - left_keys),
            "shared_business_columns": len(left_keys & right_keys),
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare open PBIX semantic model columns.")
    parser.add_argument("--left", help="Left PBIX path, file name, or unique substring.")
    parser.add_argument("--right", help="Right PBIX path, file name, or unique substring.")
    args = parser.parse_args()

    configure_adomd()
    result = compare_pbix_models(args.left, args.right)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()