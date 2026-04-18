import os
import json
import pandas as pd
import pyodbc

# ==============================
# CONFIG (UPDATE THESE)
# ==============================

PATH1_FOLDER = r"D:\MCP_Microsoft_Fabric\path1_json"
PATH2_FOLDER = r"D:\MCP_Microsoft_Fabric\path2_json"

SQL_SERVER = "YOUR_SQL_ENDPOINT"
DATABASE = "YOUR_LAKEHOUSE"

# ==============================
# ENSURE FOLDER EXISTS
# ==============================

def ensure_folder(path):
    if not os.path.exists(path):
        print(f"⚠️ Folder not found: {path}")
        print(f"👉 Creating folder: {path}")
        os.makedirs(path)
        return False
    return True

# ==============================
# READ PIPELINE JSON
# ==============================

def extract_pipeline_mappings(folder_path, tag):
    all_data = []

    if not ensure_folder(folder_path):
        print(f"⚠️ No files in {folder_path}. Please add JSON files.")
        return pd.DataFrame()

    files = [f for f in os.listdir(folder_path) if f.endswith(".json")]

    if not files:
        print(f"⚠️ No JSON files found in {folder_path}")
        return pd.DataFrame()

    for file in files:
        file_path = os.path.join(folder_path, file)

        print(f"   → Reading {file}")

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            pipeline_name = file.replace(".json", "")

            activities = data.get("properties", {}).get("activities", [])

            for act in activities:
                if act.get("type") == "Copy":

                    mappings = (
                        act.get("typeProperties", {})
                        .get("translator", {})
                        .get("mappings", [])
                    )

                    for m in mappings:
                        all_data.append({
                            "path": tag,
                            "pipeline": pipeline_name,
                            "source_column": m.get("source", {}).get("name"),
                            "destination_column": m.get("sink", {}).get("name")
                        })

        except Exception as e:
            print(f"❌ Error reading {file}: {e}")

    return pd.DataFrame(all_data)

# ==============================
# READ LAKEHOUSE SCHEMA
# ==============================

def get_lakehouse_columns():
    print("🔹 Connecting to Lakehouse SQL...")

    try:
        conn = pyodbc.connect(
            "Driver={ODBC Driver 18 for SQL Server};"
            f"Server={SQL_SERVER};"
            f"Database={DATABASE};"
            "Authentication=ActiveDirectoryInteractive"
        )

        query = """
        SELECT 
            TABLE_NAME,
            COLUMN_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'dbo'
        """

        df = pd.read_sql(query, conn)

        print("✅ Lakehouse schema fetched")
        return df

    except Exception as e:
        print("❌ SQL Connection Failed:", e)
        return pd.DataFrame()

# ==============================
# VALIDATION
# ==============================

def validate_mappings(pipeline_df, lakehouse_df):

    if pipeline_df.empty or lakehouse_df.empty:
        print("⚠️ Skipping validation (missing data)")
        return pd.DataFrame()

    merged = pipeline_df.merge(
        lakehouse_df,
        left_on="destination_column",
        right_on="COLUMN_NAME",
        how="left"
    )

    merged["status"] = merged["COLUMN_NAME"].apply(
        lambda x: "✅ Match" if pd.notnull(x) else "❌ Missing in Table"
    )

    return merged

# ==============================
# MAIN
# ==============================

def main():

    print("\n🔹 STEP 1: Reading Pipeline JSONs")

    path1_df = extract_pipeline_mappings(PATH1_FOLDER, "PATH1")
    path2_df = extract_pipeline_mappings(PATH2_FOLDER, "PATH2")

    pipeline_df = pd.concat([path1_df, path2_df], ignore_index=True)

    if pipeline_df.empty:
        print("\n❌ No pipeline mappings found.")
        print("👉 Add JSON files into folders and re-run.")
        return

    print("\n🔹 STEP 2: Reading Lakehouse Schema")

    lakehouse_df = get_lakehouse_columns()

    print("\n🔹 STEP 3: Validation")

    validation_df = validate_mappings(pipeline_df, lakehouse_df)

    # ==============================
    # EXPORT
    # ==============================

    pipeline_df.to_csv("pipeline_mappings.csv", index=False)

    if not lakehouse_df.empty:
        lakehouse_df.to_csv("lakehouse_columns.csv", index=False)

    if not validation_df.empty:
        validation_df.to_csv("validation_report.csv", index=False)

    print("\n✅ DONE — Files Generated")

# ==============================
# RUN
# ==============================

if __name__ == "__main__":
    main()