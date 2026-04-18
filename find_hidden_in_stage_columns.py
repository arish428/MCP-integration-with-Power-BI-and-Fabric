from pathlib import Path
import json

import compare_powerbi_models as model_compare


REMOVED_COLUMNS = {
    "parentId",
    "isLeaf",
    "status",
    "createdBy",
    "updatedBy",
    "updatedDate",
    "isDeleted",
    "isSystemRiskRegisterItem",
    "isOrphan",
    "riskDefinitionId",
    "libraryRiskId",
    "riskUpdateStatus",
    "riskTaxonomyId",
    "processTaxonomyId",
    "contentLibraryId",
    "guid",
    "isOrphanForProcess",
    "managementComments",
    "referencedRiskRegisterItemId",
    "riaSurveyId",
    "rpaSurveyId",
    "controlSurveyId",
    "etl_uuid",
    "comment",
    "cfrGuidance",
    "predict360.businessareasdefinitionriskregisteritemmapping",
    "predict360.contentlibrary",
    "predict360.customer",
    "predict360.kri",
    "predict360.predictedriskforriskcategory",
    "predict360.predictuser(createdBy)",
    "predict360.predictuser(updatedBy)",
    "predict360.processcategoryriskitem",
    "predict360.processtaxonomy",
    "predict360.riskappetiteregisteritem",
    "predict360.riskcategorypredictedrisk",
    "predict360.riskregister",
    "predict360.riskregisteritem(id)",
    "predict360.riskregisteritem(id) 2",
    "predict360.riskregisteritem(parentId)",
    "predict360.riskregisteritem(referencedRiskRegisterItemId)",
    "predict360.riskregisteritemapplicablity(id)",
    "predict360.riskregisteritemapplicablity(id) 2",
    "predict360.riskregisteritemmapping(id)",
    "predict360.riskregisteritemmapping(id) 2",
    "predict360.riskregisteritemnotification",
    "predict360.riskregisteritemrequirement",
    "predict360.riskregisteritemsubjarea",
    "predict360.riskregistertask",
    "predict360.risktaxonomy",
    "predict360.risktaxonomyerrorcontentlog(id)",
    "predict360.risktaxonomyerrorcontentlog(id) 2",
    "predict360.sboriskregisteritem",
}


def main() -> None:
    adomd_dir = Path(r"C:\Program Files\Microsoft.NET\ADOMD.NET\160")
    model_compare.ADOMD_DIR = adomd_dir
    model_compare.NATIVE_DIR = adomd_dir
    model_compare.configure_adomd()

    discovered = model_compare.discover_open_models()
    selected: dict[str, list[dict[str, object]]] = {}
    for row in discovered:
        pbix_path = row["pbix_path"]
        if pbix_path.endswith("Risk Taxonomy Summary Report (prod).pbix") or pbix_path.endswith(
            "Risk Taxonomy Summary Report (stage).pbix"
        ):
            selected[pbix_path] = model_compare.fetch_model_metadata(row["port"])["columns"]

    prod_path = next(path for path in selected if path.endswith("(prod).pbix"))
    stage_path = next(path for path in selected if path.endswith("(stage).pbix"))

    prod_columns = {f"{column['table']}.{column['column']}": column for column in selected[prod_path]}
    stage_columns = {f"{column['table']}.{column['column']}": column for column in selected[stage_path]}

    hidden_in_stage = []
    for qualified_name, stage_column in stage_columns.items():
        bare_name = qualified_name.split(".", 1)[1]
        if bare_name in REMOVED_COLUMNS and qualified_name in prod_columns and bool(stage_column["column_hidden"]):
            hidden_in_stage.append(
                {
                    "qualified": qualified_name,
                    "column": bare_name,
                    "stage_hidden": bool(stage_column["column_hidden"]),
                    "prod_hidden": bool(prod_columns[qualified_name]["column_hidden"]),
                }
            )

    print(json.dumps(sorted(hidden_in_stage, key=lambda item: item["qualified"]), indent=2))


if __name__ == "__main__":
    main()