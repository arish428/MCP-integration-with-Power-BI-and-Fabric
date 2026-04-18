from pathlib import Path
import json

import compare_powerbi_models as model_compare


REMOVED_COLUMNS = [
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
]


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

    prod_bare_columns = {str(column["column"]) for column in selected[prod_path]}
    stage_bare_columns = {str(column["column"]) for column in selected[stage_path]}

    missing_in_stage = [
        column_name
        for column_name in REMOVED_COLUMNS
        if column_name in prod_bare_columns and column_name not in stage_bare_columns
    ]
    present_in_both = [
        column_name
        for column_name in REMOVED_COLUMNS
        if column_name in prod_bare_columns and column_name in stage_bare_columns
    ]

    print(
        json.dumps(
            {
                "missing_in_stage": missing_in_stage,
                "present_in_both": present_in_both,
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()