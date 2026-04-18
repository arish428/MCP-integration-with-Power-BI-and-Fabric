from openpyxl import Workbook
from pathlib import Path


REMOVED_COLUMNS = [
    "parentId",
    "isLeaf",
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
]

NOT_REMOVED_COLUMNS = [
    "status",
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


def autofit_columns(worksheet) -> None:
    for column in worksheet.columns:
        width = max(len(str(cell.value or "")) for cell in column) + 2
        worksheet.column_dimensions[column[0].column_letter].width = min(width, 60)


def main() -> None:
    output_path = Path(__file__).resolve().parent / "risk_taxonomy_removed_vs_not_removed.xlsx"

    workbook = Workbook()
    removed_sheet = workbook.active
    removed_sheet.title = "Removed"
    removed_sheet.append(["Column", "Exists in prod", "Exists in stage", "Status"])
    for column_name in REMOVED_COLUMNS:
        removed_sheet.append([column_name, "Yes", "No", "Removed"])

    not_removed_sheet = workbook.create_sheet("Not Removed")
    not_removed_sheet.append(["Column", "Exists in prod only", "Status"])
    for column_name in NOT_REMOVED_COLUMNS:
        not_removed_sheet.append([column_name, "No", "Not removed"])

    autofit_columns(removed_sheet)
    autofit_columns(not_removed_sheet)
    workbook.save(output_path)
    print(output_path)


if __name__ == "__main__":
    main()