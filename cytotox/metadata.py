"""Read sponsor and study metadata from simple Excel, Word, or JSON files."""

from __future__ import annotations

import json
import re
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from docx import Document


FIELDS = {
    "sponsor_name": "Sponsor Name",
    "sponsor_address": "Sponsor Address",
    "sponsor_contact": "Sponsor Contact",
    "study_title": "Study Title",
    "study_number": "Study Number",
    "study_type": "Study Type",
    "study_objective": "Study Objective",
    "test_item_name": "Test Item Name",
    "test_item_id": "Test Item ID",
    "batch_lot": "Batch / Lot Number",
    "vehicle": "Vehicle / Solvent",
    "test_system": "Test System",
    "study_start": "Study Start Date",
    "study_end": "Study End Date",
    "method": "Experimental Method / Procedure",
    "acceptance_criteria": "Acceptance Criteria",
    "deviations": "Deviations / Observations",
    "conclusion": "Conclusion",
    "prepared_by": "Prepared By",
    "reviewed_by": "Reviewed By",
    "study_director": "Study Director",
    "study_director_name": "Study Director Name",
    "quality_assurance_unit": "Quality Assurance Unit",
    "test_facility_manager": "Test Facility Manager",
    "study_completion_date": "Study Completion Date",
    # Fields used by the supplied cytotoxicity study-plan template.
    "study_code": "Study Code",
    "monitoring_scientist": "Monitoring Scientist",
    "study_coordinator": "Study Co-ordinator",
    "study_personnel": "Study Personnel",
    "study_initiation_date": "Study Initiation Date",
    "experimental_start_date": "Experimental Starting Date",
    "experimental_completion_date": "Experimental Completion Date",
    "draft_report_submission_date": "Draft Report Submission Date",
    "study_director_email": "Study Director Email",
    "test_item_extraction_details": "Test Item Extraction Details",
    "test_item_weight": "Test Item Weight",
    "physical_appearance": "Physical Appearance",
    "quantity_sent": "Quantity Sent",
    "batch_number": "Batch Number",
    "lot_number": "Lot Number",
    "manufacturer_name": "Manufacturer Name",
    "manufacturer_address": "Manufacturer Address",
    "manufacture_date": "Date of Manufacture",
    "expiry_date": "Date of Expiry",
    "storage_condition": "Storage Condition",
    "test_item_code": "Test Item Code",
}


def _key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


_ALIASES = {_key(label): field for field, label in FIELDS.items()}
_ALIASES.update(
    {
        "sponsor": "sponsor_name",
        "studysponsor": "sponsor_name",
        "sponsorcompany": "sponsor_name",
        "address": "sponsor_address",
        "sponsorrepresentative": "sponsor_contact",
        "studysponsornameandaddress": "sponsor_name",
        "sponsoremail": "sponsor_contact",
        "sponsorrepr": "sponsor_contact",
        "studyid": "study_number",
        "protocolnumber": "study_number",
        "protocoltitle": "study_title",
        "testitem": "test_item_name",
        "testsample": "test_item_name",
        "nameoftestitem": "test_item_name",
        "nameoftestitemreferenceitem": "test_item_name",
        "sampleid": "test_item_id",
        "testitemcodebytestfacility": "test_item_id",
        "lotnumber": "batch_lot",
        "batchnumber": "batch_lot",
        "lotno": "batch_lot",
        "solvent": "vehicle",
        "cellline": "test_system",
        "procedure": "method",
        "methodology": "method",
        "clinicalintendeduseofthedevicetestitem": "study_objective",
        "nameofthesponsorsrepresentative": "sponsor_contact",
        "studydirector": "prepared_by",
        "studydirectorname": "study_director_name",
        "nameofstudydirector": "study_director_name",
        "directorname": "study_director_name",
        "qualityassuranceunit": "quality_assurance_unit",
        "qualityassurance": "quality_assurance_unit",
        "qau": "quality_assurance_unit",
        "qaunit": "quality_assurance_unit",
        "qualityassurancedepartment": "quality_assurance_unit",
        "qualityassurancepersonnel": "quality_assurance_unit",
        "testfacilityman": "test_facility_manager",
        "testfacilitymanager": "test_facility_manager",
        "studycompletiondate": "study_completion_date",
        "studydirectoremail": "study_director_email",
        "studycode": "study_code",
        "monitoringscientist": "monitoring_scientist",
        "studycoordinator": "study_coordinator",
        "studypersonnel": "study_personnel",
        "studyinitiationdate": "study_initiation_date",
        "experimentalstartingdate": "experimental_start_date",
        "experimentalcompletiondate": "experimental_completion_date",
        "draftreportsubmissiondate": "draft_report_submission_date",
        "detailstobeconsideredforextractiontesting": "test_item_extraction_details",
        "detailstobeconsidered": "test_item_extraction_details",
        "extractiondetails": "test_item_extraction_details",
        "weight": "test_item_weight",
        "quantitysent": "quantity_sent",
        "batchno": "batch_number",
        "lotno": "lot_number",
        "manufacturedsuppliedbynameandaddress": "manufacturer_name",
        "manufacturedsuppliedby": "manufacturer_name",
        "manufacturedsupplied": "manufacturer_name",
        "manufacturer": "manufacturer_name",
        "dateofmanufacture": "manufacture_date",
        "dateofexpiry": "expiry_date",
        "storagecondition": "storage_condition",
        "testitemcode": "test_item_code",
        "testitemcodebytestfacility": "test_item_code",
        "preparedby": "prepared_by",
        "reviewedby": "reviewed_by",
    }
)


def _assign(target: dict[str, str], label: Any, value: Any) -> None:
    label_key = _key(label)
    if isinstance(value, (dict, list, tuple, set)):
        return
    if bool(pd.isna(value)):
        return
    if isinstance(value, (pd.Timestamp, datetime, date)):
        text = value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    else:
        text = str(value).strip()
    if text and text.lower() != "nan":
        if label_key == "studysponsornameandaddress":
            parts = [part.strip() for part in text.splitlines() if part.strip()]
            if parts:
                target["sponsor_name"] = parts[0]
                if len(parts) > 1:
                    target["sponsor_address"] = "\n".join(parts[1:])
            return
        if label_key == "manufacturedsuppliedbynameandaddress":
            parts = [part.strip() for part in text.splitlines() if part.strip()]
            if parts:
                target["manufacturer_name"] = parts[0]
                if len(parts) > 1:
                    target["manufacturer_address"] = "\n".join(parts[1:])
            else:
                target["manufacturer_name"] = text
            return
        if label_key in {"studydirector", "studydirectorname", "nameofstudydirector", "directorname"}:
            target["prepared_by"] = text
            target["study_director"] = text
            target["study_director_name"] = text
            return
        if label_key == "testitemcodebytestfacility":
            target["test_item_code"] = text
            target["test_item_id"] = text
            return
        if label_key in {"lotnumber", "lotno"}:
            target["lot_number"] = text
            target["batch_lot"] = text
            return
        if label_key in {"batchnumber", "batchno"}:
            target["batch_number"] = text
            target["batch_lot"] = text
            return
        field = _ALIASES.get(label_key)
        if field is not None:
            target[field] = text


def read_metadata_file(filename: str, content: bytes) -> dict[str, str]:
    """Read recognized fields from key/value Excel, Word, or JSON files.

    Excel files can use key/value columns or a study-details table with
    Description and Details columns. Word files use two-column tables or
    paragraphs formatted as ``Field: value``.
    """

    suffix = Path(filename).suffix.lower()
    values: dict[str, str] = {}

    if suffix == ".json":
        payload = json.loads(content.decode("utf-8-sig"))
        if not isinstance(payload, dict):
            raise ValueError("The JSON file must contain a key/value object.")
        for label, value in payload.items():
            _assign(values, label, value)
        return values

    if suffix in {".xlsx", ".xlsm"}:
        workbook = pd.ExcelFile(BytesIO(content))
        sheet_names = list(workbook.sheet_names)
        if "Metadata" in sheet_names:
            sheet_names.remove("Metadata")
            sheet_names.insert(0, "Metadata")
        parsed_any = False
        for sheet in sheet_names:
            frame = pd.read_excel(
                workbook,
                sheet_name=sheet,
                header=None,
                keep_default_na=False,
            )
            frame = frame.replace("", pd.NA).dropna(how="all")
            if frame.shape[1] < 2:
                continue

            # Study-details sheets often have an ID column followed by
            # Description and Details, with a title and blank rows above.
            found_description_table = False
            for row_index, row in frame.iterrows():
                header_columns = {_key(value): column for column, value in row.items() if not pd.isna(value)}
                description_column = header_columns.get("description")
                details_column = header_columns.get("details")
                if description_column is not None and details_column is not None:
                    for _, data_row in frame.loc[row_index + 1 :].iterrows():
                        _assign(values, data_row.iloc[description_column], data_row.iloc[details_column])
                    parsed_any = True
                    found_description_table = True
                    break
            if found_description_table:
                continue

            for row in frame.itertuples(index=False, name=None):
                if len(row) >= 2:
                    before = len(values)
                    _assign(values, row[0], row[1])
                    parsed_any = parsed_any or len(values) > before

        if not parsed_any:
            raise ValueError("The metadata worksheet must have a Field and Value column.")
        return values

    if suffix == ".docx":
        document = Document(BytesIO(content))
        for table in document.tables:
            for row in table.rows:
                if len(row.cells) >= 2:
                    _assign(values, row.cells[0].text, row.cells[1].text)
        for paragraph in document.paragraphs:
            text = paragraph.text.strip()
            if ":" in text:
                label, value = text.split(":", 1)
                _assign(values, label, value)
        return values

    raise ValueError("Use a .xlsx, .xlsm, .docx, or .json metadata file.")
