"""Experimental spreadsheet templates and input validation."""

from __future__ import annotations

from io import BytesIO
import re

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

from .metadata import FIELDS


TEMPLATE_COLUMNS = ["Treatment", "Concentration", "Unit", "Replicate", "Response", "Response Unit"]

_COLUMN_ALIASES = {
    "treatment": "Treatment",
    "group": "Treatment",
    "sample": "Treatment",
    "testitem": "Treatment",
    "concentration": "Concentration",
    "conc": "Concentration",
    "dose": "Concentration",
    "doselevel": "Concentration",
    "unit": "Unit",
    "concentrationunit": "Unit",
    "doseunit": "Unit",
    "replicate": "Replicate",
    "rep": "Replicate",
    "replicatenumber": "Replicate",
    "response": "Response",
    "responsepercent": "Response",
    "viability": "Response",
    "cellviability": "Response",
    "cellviabilitypercent": "Response",
    "percentviability": "Response",
    "viabilitypercent": "Response",
    "result": "Response",
    "value": "Response",
    "responseunit": "Response Unit",
    "responseunits": "Response Unit",
}


def _normalize_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def create_input_template() -> bytes:
    """Create a clearly labeled blank workbook for the experimental data."""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Experimental Data"
    sheet.append(TEMPLATE_COLUMNS)
    sheet.append(["", "", "", "", "", ""])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = "A1:F2"

    for cell in sheet[1]:
        cell.font = Font(name="Aptos", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="183B56")
        cell.alignment = Alignment(vertical="center")
    sheet.row_dimensions[1].height = 25

    for column, width in {"A": 24, "B": 18, "C": 16, "D": 16, "E": 18, "F": 18}.items():
        sheet.column_dimensions[column].width = width

    table = Table(displayName="ExperimentalData", ref="A1:F2")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)
    sheet.sheet_view.showGridLines = False

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def create_metadata_template() -> bytes:
    """Create a two-column workbook for sponsor and study details."""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Metadata"
    sheet.append(["Field", "Value"])
    for label in FIELDS.values():
        sheet.append([label, ""])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:B{sheet.max_row}"
    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["B"].width = 64
    sheet.sheet_view.showGridLines = False
    for cell in sheet[1]:
        cell.font = Font(name="Aptos", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="183B56")
        cell.alignment = Alignment(vertical="center")
    sheet.row_dimensions[1].height = 25
    table = Table(displayName="StudyMetadata", ref=f"A1:B{sheet.max_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def validate_experimental_data(frame: pd.DataFrame) -> tuple[pd.DataFrame, list[str], list[str]]:
    """Normalize common headers and validate the required assay data fields."""

    data = frame.copy()
    data.columns = [str(column).strip() for column in data.columns]
    renamed: dict[object, str] = {}
    for column in data.columns:
        canonical = _COLUMN_ALIASES.get(_normalize_header(column))
        if canonical and canonical not in renamed.values():
            renamed[column] = canonical
    data = data.rename(columns=renamed).dropna(how="all").reset_index(drop=True)

    errors: list[str] = []
    warnings: list[str] = []
    required = ["Treatment", "Concentration", "Replicate", "Response"]
    missing = [column for column in required if column not in data.columns]
    if missing:
        errors.append("Missing required column(s): " + ", ".join(missing) + ".")
        return data, errors, warnings
    if data.empty:
        errors.append("The worksheet has column headers but no experimental rows.")
        return data, errors, warnings

    for column in ("Concentration", "Response"):
        original = data[column]
        converted = pd.to_numeric(original, errors="coerce")
        invalid = original.notna() & converted.isna()
        if invalid.any():
            examples = ", ".join(str(value) for value in original[invalid].head(3).tolist())
            errors.append(f"{column} must be numeric. Invalid value(s) include: {examples}.")
        non_finite = converted.isin([float("inf"), float("-inf")])
        if non_finite.any():
            errors.append(f"{column} must contain finite numeric values.")
        data[column] = converted

    for column in required:
        missing_rows = data[column].isna() | data[column].astype(str).str.strip().eq("")
        if missing_rows.any():
            errors.append(f"{int(missing_rows.sum())} row(s) are missing {column}.")

    if data["Concentration"].notna().any() and (data["Concentration"].dropna() < 0).any():
        errors.append("Concentration cannot be negative.")

    if "Unit" not in data.columns:
        data["Unit"] = ""
    data["Unit"] = data["Unit"].fillna("").astype(str).str.strip()
    data["Treatment"] = data["Treatment"].fillna("").astype(str).str.strip()
    data["Replicate"] = data["Replicate"].fillna("").astype(str).str.strip()

    all_concentration_units = set(data["Unit"].unique())
    concentration_units = all_concentration_units - {""}
    if len(concentration_units) > 1:
        errors.append("Use one concentration unit per workbook. Convert values before combining units.")
    elif concentration_units and "" in all_concentration_units:
        errors.append("Some rows are missing a concentration unit while others provide one.")
    elif not concentration_units:
        warnings.append("No concentration unit was provided. Confirm units before interpreting the graph or report.")
    if "Response Unit" in data.columns:
        data["Response Unit"] = data["Response Unit"].fillna("").astype(str).str.strip()
        all_response_units = set(data["Response Unit"].unique())
        response_units = all_response_units - {""}
        if len(response_units) > 1:
            errors.append("Use one response unit per workbook. Convert values before combining units.")
        elif response_units and "" in all_response_units:
            errors.append("Some rows are missing a response unit while others provide one.")
        elif not response_units:
            warnings.append("No response unit was provided. Confirm units before interpreting the graph or report.")
    else:
        warnings.append("No response unit was provided. Confirm units before interpreting the graph or report.")

    responses = data["Response"].dropna()
    if not responses.empty and (~responses.between(0, 100)).any():
        warnings.append("Some response values are outside 0–100. Confirm the response unit and interpretation.")

    duplicate_keys = ["Treatment", "Concentration", "Replicate"]
    if data[duplicate_keys].duplicated().any():
        count = int(data[duplicate_keys].duplicated(keep=False).sum())
        warnings.append(f"{count} rows share the same Treatment, Concentration, and Replicate. Check for duplicate records.")

    return data, errors, warnings
