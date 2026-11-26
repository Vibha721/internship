"""Read study metadata and the result fields used by the supplied report template."""

from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd

from .metadata import read_metadata_file


_RESULT_ALIASES = {
    "r1": "replicate_1",
    "r1star": "replicate_1",
    "replicate1": "replicate_1",
    "replicate1star": "replicate_1",
    "r2": "replicate_2",
    "r2star": "replicate_2",
    "replicate2": "replicate_2",
    "replicate2star": "replicate_2",
    "r3": "replicate_3",
    "r3star": "replicate_3",
    "replicate3": "replicate_3",
    "replicate3star": "replicate_3",
    "grade": "grade",
    "gradestar": "grade",
    "result": "result",
    "resultstar": "result",
}


def _key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def _text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_report_workbook(filename: str, content: bytes) -> tuple[dict[str, str], dict[str, str]]:
    """Read all workbook sheets for the report's metadata and result values.

    The supplied report template expects the key/value rows used in the
    Cytotoxicity Study workbook (TIDS, in-house study details, and R1/R2/R3,
    Grade, Result). No assay measurements are inferred from missing values.
    """

    suffix = Path(filename).suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        raise ValueError("Upload a .xlsx or .xlsm workbook for this report template.")

    metadata = read_metadata_file(filename, content)
    workbook = pd.ExcelFile(BytesIO(content))
    results: dict[str, str] = {
        "replicate_1": "",
        "replicate_2": "",
        "replicate_3": "",
        "grade": "",
        "result": "",
    }
    for sheet_name in workbook.sheet_names:
        frame = pd.read_excel(workbook, sheet_name=sheet_name, header=None)
        for _, row in frame.iterrows():
            values = list(row)
            for column, label in enumerate(values):
                result_field = _RESULT_ALIASES.get(_key(label))
                if not result_field:
                    continue
                result_value = next(
                    (_text(value) for value in values[column + 1 :] if _text(value)),
                    "",
                )
                if result_value:
                    results.setdefault(result_field, result_value)

        for column in range(frame.shape[1]):
            for row_index in range(frame.shape[0] - 1):
                label = frame.iloc[row_index, column]
                result_field = _RESULT_ALIASES.get(_key(label))
                if not result_field:
                    continue
                result_value = _text(frame.iloc[row_index + 1, column])
                if result_value:
                    results.setdefault(result_field, result_value)

    details = metadata.get("test_item_extraction_details", "")
    weight_match = re.search(
        r"weight\s*(?:\([^)]*\))?\s*:\s*([^;\n]+)", details, re.IGNORECASE
    )
    if weight_match and not metadata.get("test_item_weight"):
        metadata["test_item_weight"] = weight_match.group(1).strip()

    return metadata, results


def validate_report_results(results: dict[str, str]) -> list[str]:
    """Validate the fields and result pattern assumed by the provided report."""

    required = ("replicate_1", "replicate_2", "replicate_3", "grade", "result")
    missing = [field for field in required if not str(results.get(field, "")).strip()]
    if missing:
        return []

    expected_replicates = ["1", "1", "1"]
    actual_replicates = [str(results[field]).strip() for field in required[:3]]
    normalized_grade = str(results["grade"]).strip().casefold()
    normalized_result = str(results["result"]).strip().casefold().replace(" ", "-")
    if actual_replicates != expected_replicates or normalized_grade != "slight" or normalized_result != "non-cytotoxic":
        return [
            "This supplied Study Report template contains fixed Agarose Overlay results text for three Grade 1 (Slight) replicates and a non-cytotoxic conclusion. The uploaded Excel results do not match that text, so the report was not generated to avoid a misleading result."
        ]
    return []
