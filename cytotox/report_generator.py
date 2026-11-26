"""Build standardized Microsoft Word study plans and study reports."""

from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Any

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from .metadata import FIELDS
from .report_template_filler import fill_study_report_template
from .template_filler import fill_study_plan_template


NAVY = "183B56"
TEAL = "168C8C"
PALE = "EAF2F5"
MUTED = "52616B"


def _value(metadata: dict[str, Any], key: str, fallback: str = "Not provided") -> str:
    value = metadata.get(key)
    return str(value).strip() if value is not None and str(value).strip() else fallback


def _shade(cell: Any, fill: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    properties.append(shading)


def _set_cell_text(cell: Any, value: Any, *, bold: bool = False, color: str | None = None) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(str(value) if value is not None else "")
    run.bold = bold
    run.font.name = "Aptos"
    run.font.size = Pt(9)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _add_table(document: Document, headers: list[str], rows: list[list[Any]]) -> Any:
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for index, header in enumerate(headers):
        _set_cell_text(table.rows[0].cells[index], header, bold=True, color="FFFFFF")
        _shade(table.rows[0].cells[index], NAVY)
    for row_index, row in enumerate(rows):
        cells = table.add_row().cells
        for column_index, value in enumerate(row):
            _set_cell_text(cells[column_index], value)
            if row_index % 2 == 1:
                _shade(cells[column_index], "F4F7F8")
    document.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def _add_section(document: Document, title: str, body: str | None = None) -> None:
    document.add_heading(title, level=1)
    if body and body.strip():
        paragraph = document.add_paragraph(body.strip())
        paragraph.paragraph_format.space_after = Pt(7)


def _add_page_number(paragraph: Any) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.add_run("Page ")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run = paragraph.add_run()
    run._r.append(begin)
    run._r.append(instruction)
    run._r.append(separate)
    run._r.append(end)


def _setup_document(document: Document, kind: str) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    section.header_distance = Inches(0.35)
    section.footer_distance = Inches(0.35)

    normal = document.styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(9.5)
    normal.font.color.rgb = RGBColor.from_string("263746")
    normal.paragraph_format.space_after = Pt(5)

    for style_name, size in (("Title", 23), ("Heading 1", 14), ("Heading 2", 11)):
        style = document.styles[style_name]
        style.font.name = "Aptos Display" if style_name == "Title" else "Aptos"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(NAVY if style_name != "Heading 2" else TEAL)
        style.paragraph_format.keep_with_next = True

    header = section.header.paragraphs[0]
    header.text = f"CYTOTOXICITY STUDY  |  {kind.upper()}"
    header.style = document.styles["Caption"]
    header.runs[0].font.color.rgb = RGBColor.from_string(MUTED)
    header.runs[0].font.bold = True

    footer = section.footer.paragraphs[0]
    footer.add_run("Generated study document  •  ")
    footer.runs[0].font.size = Pt(8)
    footer.runs[0].font.color.rgb = RGBColor.from_string(MUTED)
    _add_page_number(footer)


def _add_front_matter(document: Document, kind: str, metadata: dict[str, Any]) -> None:
    title = _value(metadata, "study_title", "Cytotoxicity Study")
    title_para = document.add_paragraph(style="Title")
    title_para.paragraph_format.space_before = Pt(14)
    title_para.add_run(title)
    subtitle = document.add_paragraph()
    run = subtitle.add_run(f"Study {kind}")
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor.from_string(TEAL)

    rows = [
        ["Study Number", _value(metadata, "study_number")],
        ["Test Item", _value(metadata, "test_item_name")],
        ["Sponsor", _value(metadata, "sponsor_name")],
        ["Generated", datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %Z")],
    ]
    table = document.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for row_index, (label, value) in enumerate(rows):
        cells = table.add_row().cells
        _set_cell_text(cells[0], label, bold=True, color=NAVY)
        _shade(cells[0], PALE)
        _set_cell_text(cells[1], value)
        if row_index % 2 == 1:
            _shade(cells[1], "F8FAFB")
    document.add_paragraph()


def _add_metadata_sections(document: Document, metadata: dict[str, Any]) -> None:
    _add_section(document, "Sponsor Details")
    _add_table(
        document,
        ["Field", "Details"],
        [[FIELDS[field], _value(metadata, field)] for field in ("sponsor_name", "sponsor_address", "sponsor_contact")],
    )

    _add_section(document, "Test Item Details")
    _add_table(
        document,
        ["Field", "Details"],
        [[FIELDS[field], _value(metadata, field)] for field in ("test_item_name", "test_item_id", "batch_lot", "vehicle")],
    )

    _add_section(document, "Study Information")
    _add_table(
        document,
        ["Field", "Details"],
        [[FIELDS[field], _value(metadata, field)] for field in ("study_number", "study_type", "study_start", "study_end", "test_system", "prepared_by", "reviewed_by")],
    )
    _add_section(document, "Study Objective", _value(metadata, "study_objective"))
    _add_section(document, "Experimental Method / Procedure", _value(metadata, "method"))
    _add_section(document, "Acceptance Criteria", _value(metadata, "acceptance_criteria"))


def _value_from_data(data: pd.DataFrame, column: str, fallback: str) -> str:
    if column not in data.columns:
        return fallback
    values = data[column].dropna().astype(str).str.strip().unique()
    return values[0] if len(values) == 1 and values[0] else fallback


def _summary_rows(data: pd.DataFrame) -> list[list[str]]:
    grouped = data.groupby(["Treatment", "Concentration", "Unit"], dropna=False, sort=True)["Response"]
    summary = grouped.agg(["count", "mean", "std"]).reset_index()
    rows: list[list[str]] = []
    response_unit = _value_from_data(data, "Response Unit", "response units")
    for treatment, concentration, unit, count, mean, sd in summary.itertuples(index=False, name=None):
        concentration_text = f"{concentration:g}" if pd.notna(concentration) else ""
        rows.append(
            [
                str(treatment),
                f"{concentration_text} {unit}".strip(),
                str(int(count)),
                f"{mean:.3f} {response_unit}" if pd.notna(mean) else "",
                f"{sd:.3f} {response_unit}" if pd.notna(sd) else "Not calculated (n=1)",
            ]
        )
    return rows


def _chart_bytes(data: pd.DataFrame) -> bytes:
    fig, axis = plt.subplots(figsize=(7.2, 3.7), dpi=180)
    palette = ["#168C8C", "#E28B43", "#547AA5", "#986FA8", "#54A47B", "#D36565"]
    for index, (treatment, subset) in enumerate(data.groupby("Treatment", sort=False)):
        summary = subset.groupby("Concentration", sort=True)["Response"].agg(["mean", "std", "count"])
        axis.errorbar(
            summary.index,
            summary["mean"],
            yerr=summary["std"].fillna(0),
            marker="o",
            markersize=4,
            linewidth=1.7,
            capsize=3,
            color=palette[index % len(palette)],
            label=str(treatment),
        )
    axis.set_xlabel(f"Concentration ({_value_from_data(data, 'Unit', 'unit')})")
    axis.set_ylabel(f"Mean response ({_value_from_data(data, 'Response Unit', 'response unit')})")
    axis.set_title("Response by concentration", loc="left", color="#183B56", fontsize=11, pad=10)
    axis.spines[["top", "right"]].set_visible(False)
    axis.grid(axis="y", color="#DCE5E9", linewidth=0.7)
    if data["Treatment"].nunique() > 1:
        axis.legend(frameon=False, loc="best")
    fig.tight_layout()
    output = BytesIO()
    fig.savefig(output, format="png", bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return output.getvalue()


def _format_raw_value(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def build_study_document(
    kind: str,
    metadata: dict[str, Any],
    data: pd.DataFrame | None = None,
    source_filename: str | None = None,
    template_bytes: bytes | None = None,
    report_results: dict[str, Any] | None = None,
) -> bytes:
    """Return a study plan or report as a Microsoft Word document."""

    if kind not in {"Study Plan", "Study Report"}:
        raise ValueError("Document type must be Study Plan or Study Report.")
    if kind == "Study Plan" and template_bytes is not None:
        return fill_study_plan_template(template_bytes, metadata)
    if kind == "Study Report" and template_bytes is not None:
        if report_results is None:
            raise ValueError("The supplied Study Report template requires report result fields from the Excel workbook.")
        return fill_study_report_template(template_bytes, metadata, report_results)
    if kind == "Study Report" and (data is None or data.empty):
        raise ValueError("A study report requires validated experimental data.")

    document = Document()
    _setup_document(document, kind)
    _add_front_matter(document, kind, metadata)
    _add_metadata_sections(document, metadata)

    if kind == "Study Plan":
        _add_section(
            document,
            "Planned Data Collection",
            "Record each treatment, concentration and unit, replicate identifier, and measured response in the experimental data workbook. Preserve source observations and units when transferring data.",
        )
        _add_section(
            document,
            "Data Review and Reporting",
            "Review records for missing values, inconsistent units, and duplicate treatment/concentration/replicate combinations before analysis. Summarize response by treatment and concentration. Document any exclusions or deviations with a rationale.",
        )
        _add_section(document, "Deviations / Observations", _value(metadata, "deviations"))
        _add_section(document, "Approval")
        _add_table(
            document,
            ["Role", "Name", "Signature / Date"],
            [["Prepared by", _value(metadata, "prepared_by"), ""], ["Reviewed by", _value(metadata, "reviewed_by"), ""]],
        )
    else:
        assert data is not None
        _add_section(document, "Experimental Data Source")
        _add_table(
            document,
            ["Source file", "Records", "Response unit"],
            [[source_filename or "Uploaded workbook", str(len(data)), _value_from_data(data, "Response Unit", "Not provided")]],
        )
        _add_section(
            document,
            "Results and Observations",
            "The following table and graph summarize the observed responses by treatment and concentration using descriptive statistics.",
        )
        _add_table(
            document,
            ["Treatment", "Concentration", "n", "Mean response", "SD"],
            _summary_rows(data),
        )
        document.add_picture(BytesIO(_chart_bytes(data)), width=Inches(6.2))
        caption = document.add_paragraph("Points show mean response; error bars show standard deviation.")
        caption.style = document.styles["Caption"]
        _add_section(document, "Deviations / Observations", _value(metadata, "deviations"))
        _add_section(document, "Conclusion", _value(metadata, "conclusion"))
        _add_section(document, "Raw Experimental Data")
        visible = data
        _add_table(
            document,
            [str(column) for column in visible.columns],
            [[_format_raw_value(value) for value in row] for row in visible.itertuples(index=False, name=None)],
        )
    output = BytesIO()
    document.save(output)
    return output.getvalue()
