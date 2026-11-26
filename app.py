from __future__ import annotations

from datetime import date
import hashlib
import importlib
import json
from io import BytesIO
from pathlib import Path
import re

import pandas as pd
import streamlit as st

import cytotox.preview as preview_renderer
from cytotox.excel_io import create_input_template, create_metadata_template, validate_experimental_data
from cytotox.metadata import FIELDS, read_metadata_file
from cytotox.report_excel import read_report_workbook, validate_report_results
from cytotox.report_generator import build_study_document
from cytotox.template_input import normalize_template_upload


preview_renderer = importlib.reload(preview_renderer)
PREVIEW_RENDERER_VERSION = preview_renderer.PREVIEW_RENDERER_VERSION
render_docx_pdf = preview_renderer.render_docx_pdf


st.set_page_config(
    page_title="Cytotoxicity Study Document Generator",
    page_icon="🧪",
    layout="wide",
)

PROJECT_DIR = Path(__file__).resolve().parent
SUPPLIED_DCA_TEMPLATE = PROJECT_DIR / "templates" / "Cytotoxicity - SP - DCA - Template.docx"
SUPPLIED_TEMPLATE = PROJECT_DIR / "templates" / "Cytotoxicity Study Plan Template.docx"
SUPPLIED_REPORT_TEMPLATE = PROJECT_DIR / "templates" / "Cytotoxicity Study Report Template.docx"
METADATA_PARSER_VERSION = "8"
DOCUMENT_GENERATOR_VERSION = "7"
st.session_state.setdefault("app_page", "template")

st.markdown(
    """
    <style>
      .block-container {max-width: 1180px; padding-top: 2rem; padding-bottom: 3rem;}
      .hero {padding: 1.35rem 1.6rem; border-radius: 14px; background: linear-gradient(110deg,#183b56,#168c8c); color: white; margin-bottom: 1.35rem;}
      .hero h1 {color: white; margin: 0 0 .35rem 0; font-size: 2rem;}
      .hero p {color: #eaf5f5; margin: 0; font-size: 1rem;}
      div[data-testid="stMetric"] {background: #f2f7f8; border: 1px solid #dbe7eb; padding: .8rem 1rem; border-radius: 10px;}
      .section-note {color: #52616b; margin-top: -.6rem; margin-bottom: 1rem;}
    </style>
    <div class="hero">
      <h1>Cytotoxicity Study Document Generator</h1>
      <p>Prepare standardized study plans and reports from structured study details and experimental data.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


if st.session_state["app_page"] == "template":
    st.header("Select a study document template")
    st.markdown(
        "Choose the supplied DCA study plan, standard study plan, or report template, or upload any Word document for a study plan."
    )
    template_choice = st.radio(
        "Template source",
        [
            "Supplied DCA study plan",
            "Supplied cytotoxicity study plan",
            "Supplied cytotoxicity study report",
            "Upload another Word template",
        ],
        key="template_choice",
    )
    uploaded_template = None
    uploaded_template_bytes = None
    uploaded_template_sha256 = ""
    template_upload_error = ""
    if template_choice == "Supplied DCA study plan":
        if SUPPLIED_DCA_TEMPLATE.exists():
            st.info("Cytotoxicity - SP - DCA - Template.docx is ready to use.")
        else:
            st.error("The supplied DCA study-plan template is missing from the project templates folder.")
        st.session_state.pop("_normalized_template_upload", None)
    elif template_choice == "Supplied cytotoxicity study plan":
        if SUPPLIED_TEMPLATE.exists():
            st.info("Cytotoxicity Study Plan Template.docx is ready to use.")
        else:
            st.error("The supplied template file is missing from the project templates folder.")
        st.session_state.pop("_normalized_template_upload", None)
    elif template_choice == "Supplied cytotoxicity study report":
        if SUPPLIED_REPORT_TEMPLATE.exists():
            st.info(
                "Cytotoxicity Study Report Template.docx is ready. This report uses one Excel workbook for study details and results."
            )
        else:
            st.error("The supplied report template file is missing from the project templates folder.")
        st.session_state.pop("_normalized_template_upload", None)
    else:
        uploaded_template = st.file_uploader(
            "Choose a Word-compatible document",
            type=["doc", "docx", "docm", "dotx", "rtf", "odt", "xml"],
            help="DOC, DOCM, DOTX, RTF, ODT, and XML files are converted to DOCX when LibreOffice is available. DOCX files preserve their original layout.",
            key="custom_template_upload",
        )
        if uploaded_template is not None:
            uploaded_content = uploaded_template.getvalue()
            uploaded_template_sha256 = hashlib.sha256(uploaded_content).hexdigest()
            upload_cache = st.session_state.get("_normalized_template_upload", {})
            cache_key = (uploaded_template.name, uploaded_template_sha256)
            if upload_cache.get("key") == cache_key:
                uploaded_template_bytes = upload_cache.get("bytes")
            else:
                try:
                    uploaded_template_bytes = normalize_template_upload(
                        uploaded_template.name,
                        uploaded_content,
                        supplied_docx=SUPPLIED_TEMPLATE.read_bytes() if SUPPLIED_TEMPLATE.exists() else None,
                    )
                    st.session_state["_normalized_template_upload"] = {
                        "key": cache_key,
                        "bytes": uploaded_template_bytes,
                    }
                except Exception as exc:
                    template_upload_error = str(exc)
            if template_upload_error:
                st.error(f"Could not load template: {template_upload_error}")
            elif uploaded_template.name.lower().endswith(".doc"):
                st.success(f"Loaded {uploaded_template.name} as an editable Word template.")
                st.caption("The original Word layout is retained; only recognized text fields and table values are updated.")
            else:
                st.success(f"Selected {uploaded_template.name}")
                st.caption("The original Word layout is retained; only recognized text fields and table values are updated.")

    ready = (
        SUPPLIED_DCA_TEMPLATE.exists()
        if template_choice == "Supplied DCA study plan"
        else SUPPLIED_TEMPLATE.exists()
        if template_choice == "Supplied cytotoxicity study plan"
        else SUPPLIED_REPORT_TEMPLATE.exists()
        if template_choice == "Supplied cytotoxicity study report"
        else uploaded_template is not None and uploaded_template_bytes is not None and not template_upload_error
    )
    template_preview_bytes = None
    if template_choice == "Supplied DCA study plan" and ready:
        template_preview_bytes = SUPPLIED_DCA_TEMPLATE.read_bytes()
    elif template_choice == "Supplied cytotoxicity study plan" and ready:
        template_preview_bytes = SUPPLIED_TEMPLATE.read_bytes()
    elif template_choice == "Supplied cytotoxicity study report" and ready:
        template_preview_bytes = SUPPLIED_REPORT_TEMPLATE.read_bytes()
    elif template_choice == "Upload another Word template" and ready:
        template_preview_bytes = uploaded_template_bytes

    if template_preview_bytes:
        template_preview_signature = hashlib.sha256(
            PREVIEW_RENDERER_VERSION.encode("utf-8")
            + b"\0"
            + template_preview_bytes
        ).hexdigest()
        with st.expander("Preview selected template"):
            st.caption(
                "This page preview is rendered from the selected Word template. The generated document uses this same page layout."
            )
            preview_cache = st.session_state.get("_template_preview_pdf", {})
            if preview_cache.get("signature") == template_preview_signature:
                st.pdf(preview_cache["bytes"], height=900)
            elif st.button("Render template preview", key=f"render_template_{template_preview_signature[:12]}"):
                try:
                    preview_pdf = render_docx_pdf(template_preview_bytes)
                    st.session_state["_template_preview_pdf"] = {
                        "signature": template_preview_signature,
                        "bytes": preview_pdf,
                    }
                    st.pdf(preview_pdf, height=900)
                except Exception as exc:
                    st.error(f"Could not render the selected template preview: {exc}")

    uploaded_template_name = uploaded_template.name if template_choice == "Upload another Word template" and uploaded_template is not None else ""
    uploaded_template_normalized = re.sub(r"[^a-z0-9]+", " ", uploaded_template_name.lower())
    uploaded_is_report = " study report " in f" {uploaded_template_normalized} " or " sr " in f" {uploaded_template_normalized} "
    continue_label = (
        "Continue to Excel upload"
        if template_choice == "Supplied cytotoxicity study report" or uploaded_is_report
        else "Continue to study details"
    )
    if st.button(continue_label, type="primary", disabled=not ready):
        if template_choice == "Supplied DCA study plan":
            template_name = SUPPLIED_DCA_TEMPLATE.name
            template_bytes = SUPPLIED_DCA_TEMPLATE.read_bytes()
            template_kind = "Study Plan"
        elif template_choice == "Supplied cytotoxicity study plan":
            template_name = SUPPLIED_TEMPLATE.name
            template_bytes = SUPPLIED_TEMPLATE.read_bytes()
            template_kind = "Study Plan"
        elif template_choice == "Supplied cytotoxicity study report":
            template_name = SUPPLIED_REPORT_TEMPLATE.name
            template_bytes = SUPPLIED_REPORT_TEMPLATE.read_bytes()
            template_kind = "Study Report"
        else:
            template_name = uploaded_template.name
            template_bytes = uploaded_template_bytes
            normalized_template_name = re.sub(r"[^a-z0-9]+", " ", template_name.lower())
            template_kind = (
                "Study Report"
                if (" study report " in f" {normalized_template_name} " or " sr " in f" {normalized_template_name} ")
                else "Study Plan"
            )
        st.session_state["selected_template"] = {
            "name": template_name,
            "bytes": template_bytes,
            "sha256": uploaded_template_sha256 or hashlib.sha256(template_bytes).hexdigest(),
            "kind": template_kind,
        }
        st.session_state["app_page"] = "study"
        st.rerun()
    st.stop()


selected_template = st.session_state.get("selected_template")
report_template_flow = bool(selected_template and selected_template.get("kind") == "Study Report")

with st.sidebar:
    st.header("Workflow")
    if report_template_flow:
        st.markdown("1. Upload one Excel workbook\n2. Generate the study report")
        st.caption("The workbook must include study details and R1*, R2*, R3*, Grade*, and Result* rows.")
    else:
        st.markdown("1. Upload study details\n2. Upload experimental data\n3. Choose document type\n4. Generate document")
        st.divider()
        st.download_button(
            "Download experimental data template",
            data=create_input_template(),
            file_name="cytotoxicity_experimental_data_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        st.download_button(
            "Download study details template",
            data=create_metadata_template(),
            file_name="cytotoxicity_study_metadata_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        st.caption("Upload a completed study details file. Experimental response values should be numeric.")


template_col, change_col = st.columns([5, 1])
with template_col:
    if selected_template:
        kind_label = selected_template.get("kind", "Study Plan").lower()
        st.info(f"Selected {kind_label} template: {selected_template['name']}")
with change_col:
    if st.button("Change template", width="stretch"):
        st.session_state["app_page"] = "template"
        st.rerun()


if report_template_flow:
    st.header("1. Upload the Excel study file")
    st.markdown(
        '<p class="section-note">Upload one workbook containing the TIDS and in-house study details, plus the R1*, R2*, R3*, Grade*, and Result* rows.</p>',
        unsafe_allow_html=True,
    )
    report_file = st.file_uploader(
        "Choose report details Excel file",
        type=["xlsx", "xlsm"],
        help="This single workbook supplies the study information and the results used in the provided Agarose Overlay report template.",
        key="report_template_excel_upload",
    )
    report_metadata: dict[str, str] = {}
    report_results: dict[str, str] = {}
    report_errors: list[str] = []
    report_warnings: list[str] = []
    if report_file is not None:
        try:
            report_metadata, report_results = read_report_workbook(
                report_file.name, report_file.getvalue()
            )
            if not report_metadata:
                report_metadata = {}
            if selected_template["name"] == SUPPLIED_REPORT_TEMPLATE.name:
                report_warnings.extend(validate_report_results(report_results))
            required_report_fields = (
                "sponsor_name",
                "sponsor_address",
                "study_number",
                "test_item_name",
            )
            if selected_template["name"] == SUPPLIED_REPORT_TEMPLATE.name:
                required_report_fields += (
                    "study_director_name",
                    "quality_assurance_unit",
                )
            report_metadata["study_director_name"] = (
                str(report_metadata.get("study_director_name", "")).strip()
                or str(report_metadata.get("study_director", "")).strip()
                or str(report_metadata.get("prepared_by", "")).strip()
            )
            report_result_labels = (
                ("replicate_1", "R1*"),
                ("replicate_2", "R2*"),
                ("replicate_3", "R3*"),
                ("grade", "Grade*"),
                ("result", "Result*"),
            )
            missing_report_results = []
            if report_errors:
                for error in report_errors:
                    st.error(error)
            elif report_warnings:
                for warning in report_warnings:
                    st.error(warning)
            else:
                st.success(
                    f"Loaded {len(report_metadata)} available study detail field(s) and available report values from {report_file.name}. Missing fields will be left blank."
                )
                with st.expander("Review imported Excel values"):
                    st.dataframe(
                        pd.DataFrame(
                            [(FIELDS.get(key, key.replace("_", " ").title()), value) for key, value in report_metadata.items()],
                            columns=["Study detail", "Imported value"],
                        ),
                        width="stretch",
                        hide_index=True,
                    )
                    st.dataframe(
                        pd.DataFrame(
                            [
                                (label, report_results.get(field, ""))
                                for field, label in (
                                    ("replicate_1", "R1*"),
                                    ("replicate_2", "R2*"),
                                    ("replicate_3", "R3*"),
                                    ("grade", "Grade*"),
                                    ("result", "Result*"),
                                )
                            ],
                            columns=["Report result", "Imported value"],
                        ),
                        width="stretch",
                        hide_index=True,
                    )
                with st.expander("Live filled document preview", expanded=True):
                    try:
                        live_report_docx = build_study_document(
                            kind="Study Report",
                            metadata=report_metadata,
                            source_filename=report_file.name,
                            template_bytes=selected_template["bytes"],
                            report_results=report_results,
                        )
                        live_report_pdf = render_docx_pdf(live_report_docx)
                        st.pdf(live_report_pdf, height=1050)
                    except Exception as exc:
                        st.error(f"Could not render the live filled preview: {exc}")
        except Exception as exc:
            report_errors.append(f"Could not read the Excel file: {exc}")
            st.error(report_errors[-1])

    st.header("2. Generate document")
    report_signature = hashlib.sha256(
        json.dumps(
            {
                "template": selected_template["sha256"],
                "upload": (
                    report_file.name + ":" + hashlib.sha256(report_file.getvalue()).hexdigest()
                    if report_file is not None
                    else ""
                ),
                "metadata": report_metadata,
                "results": report_results,
                "document_generator_version": DOCUMENT_GENERATOR_VERSION,
                "preview_renderer_version": PREVIEW_RENDERER_VERSION,
            },
            sort_keys=True,
            ensure_ascii=False,
        ).encode("utf-8")
    ).hexdigest()
    report_can_generate = bool(
        report_file is not None
        and not report_errors
        and not report_warnings
    )
    if st.button("Generate Word document", type="primary", disabled=not report_can_generate, key="generate_report_document"):
        try:
            report_docx = build_study_document(
                kind="Study Report",
                metadata=report_metadata,
                source_filename=report_file.name if report_file is not None else None,
                template_bytes=selected_template["bytes"],
                report_results=report_results,
            )
            safe_number = "".join(
                character
                for character in report_metadata.get("study_number", "")
                if character.isalnum() or character in "-_ "
            ).strip()
            suffix = f"_{safe_number}" if safe_number else ""
            st.session_state["generated_document"] = {
                "bytes": report_docx,
                "filename": f"study_report{suffix}.docx",
                "kind": "Study Report",
                "signature": report_signature,
            }
            st.success("Study Report generated from your Excel file. Review the preview, then download the Word document below.")
        except Exception as exc:
            st.error(f"Document generation failed: {exc}")

    generated = st.session_state.get("generated_document")
    if generated and generated.get("kind") == "Study Report":
        if generated.get("signature") == report_signature:
            st.subheader("Document preview")
            st.caption("Preview rendered from the supplied Study Report Word template.")
            try:
                preview_pdf = render_docx_pdf(generated["bytes"])
                st.pdf(preview_pdf, height=1050)
            except Exception as exc:
                st.error(f"Could not render the Word document preview: {exc}")
            st.download_button(
                "Download Study Report (.docx)",
                data=generated["bytes"],
                file_name=generated["filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
        else:
            st.info("The Excel file, template, or document generation rules changed. Generate a new document to include the latest version.")
    st.stop()


st.header("1. Upload study details")
st.markdown(
    '<p class="section-note">Upload your Study Details Excel. You may also upload the Medical Device TIDS Word file to fill any missing details. Excel values take priority when both files provide the same field.</p>',
    unsafe_allow_html=True,
)
metadata_files = st.file_uploader(
    "Choose study details files",
    type=["xlsx", "xlsm", "docx", "json"],
    accept_multiple_files=True,
    help="The Study Details Excel can be uploaded by itself. Optionally include the Medical Device TIDS Word file; Excel values take priority and TIDS values fill gaps.",
    key="metadata_upload",
)
metadata_sources = []
metadata_upload_parts = []
for metadata_file in metadata_files or []:
    metadata_bytes = metadata_file.getvalue()
    metadata_sources.append((metadata_file, metadata_bytes))
    metadata_upload_parts.append(
        (
            metadata_file.name,
            hashlib.sha256(metadata_bytes).hexdigest(),
            METADATA_PARSER_VERSION,
        )
    )
current_metadata_upload = tuple(metadata_upload_parts) if metadata_upload_parts else None
if metadata_sources:
    st.session_state.pop("_metadata_fallback_signature", None)
    st.session_state.pop("_metadata_fallback_source_name", None)
    if st.session_state.get("_metadata_signature") != current_metadata_upload:
        for field in FIELDS:
            st.session_state.pop(field, None)
        st.session_state.pop("_metadata_workbook_signature", None)
        st.session_state.pop("_metadata_workbook_source_name", None)
        try:
            # Load device-specific fields from the TIDS first, then let the
            # study-details workbook take precedence for overlapping values.
            ordered_sources = sorted(
                metadata_sources,
                key=lambda source: Path(source[0].name).suffix.lower() in {".xlsx", ".xlsm"},
            )
            imported = {}
            for source_file, source_bytes in ordered_sources:
                source_values = read_metadata_file(source_file.name, source_bytes)
                if not source_values:
                    raise ValueError(
                        f"No recognized study fields were found in {source_file.name}. Check its field labels."
                    )
                imported.update(source_values)
            if not imported:
                raise ValueError("No recognized study fields were found. Use the study details template or check the field labels.")
            st.session_state.update(imported)
            st.session_state["_metadata_base_values"] = dict(imported)
            st.session_state["_metadata_signature"] = current_metadata_upload
            st.session_state["_metadata_import_count"] = len(imported)
            st.session_state["_metadata_import_error"] = ""
        except Exception as exc:
            st.session_state["_metadata_base_values"] = {}
            st.session_state["_metadata_signature"] = current_metadata_upload
            st.session_state["_metadata_import_count"] = 0
            st.session_state["_metadata_import_error"] = str(exc)

metadata_values = {field: st.session_state.get(field, "") for field in FIELDS}
fallback_metadata_signature = tuple(st.session_state.get("_metadata_fallback_signature", ()))
metadata_signature = (
    json.dumps(current_metadata_upload, ensure_ascii=False)
    if current_metadata_upload
    else ":".join(fallback_metadata_signature)
)
current_metadata_is_loaded = (
    current_metadata_upload is not None
    and st.session_state.get("_metadata_signature") == current_metadata_upload
)
fallback_metadata_is_loaded = (
    current_metadata_upload is None and bool(fallback_metadata_signature)
)
metadata_ready = bool(
    st.session_state.get("_metadata_import_count", 0) > 0
    and not st.session_state.get("_metadata_import_error")
    and (current_metadata_is_loaded or fallback_metadata_is_loaded)
)
if metadata_ready:
    if metadata_sources:
        source_names = ", ".join(source_file.name for source_file, _ in metadata_sources)
        workbook_source_name = st.session_state.get("_metadata_workbook_source_name")
        if workbook_source_name:
            source_names += f", {workbook_source_name}"
        st.success(f"Loaded {st.session_state['_metadata_import_count']} study field(s) from {source_names}.")
    else:
        source_name = st.session_state.get("_metadata_fallback_source_name", "the experimental workbook")
        st.success(f"Recognized {source_name} as study details and loaded {st.session_state['_metadata_import_count']} field(s).")
elif metadata_sources:
    st.error(f"Could not load study details: {st.session_state.get('_metadata_import_error', 'Please reselect the file.')}")
else:
    st.info("Upload a study details file to continue.")


st.header("2. Experimental data")
st.caption(
    "Optional for a Study Plan; required for a Study Report. If you upload a Study Details workbook here, its fields are merged with the details above; it is not treated as assay data."
)
data: pd.DataFrame | None = None
data_file = None
errors: list[str] = []
warnings: list[str] = []
data_file = st.file_uploader(
    "Upload experimental data workbook (.xlsx)",
    type=["xlsx", "xlsm"],
    help="Required columns: Treatment, Concentration, Replicate, Response. Unit and Response Unit are optional.",
    key="experimental_upload",
)
if data_file is not None:
    try:
        experimental_bytes = data_file.getvalue()
        experimental_signature = (
            data_file.name,
            hashlib.sha256(experimental_bytes).hexdigest(),
            METADATA_PARSER_VERSION,
        )
        excel = pd.ExcelFile(BytesIO(experimental_bytes))
        assay_sheets = []
        required_assay_columns = {
            "Treatment",
            "Concentration",
            "Replicate",
            "Response",
        }
        for candidate in excel.sheet_names:
            headers = pd.read_excel(excel, sheet_name=candidate, nrows=0)
            normalized_headers, _, _ = validate_experimental_data(headers)
            if required_assay_columns.issubset(normalized_headers.columns):
                assay_sheets.append(candidate)

        if not assay_sheets:
            try:
                imported = read_metadata_file(data_file.name, experimental_bytes)
            except Exception:
                imported = {}

            if imported:
                merge_signature = (experimental_signature, current_metadata_upload)
                if st.session_state.get("_metadata_workbook_signature") != merge_signature:
                    base_values = (
                        dict(st.session_state.get("_metadata_base_values", {}))
                        if current_metadata_upload is not None
                        else {}
                    )
                    merged_values = {**base_values, **imported}
                    for field in FIELDS:
                        st.session_state.pop(field, None)
                    st.session_state.update(merged_values)
                    st.session_state["_metadata_base_values"] = (
                        base_values if current_metadata_upload is not None else dict(imported)
                    )
                    st.session_state["_metadata_workbook_signature"] = merge_signature
                    st.session_state["_metadata_workbook_source_name"] = data_file.name
                    if current_metadata_upload is None:
                        st.session_state.pop("_metadata_signature", None)
                        st.session_state["_metadata_fallback_signature"] = experimental_signature
                        st.session_state["_metadata_fallback_source_name"] = data_file.name
                    st.session_state["_metadata_import_count"] = len(merged_values)
                    st.session_state["_metadata_import_error"] = ""
                    st.rerun()

                st.info(
                    "This workbook contains study details, not assay measurements. "
                    "Its fields have been combined with the Study Details/TIDS upload above. "
                    "A Study Report needs a separate assay workbook with Treatment, Concentration, Replicate, and Response columns."
                )
            else:
                source = pd.read_excel(excel, sheet_name=excel.sheet_names[0])
                _, errors, _ = validate_experimental_data(source)
                for error in errors:
                    st.error(error)
        else:
            available_sheets = assay_sheets
            selected_sheet = st.session_state.get("experimental_sheet")
            if selected_sheet not in available_sheets:
                selected_sheet = available_sheets[0]
                st.session_state["experimental_sheet"] = selected_sheet
            if len(available_sheets) > 1:
                sheet_name = st.selectbox(
                    "Worksheet", available_sheets, key="experimental_sheet"
                )
            else:
                sheet_name = available_sheets[0]
            source = pd.read_excel(excel, sheet_name=sheet_name)
            data, errors, warnings = validate_experimental_data(source)
            if errors:
                for error in errors:
                    st.error(error)
            else:
                st.success(f"Loaded {len(data):,} experimental row(s) from “{sheet_name}”.")
                columns = st.columns(3)
                columns[0].metric("Data rows", f"{len(data):,}")
                columns[1].metric("Treatments", str(data["Treatment"].nunique()))
                columns[2].metric("Concentrations", str(data["Concentration"].nunique()))
                st.dataframe(data, width="stretch", hide_index=True)
                for warning in warnings:
                    st.warning(warning)
    except Exception as exc:
        st.error(f"Could not read experimental workbook: {exc}")


st.header("3. Document type")
document_kind = st.radio("Choose what to generate", ["Study Plan", "Study Report"], horizontal=True)
if document_kind == "Study Plan" and data_file is None:
    st.info("A study plan does not require assay results. You can leave the experimental data upload empty.")

st.header("4. Generate document")
metadata_warnings: list[str] = []
required_fields = ("sponsor_name", "test_item_name")
optional_fields = ("study_title", "study_number")
if document_kind == "Study Plan":
    required_fields += ("study_number",)
    optional_fields = ("study_title",)
missing_required_fields = []
for field in required_fields:
    if not str(metadata_values.get(field, "")).strip():
        missing_required_fields.append(field)
        metadata_warnings.append(f"Required field missing: {FIELDS[field]}.")
missing_optional_fields = [
    field for field in optional_fields if not str(metadata_values.get(field, "")).strip()
]
parsed_dates: dict[str, date] = {}
for field in ("study_start", "study_end"):
    raw_date = str(metadata_values.get(field, "")).strip()
    if raw_date:
        try:
            parsed_dates[field] = date.fromisoformat(raw_date)
        except ValueError:
            metadata_warnings.append(f"{FIELDS[field]} should use YYYY-MM-DD format.")
if "study_start" in parsed_dates and "study_end" in parsed_dates and parsed_dates["study_end"] < parsed_dates["study_start"]:
    metadata_warnings.append("Study End Date is earlier than Study Start Date.")
if metadata_ready:
    for warning in metadata_warnings:
        st.warning(warning)
    if missing_required_fields:
        if document_kind == "Study Plan" and "study_number" in missing_required_fields:
            st.info("A Study Plan needs a study number. Upload the Study Details Excel with the Medical Device TIDS so the number can be filled from your files.")
        else:
            st.info("Add the missing required fields to the uploaded study details file, then upload it again.")
    if missing_optional_fields:
        labels = ", ".join(FIELDS[field] for field in missing_optional_fields)
        missing_note = (
            "The template's default title wording will be used."
            if document_kind == "Study Plan"
            else "The generated document will mark these as not provided."
        )
        st.caption(f"Optional fields not supplied: {labels}. {missing_note}")
if document_kind == "Study Report" and data_file is None:
    st.info("Upload experimental data to generate a study report.")
if document_kind == "Study Report" and data is not None and errors:
    st.info("Resolve the data validation errors above before generating a study report.")

required_metadata_present = not missing_required_fields
can_generate = metadata_ready and required_metadata_present and (
    document_kind == "Study Plan" or (data is not None and not errors and not data.empty)
)
source_signature = (
    data_file.name + ":" + hashlib.sha256(data_file.getvalue()).hexdigest()
    if data_file is not None
    else ""
)
current_signature = hashlib.sha256(
    json.dumps(
        {
            "kind": document_kind,
            "document_generator_version": DOCUMENT_GENERATOR_VERSION,
            "preview_renderer_version": PREVIEW_RENDERER_VERSION,
            "metadata": metadata_values,
            "metadata_source": metadata_signature,
            "source": source_signature,
            "sheet": st.session_state.get("experimental_sheet", ""),
            "template": (
                st.session_state.get("selected_template", {}).get("sha256", "")
            ),
        },
        sort_keys=True,
        ensure_ascii=False,
    ).encode("utf-8")
).hexdigest()
if st.button("Generate Word document", type="primary", disabled=not can_generate, key="generate_standard_document"):
    metadata = metadata_values
    try:
        result = build_study_document(
            kind=document_kind,
            metadata=metadata,
            data=data if document_kind == "Study Report" else None,
            source_filename=data_file.name if data_file is not None else None,
            template_bytes=(
                st.session_state.get("selected_template", {}).get("bytes")
                if document_kind == "Study Plan"
                else None
            ),
        )
        safe_number = "".join(character for character in metadata.get("study_number", "") if character.isalnum() or character in "-_ ").strip()
        suffix = f"_{safe_number}" if safe_number else ""
        st.session_state["generated_document"] = {
            "bytes": result,
            "filename": f"{document_kind.lower().replace(' ', '_')}{suffix}.docx",
            "kind": document_kind,
            "signature": current_signature,
        }
        st.success(f"{document_kind} generated. Download the editable Microsoft Word document below.")
    except Exception as exc:
        st.error(f"Document generation failed: {exc}")

generated = st.session_state.get("generated_document")
if generated:
    if generated["signature"] == current_signature:
        st.subheader("Document preview")
        template_name = st.session_state.get("selected_template", {}).get("name", "selected template")
        st.caption(
            f"Page preview rendered from the filled Word document using the {template_name} layout. Download the editable .docx below."
        )
        try:
            preview_pdf = render_docx_pdf(generated["bytes"])
            st.pdf(preview_pdf, height=1050)
        except Exception as exc:
            st.error(f"Could not render the Word document preview: {exc}")
        st.download_button(
            f"Download {generated['kind']} (.docx)",
            data=generated["bytes"],
            file_name=generated["filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )
    else:
        st.info("Study details, source data, template, or document generation rules changed. Generate a new document to include the latest version.")

# The standard document uses the generate/preview block above; duplicate widgets are intentionally not rendered here.
