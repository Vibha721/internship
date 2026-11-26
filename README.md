# Cytotoxicity Study Document Generator

A local Streamlit prototype for turning structured sponsor, test item, and study information into standardized Microsoft Word study plans and reports. Study reports include a validated experimental-data summary table, a concentration-response graph, and the imported observations.

## Technology

- Python and Streamlit for the interface
- pandas and openpyxl for Excel import, templates, and validation
- python-docx for editable Word output
- matplotlib for study-report charts

The supplied DCA study plan, standard cytotoxicity study plan, and study report are included in `templates/` and are available from the app's opening template-selection page. Study Plans and the supplied Study Report fill the selected Word template while preserving its layout, tables, page setup, fields, and embedded images. You can also upload a compatible `.doc` or `.docx` plan template; legacy `.doc` files are converted to editable `.docx` first.

## Run locally

Use Python 3.10 or newer.

The requirements install Streamlit's PDF viewer for the in-app page-accurate preview. Install LibreOffice for rendering the preview and for converting custom legacy `.doc` templates. Without LibreOffice, the editable Word file can still be generated, but the app cannot render the preview.

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run app.py
```

Open the local URL printed by Streamlit. The app runs locally; uploaded study inputs are processed in the current Streamlit session and are not written to a project database.

## Workflow

1. Select the supplied DCA study plan, standard study plan, supplied study report, or upload a compatible Word plan template. Other legacy `.doc` files require LibreOffice for conversion. The chosen template is retained for the current session; use **Change template** to return to the selection page.
2. For a Study Plan, upload one or more study-details files, then optionally upload experimental data. The Study Details Excel and Medical Device TIDS Word file can be uploaded together; Excel values take priority when the same field appears in both.
3. For the supplied Study Report template, upload one Excel workbook containing study details and the `R1*`, `R2*`, `R3*`, `Grade*`, and `Result*` rows. This template's fixed narrative is for three Grade 1 (Slight) replicates and a non-cytotoxic conclusion; mismatching results are blocked so the report cannot misstate the uploaded findings.
4. Generate the document, inspect the page-accurate Word preview, then download the editable `.docx` file.

Metadata files can be `.xlsx`, `.xlsm`, `.docx`, or `.json`. Excel metadata can use a two-column key/value sheet or a study-details table with `Description` and `Details` columns. Word metadata can use two-column tables or paragraphs such as `Sponsor Name: Example Sponsor`. The Medical Device TIDS format is recognized, including sponsor, test item, intended use, lot number, study number, and sponsor representative. JSON should contain field names and values. The app reads recognized labels directly from uploaded files, combines their fields, and ignores unknown labels. When a field appears in multiple files, Excel values take priority. Sponsor and test item are required for all documents; study number is required for a Study Plan. Study title remains optional.

If a study-details workbook is uploaded in the experimental-data section, the app recognizes it and merges its fields with the study-details/TIDS files rather than treating it as assay measurements.

## Experimental workbook format

Required columns:

| Column | Meaning |
| --- | --- |
| `Treatment` | Control or test treatment/group name |
| `Concentration` | Numeric concentration; zero is allowed for controls |
| `Replicate` | Replicate identifier, such as `1` or `R1` |
| `Response` | Numeric measured response |

`Unit` and `Response Unit` are optional. Common header aliases such as `Dose`, `Group`, `Rep`, and `Viability (%)` are recognized. Use one concentration unit and one response unit throughout a workbook. Negative concentrations, missing required values, mixed units, and nonnumeric concentration/response values block report generation. Response values outside 0–100 are flagged for review, but not blocked because other response scales may be used.

Study reports summarize the mean and standard deviation by treatment and concentration. The chart shows mean response with standard-deviation error bars. A single replicate has no standard deviation. The app does not fit an IC50 model or apply assay-specific acceptance rules; enter protocol-specific criteria and conclusions in the study details.

## Project layout

```text
app.py
cytotox/
  excel_io.py
  metadata.py
  preview.py
  report_excel.py
  report_generator.py
  report_template_filler.py
  template_filler.py
  template_input.py
requirements.txt
templates/
  Cytotoxicity - SP - DCA - Template.docx
  Cytotoxicity Study Plan Template.doc
  Cytotoxicity Study Plan Template.docx
  Cytotoxicity Study Report Template.docx
```

The legacy `.doc` is retained as the original reference; the `.docx` copy is used to create editable Study Plans. Uploaded replacement templates must retain the supplied study-plan table structure and highlighted fields. Review generated study documents before controlled use.
