"""Fill the supplied Cytotoxicity Study Plan without rebuilding its layout.

The Word package is patched at the XML-part level so its styles, fields,
content controls, relationships, images, page setup, and fixed protocol text
remain those of the selected template.
"""

from __future__ import annotations

import hashlib
import re
from io import BytesIO
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from lxml import etree

from .metadata import FIELDS


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"
SUPPLIED_TEMPLATE_SHA256 = "0debaae0e66c276267caf527751265f610b1ca566a78a302c58396ecd380735f"
DCA_TEMPLATE_SHA256 = "7f4bdfd7697f4d9d6506a79be39a52cf9135b29fbe3f6e061e066c4e816994cf"


def _set_dca_toc_run_properties(container: etree._Element, *, caps: bool) -> None:
    """Apply the uploaded DCA template's blue, regular Times New Roman TOC text."""

    run_properties = container.find(f"{W}rPr")
    if run_properties is None:
        run_properties = etree.Element(f"{W}rPr")
        if container.tag == f"{W}pPr":
            container.append(run_properties)
        else:
            container.insert(0, run_properties)

    preserved = {
        child.tag: child
        for child in run_properties
        if child.tag in {f"{W}rStyle", f"{W}lang"}
    }
    for child in list(run_properties):
        run_properties.remove(child)

    if f"{W}rStyle" in preserved:
        run_properties.append(preserved[f"{W}rStyle"])

    fonts = etree.SubElement(run_properties, f"{W}rFonts")
    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(f"{W}{attribute}", "Times New Roman")

    for tag in ("b", "bCs"):
        element = etree.SubElement(run_properties, f"{W}{tag}")
        element.set(f"{W}val", "0")

    caps_element = etree.SubElement(run_properties, f"{W}caps")
    caps_element.set(f"{W}val", "1" if caps else "0")
    color = etree.SubElement(run_properties, f"{W}color")
    color.set(f"{W}val", "0000FF")
    small_caps = etree.SubElement(run_properties, f"{W}smallCaps")
    small_caps.set(f"{W}val", "0")
    for tag in ("sz", "szCs"):
        size = etree.SubElement(run_properties, f"{W}{tag}")
        size.set(f"{W}val", "24")

    if f"{W}lang" in preserved:
        run_properties.append(preserved[f"{W}lang"])


def _format_dca_toc(
    document_root: etree._Element, styles_root: etree._Element | None = None
) -> int:
    """Format cached DCA table-of-contents entries like the uploaded file."""

    instruction = " ".join(
        document_root.xpath(".//w:instrText/text()", namespaces=NS)
    )
    if "Table,6,Figure,7,Appendix,8,Annexure,9" not in instruction:
        return 0

    contents = document_root.xpath(
        ".//w:sdtContent[.//w:instrText[contains(., 'TOC')]]",
        namespaces=NS,
    )
    if not contents:
        return 0

    left_position = "720"
    right_position = "8723"
    if styles_root is not None:
        toc_style = styles_root.xpath(
            ".//w:style[@w:styleId='TOC1']", namespaces=NS
        )
        if toc_style:
            style_properties = toc_style[0].find(f"{W}pPr")
            if style_properties is not None:
                indent = style_properties.find(f"{W}ind")
                indent_position = (
                    (indent.get(f"{W}start") or indent.get(f"{W}left"))
                    if indent is not None
                    else None
                )
                tabs = style_properties.findall(f"{W}tabs/{W}tab")
                left_positions = {
                    tab.get(f"{W}pos")
                    for tab in tabs
                    if tab.get(f"{W}val") == "clear"
                    and tab.get(f"{W}pos") == indent_position
                }
                right_tabs = [
                    tab
                    for tab in tabs
                    if tab.get(f"{W}val") == "right"
                    and tab.get(f"{W}leader") == "dot"
                    and tab.get(f"{W}pos", "").isdigit()
                ]
                if left_positions:
                    left_position = next(iter(left_positions))
                if right_tabs:
                    right_position = max(
                        (tab.get(f"{W}pos") for tab in right_tabs), key=int
                    )

    formatted = 0
    for paragraph in contents[0].xpath(".//w:p", namespaces=NS):
        style = paragraph.xpath(
            "string(./w:pPr/w:pStyle/@w:val)", namespaces=NS
        )
        if not style.startswith("TOC") or style == "TOCHeading":
            continue

        caps = style == "TOC1"
        properties = paragraph.find(f"{W}pPr")
        if properties is None:
            properties = etree.Element(f"{W}pPr")
            paragraph.insert(0, properties)
        _set_dca_toc_run_properties(properties, caps=caps)

        for run in paragraph.xpath(".//w:r", namespaces=NS):
            _set_dca_toc_run_properties(run, caps=caps)

        if style == "TOC1":
            tabs = properties.find(f"{W}tabs")
            if tabs is not None:
                properties.remove(tabs)
            tabs = etree.Element(f"{W}tabs")
            left = etree.SubElement(tabs, f"{W}tab")
            left.set(f"{W}val", "left")
            left.set(f"{W}pos", left_position)
            right = etree.SubElement(tabs, f"{W}tab")
            right.set(f"{W}val", "right")
            right.set(f"{W}pos", right_position)
            right.set(f"{W}leader", "dot")
            style_element = properties.find(f"{W}pStyle")
            tabs_index = (
                list(properties).index(style_element) + 1
                if style_element is not None
                else 0
            )
            properties.insert(tabs_index, tabs)

            alignment = properties.find(f"{W}jc")
            if alignment is None:
                alignment = etree.Element(f"{W}jc")
                run_properties = properties.find(f"{W}rPr")
                alignment_index = (
                    list(properties).index(run_properties)
                    if run_properties is not None
                    else len(properties)
                )
                properties.insert(alignment_index, alignment)
            alignment.set(f"{W}val", "start")
        formatted += 1
    return formatted


def apply_dca_toc_formatting(document_bytes: bytes) -> bytes:
    """Return the DCA Word package with its TOC restored to the source styling."""

    try:
        with ZipFile(BytesIO(document_bytes), "r") as source_zip:
            names = set(source_zip.namelist())
            if "word/document.xml" not in names:
                return document_bytes
            document = etree.fromstring(source_zip.read("word/document.xml"))
            styles = (
                etree.fromstring(source_zip.read("word/styles.xml"))
                if "word/styles.xml" in names
                else None
            )
            if not _format_dca_toc(document, styles):
                return document_bytes

            output = BytesIO()
            with ZipFile(output, "w", ZIP_DEFLATED) as destination:
                for info in source_zip.infolist():
                    data = source_zip.read(info.filename)
                    if info.filename == "word/document.xml":
                        data = etree.tostring(
                            document,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone=True,
                        )
                    destination.writestr(info, data)
            return output.getvalue()
    except Exception as exc:
        raise ValueError("Could not restore the DCA table-of-contents formatting.") from exc


def _plain(value: Any) -> str:
    return str(value).strip() if value is not None and str(value).strip() else ""


def _inline(value: Any) -> str:
    return re.sub(r"\s+", " ", _plain(value)).strip()


def _format_like(value: str, original: str) -> str:
    if original.isupper():
        return value.upper()
    if original.islower():
        return value.lower()
    if original.istitle():
        return value.title()
    return value


def _text_nodes(root: etree._Element) -> list[etree._Element]:
    return root.xpath(".//w:t", namespaces=NS)


def _replace_text(root: etree._Element, old: str, new: str, *, preserve_case: bool = True) -> int:
    if not old:
        return 0
    count = 0
    pattern = re.compile(re.escape(old), re.IGNORECASE)
    for node in _text_nodes(root):
        current = node.text or ""
        if pattern.search(current):
            node.text = pattern.sub(
                lambda match: _format_like(new, match.group(0)) if preserve_case else new,
                current,
            )
            if node.text[:1].isspace() or node.text[-1:].isspace():
                node.set(f"{{{XML_NS}}}space", "preserve")
            count += 1
    return count


def _run_text(run: etree._Element) -> str:
    return "".join(node.text or "" for node in run.xpath(".//w:t", namespaces=NS))


def _set_run_text(run: etree._Element, value: str) -> None:
    for child in list(run):
        if child.tag != f"{W}rPr":
            run.remove(child)
    text_node = etree.SubElement(run, f"{W}t")
    text_node.text = value
    if value[:1].isspace() or value[-1:].isspace():
        text_node.set(f"{{{XML_NS}}}space", "preserve")


def _replace_in_paragraph(paragraph: etree._Element, old: str, new: str, *, preserve_case: bool = True) -> bool:
    for run in paragraph.xpath("./w:r", namespaces=NS):
        current = _run_text(run)
        if current == old:
            _set_run_text(run, _format_like(new, old) if preserve_case else new)
            return True
    return False


def _set_paragraph_text(
    paragraph: etree._Element, value: str, *, prefer_highlight: bool = True
) -> None:
    runs = paragraph.xpath("./w:r", namespaces=NS)
    if not runs:
        run = etree.SubElement(paragraph, f"{W}r")
        runs = [run]
    highlighted = [
        run
        for run in runs
        if run.xpath("./w:rPr/w:highlight", namespaces=NS)
    ]
    non_highlighted = [run for run in runs if run not in highlighted]
    preferred = highlighted if prefer_highlight else non_highlighted
    target = preferred[0] if preferred else next(
        (run for run in runs if run.xpath(".//w:t", namespaces=NS)), runs[0]
    )
    _set_run_text(target, value)
    for run in runs:
        if run is not target and run.xpath(".//w:t", namespaces=NS):
            _set_run_text(run, "")


def _rows(table: etree._Element) -> list[etree._Element]:
    return table.findall(f"{W}tr")


def _cells(row: etree._Element) -> list[etree._Element]:
    return row.findall(f"{W}tc")


def _cell_width_twips(cell: etree._Element) -> int | None:
    width = cell.find(f"{W}tcPr/{W}tcW")
    if width is None or width.get(f"{W}type", "dxa") != "dxa":
        return None
    try:
        return int(width.get(f"{W}w", ""))
    except ValueError:
        return None


def _value_cell_is_wide_enough(
    label_cell: etree._Element, value_cell: etree._Element, value: str
) -> bool:
    value_width = _cell_width_twips(value_cell)
    if value_width is None:
        return True

    # Do not treat narrow spacer columns as the value column in a complex
    # template. Estimate the unbreakable width so longer words cannot stack
    # one character per line in the uploaded layout.
    longest_word = max((len(word) for word in re.findall(r"\S+", value)), default=0)
    required_width = 240 + longest_word * 80
    label_width = _cell_width_twips(label_cell)
    if label_width is not None:
        required_width = max(required_width, label_width // 3)
    return value_width >= required_width


def _paragraphs(cell: etree._Element) -> list[etree._Element]:
    return cell.findall(f"{W}p")


def _cell_at(tables: list[etree._Element], table_index: int, row_index: int, cell_index: int) -> etree._Element:
    try:
        return _cells(_rows(tables[table_index])[row_index])[cell_index]
    except IndexError as exc:
        raise ValueError(
            "This Word template does not have the Cytotoxicity Study Plan table structure. "
            "Choose the supplied template or a compatible copy."
        ) from exc


def _set_cell_paragraph(
    tables: list[etree._Element],
    table_index: int,
    row_index: int,
    cell_index: int,
    paragraph_index: int,
    value: str,
    *,
    prefer_highlight: bool = True,
) -> None:
    cell = _cell_at(tables, table_index, row_index, cell_index)
    paragraphs = _paragraphs(cell)
    if paragraph_index < len(paragraphs):
        _set_paragraph_text(paragraphs[paragraph_index], value, prefer_highlight=prefer_highlight)


def _replace_in_cell(
    tables: list[etree._Element], table_index: int, row_index: int, cell_index: int, old: str, new: str
) -> bool:
    cell = _cell_at(tables, table_index, row_index, cell_index)
    for paragraph in _paragraphs(cell):
        if _replace_in_paragraph(paragraph, old, new):
            return True
    return False


def _replace_cell_text_exact(
    tables: list[etree._Element], table_index: int, row_index: int, cell_index: int, old: str, new: str
) -> bool:
    cell = _cell_at(tables, table_index, row_index, cell_index)
    for paragraph in _paragraphs(cell):
        if _replace_in_paragraph(paragraph, old, new):
            return True
    return False


def _find_body_paragraph(root: etree._Element, begins_with: str) -> etree._Element | None:
    for paragraph in root.xpath("./w:body/w:p", namespaces=NS):
        text = "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
        if text.startswith(begins_with):
            return paragraph
    return None


def _normalise_ratio(value: str) -> str:
    value = value.strip()
    return re.sub(r"(?<=\d)(?=[A-Za-zµμ])", " ", value)


def fill_dca_study_plan_template(template_bytes: bytes, metadata: dict[str, Any]) -> bytes:
    """Fill the DCA plan's labeled slots while preserving its Word package layout."""

    sponsor_name = _inline(metadata.get("sponsor_name"))
    test_item = _inline(metadata.get("test_item_name"))
    study_number = _inline(metadata.get("study_number"))
    if not sponsor_name or not test_item or not study_number:
        raise ValueError("Sponsor Name, Test Item Name, and Study Number are required by the DCA study-plan template.")

    study_director = _inline(
        metadata.get("study_director_name")
        or metadata.get("study_director")
        or metadata.get("prepared_by")
    )
    manufacturer = ", ".join(
        value
        for value in (
            _inline(metadata.get("manufacturer_name")),
            _inline(metadata.get("manufacturer_address")),
        )
        if value
    )
    replacements = {
        "Study Sponsor*": sponsor_name,
        "STUDY SPONSOR*": sponsor_name,
        "Address*": _inline(metadata.get("sponsor_address")),
        "Study Number*": study_number,
        "STUDY NUMBER*": study_number,
        "SPONSOR REPRESENTATIVE*": _inline(metadata.get("sponsor_contact")),
        "MONITORING SCIENTIST*": _inline(metadata.get("monitoring_scientist")),
        "Study Director*": study_director,
        "STUDY DIRECTOR NAME*": study_director,
        "Study Director Email*": _inline(metadata.get("study_director_email")),
        "Study Co-ordinator*": _inline(metadata.get("study_coordinator")),
        "Study Personnel*": _inline(metadata.get("study_personnel")),
        "Study Initiation Date*": _inline(metadata.get("study_initiation_date")),
        "Experimental Starting Date*": _inline(metadata.get("experimental_start_date")),
        "Experimental Completion Date*": _inline(metadata.get("experimental_completion_date")),
        "Draft Report Submission Date*": _inline(metadata.get("draft_report_submission_date")),
        "Batch No*": _inline(metadata.get("batch_number") or metadata.get("batch_lot")),
        "Manufactured/Supplied*": manufacturer,
        "Date of Manufacture*": _inline(metadata.get("manufacture_date")),
        "Date of Expiry*": _inline(metadata.get("expiry_date")),
        "Quantity Sent*": _inline(metadata.get("quantity_sent")),
        "Storage Condition*": _inline(metadata.get("storage_condition")),
        "Test Item Code by Test Facility*": _inline(
            metadata.get("test_item_code") or metadata.get("test_item_id")
        ),
        "TEST FACILITY MAN*": _inline(metadata.get("test_facility_manager")),
        "QUALITY ASSURANCE UNIT*": _inline(metadata.get("quality_assurance_unit")),
    }

    study_title = _inline(metadata.get("study_title"))
    title_replacements = (
        "In vitro Cytotoxicity Study of TEST ITEM* by Direct Contact Method",
        "In vitro Cytotoxicity study of TEST ITEM* by Direct Contact Method",
    )

    output = BytesIO()
    try:
        with ZipFile(BytesIO(template_bytes), "r") as source_zip:
            names = source_zip.namelist()
            if "word/document.xml" not in names:
                raise ValueError("The DCA study-plan template is not a valid editable DOCX.")
            parts: dict[str, bytes] = {}
            for part_name in names:
                if not part_name.startswith("word/") or not part_name.endswith(".xml"):
                    continue
                root = etree.fromstring(source_zip.read(part_name))
                for paragraph in root.xpath(".//w:p", namespaces=NS):
                    if study_title:
                        for original_title in title_replacements:
                            _replace_across_paragraph_prefer_plain_run(
                                paragraph, original_title, study_title
                            )
                    _replace_test_item_placeholders(paragraph, test_item)
                    for marker, value in replacements.items():
                        _replace_across_paragraph(paragraph, {marker: value})
                    for marker, value in _starred_placeholder_replacements(metadata).items():
                        _replace_across_paragraph(paragraph, {marker: value})

                # Physical appearance is the only editable value represented by
                # an empty table cell rather than a visible placeholder.
                if part_name == "word/document.xml":
                    tables = root.xpath("./w:body/w:tbl", namespaces=NS)
                    appearance = _inline(metadata.get("physical_appearance"))
                    if appearance and len(tables) > 2:
                        _set_cell_paragraph(tables, 2, 1, 2, 0, appearance)
                    study_code = _inline(metadata.get("study_code"))
                    if study_code and len(tables) > 1:
                        _set_cell_paragraph(tables, 1, 2, 3, 0, study_code)

                parts[part_name] = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )

            with ZipFile(output, "w", compression=ZIP_DEFLATED) as destination:
                for info in source_zip.infolist():
                    destination.writestr(
                        info,
                        parts.get(info.filename, source_zip.read(info.filename)),
                    )
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("Could not fill the DCA study-plan template.") from exc
    return apply_dca_toc_formatting(output.getvalue())


_GENERIC_ALIASES = {
    re.sub(r"[^a-z0-9]+", "", label.lower()): field
    for field, label in FIELDS.items()
}
_GENERIC_ALIASES.update({
    "sponsor": "sponsor_name",
    "sponsorcompany": "sponsor_name",
    "studyid": "study_number",
    "protocolnumber": "study_number",
    "protocoltitle": "study_title",
    "testitem": "test_item_name",
    "testsample": "test_item_name",
    "lotnumber": "batch_lot",
    "batchnumber": "batch_lot",
    "solvent": "vehicle",
    "cellline": "test_system",
    "procedure": "method",
    "methodology": "method",
    "preparedby": "prepared_by",
    "reviewedby": "reviewed_by",
})


def _normalise_field_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def _set_template_text_node(node: etree._Element, value: str) -> None:
    node.text = value
    if value[:1].isspace() or value[-1:].isspace():
        node.set(f"{{{XML_NS}}}space", "preserve")
    else:
        node.attrib.pop(f"{{{XML_NS}}}space", None)


def _replace_across_paragraph(paragraph: etree._Element, replacements: dict[str, str]) -> int:
    """Replace text across runs while preserving every existing run's Word formatting."""
    count = 0
    nodes = paragraph.xpath(".//w:t", namespaces=NS)
    for marker, value in replacements.items():
        if not marker:
            continue
        text = "".join(node.text or "" for node in nodes)
        if marker not in text:
            continue

        start = text.find(marker)
        end = start + len(marker)
        offset = 0
        overlaps: list[tuple[etree._Element, int, int]] = []
        for node in nodes:
            node_text = node.text or ""
            node_end = offset + len(node_text)
            overlap_start = max(start, offset)
            overlap_end = min(end, node_end)
            if overlap_start < overlap_end:
                overlaps.append((node, overlap_start - offset, overlap_end - offset))
            offset = node_end
            if offset >= end:
                break

        if not overlaps:
            continue

        if len(overlaps) == 1:
            node, local_start, local_end = overlaps[0]
            node_text = node.text or ""
            _set_template_text_node(node, node_text[:local_start] + value + node_text[local_end:])
            count += 1
            continue

        marker_lengths = [overlap_end - overlap_start for _, overlap_start, overlap_end in overlaps]
        total_marker_length = sum(marker_lengths)
        cursor = 0
        segments: list[str] = []
        for index, length in enumerate(marker_lengths):
            if index == len(marker_lengths) - 1:
                segments.append(value[cursor:])
            else:
                take = round(len(value) * length / total_marker_length) if total_marker_length else 0
                segments.append(value[cursor:cursor + take])
                cursor += take

        for (node, local_start, local_end), segment in zip(overlaps, segments):
            node_text = node.text or ""
            _set_template_text_node(node, node_text[:local_start] + segment + node_text[local_end:])

        count += 1
    return count


def _replace_test_item_placeholders(paragraph: etree._Element, value: str) -> int:
    """Replace every starred TEST ITEM marker, including markers split across runs."""
    if not value:
        return 0

    count = 0
    while True:
        text = "".join(node.text or "" for node in paragraph.xpath(".//w:t", namespaces=NS))
        match = re.search(r"test\s+item\s*\*", text, flags=re.IGNORECASE)
        if match is None:
            return count
        marker = text[match.start():match.end()]
        count += _replace_across_paragraph(paragraph, {marker: value})


def _replace_across_paragraph_prefer_plain_run(
    paragraph: etree._Element, marker: str, value: str
) -> bool:
    """Replace a multi-run phrase using an unitalicized source run when possible."""
    italic_prefix = "In vitro "
    if marker.startswith(italic_prefix) and value.startswith(italic_prefix):
        marker = marker[len(italic_prefix):]
        value = value[len(italic_prefix):]
    nodes = paragraph.xpath(".//w:t", namespaces=NS)
    text = "".join(node.text or "" for node in nodes)
    start = text.find(marker)
    if start < 0:
        return False
    end = start + len(marker)
    offset = 0
    overlaps: list[tuple[etree._Element, int, int, int]] = []
    for node in nodes:
        node_text = node.text or ""
        node_end = offset + len(node_text)
        overlap_start = max(start, offset)
        overlap_end = min(end, node_end)
        if overlap_start < overlap_end:
            run = node.getparent()
            italic = bool(run is not None and run.xpath("./w:rPr/w:i[not(@w:val='0')]", namespaces=NS))
            overlaps.append((node, overlap_start - offset, overlap_end - offset, int(italic)))
        offset = node_end
        if offset >= end:
            break
    target = next((node for node, _, _, italic in overlaps if not italic), overlaps[0][0])
    offset = 0
    for node in nodes:
        node_text = node.text or ""
        node_end = offset + len(node_text)
        overlap_start = max(start, offset)
        overlap_end = min(end, node_end)
        if overlap_start < overlap_end:
            local_start = overlap_start - offset
            local_end = overlap_end - offset
            replacement = value if node is target else ""
            node.text = node_text[:local_start] + replacement + node_text[local_end:]
            if node.text[:1].isspace() or node.text[-1:].isspace():
                node.set(f"{{{XML_NS}}}space", "preserve")
        offset = node_end
        if offset >= end:
            break
    return True


def _set_paragraph_value(paragraph: etree._Element, value: str) -> None:
    runs = paragraph.xpath("./w:r", namespaces=NS)
    if runs:
        _set_paragraph_text(paragraph, value)
    else:
        text_node = etree.SubElement(paragraph, f"{W}r/{W}t")


def fill_study_plan_template(template_bytes: bytes, metadata: dict[str, Any]) -> bytes:
    """Return the selected template with its highlighted study values filled."""

    if not template_bytes:
        return _build_fallback_word_document(metadata)

    if hashlib.sha256(template_bytes).hexdigest() == DCA_TEMPLATE_SHA256:
        return fill_dca_study_plan_template(template_bytes, metadata)

    try:
        with ZipFile(BytesIO(template_bytes), "r") as source_zip:
            if "word/document.xml" not in source_zip.namelist():
                return _build_fallback_word_document(metadata)
            doc_hash = hashlib.sha256(template_bytes).hexdigest()
            if doc_hash != SUPPLIED_TEMPLATE_SHA256:
                filled, _ = fill_generic_template(template_bytes, metadata)
                return filled
    except Exception:
        return _build_fallback_word_document(metadata)

    sponsor_name = _inline(metadata.get("sponsor_name"))
    test_item = _inline(metadata.get("test_item_name"))
    if not sponsor_name or not test_item:
        raise ValueError("Sponsor Name and Test Item Name are required by the study-plan template.")

    source = BytesIO(template_bytes)
    output = BytesIO()
    with ZipFile(source, "r") as input_zip:
        names = input_zip.namelist()
        if "word/document.xml" not in names:
            return _build_fallback_word_document(metadata)
        xml_parts = {
            name: input_zip.read(name)
            for name in names
            if name.startswith("word/") and name.endswith(".xml")
        }
        document_xml = xml_parts["word/document.xml"]
        root = etree.fromstring(document_xml)
        tables = root.xpath("./w:body/w:tbl", namespaces=NS)
        if len(tables) < 8 or not root.xpath(".//w:sdt/w:sdtContent", namespaces=NS):
            raise ValueError(
                "The selected template does not include the supplied Cytotoxicity Study Plan fields. "
                "Please select the supplied template or a compatible copy."
            )

        # Cached TOC text, body paragraphs, and annexure titles carry the item
        # name and study number in addition to the highlighted data cells.
        _replace_text(root, "JUVÉDERM VOLBELLA XC", test_item)
        study_number = _inline(metadata.get("study_number"))
        _replace_text(root, "AB-GNT 166", study_number, preserve_case=False)
        _replace_text(root, "PHARMA CO., LTD", sponsor_name)
        _replace_text(root, "Pharma Co., Ltd", sponsor_name)

        sponsor_address = _inline(metadata.get("sponsor_address"))
        old_address = "864, Songpa-ro, Gangnam-gu, Seoul, Republic of Korea, 06734"
        _replace_text(root, old_address, sponsor_address, preserve_case=False)

        contact = _inline(metadata.get("sponsor_contact"))
        _replace_text(root, "Seo-yeon", contact)
        _replace_text(root, "SEO-YEON", contact)

        monitor = _inline(metadata.get("monitoring_scientist"))
        _replace_text(root, "Ha-joon", monitor)
        _replace_text(root, "HA-JOON", monitor)

        # The monitoring-scientist address is split across two runs in the
        # source table, unlike the single-run sponsor address elsewhere.
        if len(tables) > 1:
            _replace_in_cell(tables, 1, 6, 3, "864, Songpa-ro, Gangnam-gu, ", sponsor_address)
            _replace_in_cell(tables, 1, 6, 3, "Seoul, Republic of Korea, 06734", "")

        study_code = _inline(metadata.get("study_code"))
        _set_cell_paragraph(tables, 1, 2, 3, 0, study_code)

        study_director = _inline(metadata.get("study_director") or metadata.get("prepared_by"))
        _set_cell_paragraph(tables, 1, 9, 3, 0, study_director)
        director_email = _plain(metadata.get("study_director_email"))
        _set_cell_paragraph(
            tables,
            1,
            9,
            3,
            4,
            f"E-mail: {director_email}" if director_email else "",
        )
        _set_cell_paragraph(tables, 1, 10, 3, 0, _inline(metadata.get("study_coordinator")))
        _set_cell_paragraph(tables, 1, 11, 3, 0, _inline(metadata.get("study_personnel")))

        date_fields = {
            14: "study_initiation_date",
            15: "experimental_start_date",
            16: "experimental_completion_date",
            17: "draft_report_submission_date",
        }
        for row_index, field in date_fields.items():
            value = _plain(metadata.get(field))
            # Preserve source placeholders such as "xx October 2025" so the
            # plan reflects the uploaded study-details workbook accurately.
            _set_cell_paragraph(tables, 1, row_index, 3, 0, value)

        intended_use = _inline(metadata.get("study_objective"))
        intended_use_paragraph = _find_body_paragraph(root, "Intended to be used")
        if intended_use and intended_use_paragraph is not None:
            _set_paragraph_text(intended_use_paragraph, intended_use)

        # Test-item specification table.
        extraction = _inline(metadata.get("test_item_extraction_details"))
        extraction_label = "Weight (mass/volume):"
        if extraction.lower().startswith("weight"):
            extraction_value = extraction
        else:
            extraction_value = f"{extraction_label}{extraction}" if extraction else ""
        _set_cell_paragraph(tables, 2, 1, 2, 0, extraction_value)
        _set_cell_paragraph(tables, 2, 2, 2, 0, _plain(metadata.get("test_item_weight")))

        appearance = _plain(metadata.get("physical_appearance"))
        appearance_lines = [line.strip() for line in appearance.splitlines() if line.strip()]
        appearance_lines += [""] * max(0, 2 - len(appearance_lines))
        if len(appearance_lines) > 2:
            appearance_lines = [appearance_lines[0], " ".join(appearance_lines[1:])]
        _set_cell_paragraph(tables, 2, 3, 2, 1, appearance_lines[0])
        _set_cell_paragraph(tables, 2, 3, 2, 2, appearance_lines[1])
        _set_cell_paragraph(tables, 2, 4, 2, 0, _plain(metadata.get("quantity_sent")))
        _set_cell_paragraph(tables, 2, 5, 2, 0, _plain(metadata.get("batch_number")) or "N/A")
        _set_cell_paragraph(tables, 2, 6, 2, 0, _plain(metadata.get("lot_number") or metadata.get("batch_lot")))
        _set_cell_paragraph(tables, 2, 7, 2, 0, _inline(metadata.get("manufacturer_name")))
        _set_cell_paragraph(tables, 2, 7, 2, 1, _inline(metadata.get("manufacturer_address")))
        _set_cell_paragraph(tables, 2, 8, 2, 0, _plain(metadata.get("manufacture_date")))
        _set_cell_paragraph(tables, 2, 9, 2, 0, _plain(metadata.get("expiry_date")))
        _set_cell_paragraph(tables, 2, 10, 2, 0, _plain(metadata.get("storage_condition")))
        test_item_code = _plain(metadata.get("test_item_code") or metadata.get("test_item_id"))
        _set_cell_paragraph(tables, 2, 11, 2, 0, test_item_code)

        # Keep the narrative method ratio in step with the uploaded details.
        ratio = _normalise_ratio(extraction)
        for paragraph in root.xpath(".//w:p", namespaces=NS):
            narrative_ratio = f"{extraction} (weight/volume)" if extraction else ""
            if _replace_in_paragraph(paragraph, "0.2g/mL (weight/volume)", narrative_ratio):
                break
        if len(tables) > 4:
            _set_cell_paragraph(tables, 4, 4, 2, 0, ratio)

        # The supplied agreement keeps its fixed facility-side personnel;
        # sponsor and study-specific signatories follow the uploaded fields.
        if len(tables) > 7:
            director_name = re.sub(r"^(?:Ms\.?|Mr\.?|Dr\.?)\s*", "", study_director, flags=re.IGNORECASE)
            honorific = re.match(r"^(Ms\.?|Mr\.?|Dr\.?)\s*", study_director, flags=re.IGNORECASE)
            director_signature = f"{honorific.group(1) if honorific else 'Ms.'} {director_name.upper()}" if director_name else ""
            sponsor_signature = contact.upper()
            monitor_signature = monitor.upper()
            _replace_cell_text_exact(tables, 7, 2, 1, "Ms. POOJANA H. C.", director_signature)
            _replace_cell_text_exact(tables, 7, 2, 3, "SEO-YEON", sponsor_signature)
            _replace_cell_text_exact(tables, 7, 4, 3, "HA-JOON", monitor_signature)

        # If an explicit study title was supplied, use it in the cover, details
        # table, and agreement while retaining the template's paragraph style.
        title = _inline(metadata.get("study_title"))
        if title:
            _set_cell_paragraph(tables, 0, 8, 2, 0, title, prefer_highlight=False)
            _set_cell_paragraph(tables, 1, 0, 3, 0, title, prefer_highlight=False)
            agreement = _find_body_paragraph(root, "The Study Plan for Study No.")
            if agreement is not None:
                agreement_text = (
                    f'The Study Plan for Study No. {study_number}, “{title}” has been mutually agreed and signed.'
                )
                _set_paragraph_text(agreement, agreement_text)

        xml_parts["word/document.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        # Study number can also appear in footer XML; patch those targeted
        # Word parts without changing relationships or embedded media.
        for part_name, part_bytes in list(xml_parts.items()):
            if part_name == "word/document.xml":
                continue
            if "footer" not in part_name.lower() and "header" not in part_name.lower():
                continue
            part_root = etree.fromstring(part_bytes)
            if _replace_text(part_root, "AB-GNT 166", study_number, preserve_case=False):
                xml_parts[part_name] = etree.tostring(
                    part_root, xml_declaration=True, encoding="UTF-8", standalone=True
                )

        with ZipFile(output, "w", compression=ZIP_DEFLATED) as output_zip:
            for info in input_zip.infolist():
                output_zip.writestr(info, xml_parts.get(info.filename, input_zip.read(info.filename)))

    return output.getvalue()


def _starred_placeholder_replacements(metadata: dict[str, Any]) -> dict[str, str]:
    """Map common labeled ``*`` placeholders to uploaded metadata values."""
    director = _inline(metadata.get("study_director_name") or metadata.get("study_director") or metadata.get("prepared_by"))
    manufacturer = ", ".join(
        value
        for value in (
            _inline(metadata.get("manufacturer_name")),
            _inline(metadata.get("manufacturer_address")),
        )
        if value
    )
    replacements = {
        "Study Sponsor*": _inline(metadata.get("sponsor_name")),
        "STUDY SPONSOR*": _inline(metadata.get("sponsor_name")),
        "Sponsor Name*": _inline(metadata.get("sponsor_name")),
        "SPONSOR REPRESENTATIVE*": _inline(metadata.get("sponsor_contact")),
        "Address*": _inline(metadata.get("sponsor_address")),
        "ADDRESS*": _inline(metadata.get("sponsor_address")),
        "Study Number*": _inline(metadata.get("study_number")),
        "STUDY NUMBER*": _inline(metadata.get("study_number")),
        "Study Title*": _inline(metadata.get("study_title")),
        "STUDY TITLE*": _inline(metadata.get("study_title")),
        "Test Item*": _inline(metadata.get("test_item_name")),
        "TEST ITEM*": _inline(metadata.get("test_item_name")),
        "Test Item Name*": _inline(metadata.get("test_item_name")),
        "Test Item Code by Test Facility*": _inline(metadata.get("test_item_code") or metadata.get("test_item_id")),
        "Study Director*": director,
        "STUDY DIRECTOR*": director,
        "Study Director Name*": director,
        "STUDY DIRECTOR NAME*": director,
        "Study Director Email*": _inline(metadata.get("study_director_email")),
        "STUDY DIRECTOR EMAIL*": _inline(metadata.get("study_director_email")),
        "Study Co-ordinator*": _inline(metadata.get("study_coordinator")),
        "STUDY CO-ORDINATOR*": _inline(metadata.get("study_coordinator")),
        "Monitoring Scientist*": _inline(metadata.get("monitoring_scientist")),
        "MONITORING SCIENTIST*": _inline(metadata.get("monitoring_scientist")),
        "Study Personnel*": _inline(metadata.get("study_personnel")),
        "STUDY PERSONNEL*": _inline(metadata.get("study_personnel")),
        "Study Initiation Date*": _inline(metadata.get("study_initiation_date")),
        "STUDY INITIATION DATE*": _inline(metadata.get("study_initiation_date")),
        "Experimental Starting Date*": _inline(metadata.get("experimental_start_date")),
        "EXPERIMENTAL STARTING DATE*": _inline(metadata.get("experimental_start_date")),
        "Experimental Completion Date*": _inline(metadata.get("experimental_completion_date")),
        "EXPERIMENTAL COMPLETION DATE*": _inline(metadata.get("experimental_completion_date")),
        "Draft Report Submission Date*": _inline(metadata.get("draft_report_submission_date")),
        "DRAFT REPORT SUBMISSION DATE*": _inline(metadata.get("draft_report_submission_date")),
        "Quantity Sent*": _inline(metadata.get("quantity_sent")),
        "QUANTITY SENT*": _inline(metadata.get("quantity_sent")),
        "Batch No*": _inline(metadata.get("batch_number") or metadata.get("batch_lot") or metadata.get("batch_lot")),
        "BATCH NO*": _inline(metadata.get("batch_number") or metadata.get("batch_lot") or metadata.get("batch_lot")),
        "Lot No*": _inline(metadata.get("lot_number") or metadata.get("batch_lot")),
        "LOT NO*": _inline(metadata.get("lot_number") or metadata.get("batch_lot")),
        "Manufactured/Supplied*": manufacturer,
        "MANUFACTURED/SUPPLIED*": manufacturer,
        "Date of Manufacture*": _inline(metadata.get("manufacture_date")),
        "DATE OF MANUFACTURE*": _inline(metadata.get("manufacture_date")),
        "Date of Expiry*": _inline(metadata.get("expiry_date")),
        "DATE OF EXPIRY*": _inline(metadata.get("expiry_date")),
        "Storage Condition*": _inline(metadata.get("storage_condition")),
        "STORAGE CONDITION*": _inline(metadata.get("storage_condition")),
        "Physical Appearance*": _inline(metadata.get("physical_appearance")),
        "PHYSICAL APPEARANCE*": _inline(metadata.get("physical_appearance")),
        "Quality Assurance Unit*": _inline(metadata.get("quality_assurance_unit")),
        "QUALITY ASSURANCE UNIT*": _inline(metadata.get("quality_assurance_unit")),
        "TEST FACILITY MAN*": _inline(metadata.get("test_facility_manager")),
    }
    for field, label in FIELDS.items():
        value = _inline(metadata.get(field))
        if value:
            replacements[f"{label}*"] = value
            replacements[f"{label} *"] = value
    return {marker: value for marker, value in replacements.items() if value}


def fill_generic_template(template_bytes: bytes, metadata: dict[str, Any]) -> tuple[bytes, int]:
    """Patch text in-place while preserving the uploaded document's layout and formatting.

    The Word package is copied as-is; only text nodes and recognized value cells are changed.
    Paragraph properties, run formatting, fonts, spacing, tables, sections, headers, and footers
    remain sourced from the uploaded template rather than being rebuilt.
    """
    count = 0
    replacements = _starred_placeholder_replacements(metadata)

    if len(template_bytes) < 4:
        return _build_fallback_word_document(metadata), 0

    if not template_bytes:
        return _build_fallback_word_document(metadata), 0

    for field, label in FIELDS.items():
        value = _plain(metadata.get(field))
        if not value:
            continue
        replacements[field] = value
        replacements[label] = value
        replacements[field.replace("_", " ")] = value
        replacements[label.replace(" ", "_")] = value

    test_item = _plain(metadata.get("test_item_name"))
    match_count = 0
    output = BytesIO()
    try:
        with ZipFile(BytesIO(template_bytes), "r") as input_zip:
            names = input_zip.namelist()
            if "word/document.xml" not in names:
                return _build_fallback_word_document(metadata), 0

            with ZipFile(output, "w", compression=ZIP_DEFLATED) as output_zip:
                for info in input_zip.infolist():
                    part_bytes = input_zip.read(info.filename)
                    if not info.filename.startswith("word/") or not info.filename.endswith(".xml"):
                        output_zip.writestr(info, part_bytes)
                        continue

                    try:
                        root = etree.fromstring(part_bytes)
                    except Exception:
                        output_zip.writestr(info, part_bytes)
                        continue

                    for paragraph in root.xpath(".//w:p", namespaces=NS):
                        match_count += _replace_test_item_placeholders(paragraph, test_item)
                        nodes = paragraph.xpath(".//w:t", namespaces=NS)
                        text = "".join(node.text or "" for node in nodes)

                        for marker, value in _starred_placeholder_replacements(metadata).items():
                            if marker in text:
                                match_count += _replace_across_paragraph(paragraph, {marker: value})
                                text = "".join(node.text or "" for node in nodes)

                        for field, value in replacements.items():
                            for wrapper in (
                                f"{{{{{field}}}}}", f"{{{field}}}",
                                f"[[{field}]]", f"<<{field}>>",
                            ):
                                if wrapper in text:
                                    match_count += _replace_across_paragraph(
                                        paragraph, {wrapper: value}
                                    )
                                    text = "".join(node.text or "" for node in nodes)

                        for field, label in FIELDS.items():
                            value = _plain(metadata.get(field))
                            if not value:
                                continue
                            pattern = re.compile(
                                re.escape(label) + r"\s*[:：]\s*",
                                re.IGNORECASE,
                            )
                            label_match = pattern.search(text)
                            if label_match:
                                tail = text[label_match.end():].strip()
                                if tail:
                                    match_count += _replace_across_paragraph(
                                        paragraph, {tail: value}
                                    )
                                else:
                                    marker = label_match.group(0)
                                    match_count += _replace_across_paragraph(
                                        paragraph, {marker: marker + value}
                                    )
                                text = "".join(node.text or "" for node in nodes)

                    for table in root.xpath(".//w:tbl", namespaces=NS):
                        rows = table.findall(f"{W}tr")
                        for row_index, row in enumerate(rows):
                            cells = row.findall(f"{W}tc")
                            if len(cells) < 2:
                                continue
                            label_text = "".join(cells[0].xpath(".//w:t/text()", namespaces=NS)).strip()
                            field = _GENERIC_ALIASES.get(_normalise_field_key(label_text))
                            if not field:
                                continue
                            value = _plain(metadata.get(field))
                            if not value:
                                continue

                            value_cell = cells[1]
                            if not _value_cell_is_wide_enough(cells[0], value_cell, value):
                                # Some cover pages put category headings across
                                # several columns, with a narrow spacer beside
                                # each heading and the value directly below it.
                                if len(cells) < 3 or row_index + 1 >= len(rows):
                                    continue
                                next_cells = rows[row_index + 1].findall(f"{W}tc")
                                if not next_cells:
                                    continue
                                candidate = next_cells[0]
                                candidate_text = "".join(
                                    candidate.xpath(".//w:t/text()", namespaces=NS)
                                ).strip()
                                if candidate_text or not _value_cell_is_wide_enough(
                                    cells[0], candidate, value
                                ):
                                    continue
                                value_cell = candidate

                            paragraphs = value_cell.findall(f"{W}p")
                            if paragraphs:
                                _set_paragraph_text(paragraphs[0], value)
                                match_count += 1

                    patched = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                    output_zip.writestr(info, patched)
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("Could not read the selected Word template. Upload a valid .doc or .docx file.") from exc

    return output.getvalue(), match_count


def _build_fallback_word_document(metadata: dict[str, Any]) -> bytes:
    """Create a valid Word document when an uploaded file cannot be parsed in place."""
    document = Document()
    title = _inline(metadata.get("study_title") or "Cytotoxicity Study Plan")
    document.add_heading(title, level=1)
    table = document.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    for field, label in FIELDS.items():
        value = _plain(metadata.get(field))
        if value:
            cells = table.add_row().cells
            cells[0].text = label
            cells[1].text = value
    output = BytesIO()
    document.save(output)
    return output.getvalue()
