"""Fill the supplied Cytotoxicity Study Report Word template in place."""

from __future__ import annotations

import re
from io import BytesIO
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def _value(metadata: dict[str, Any], key: str) -> str:
    value = metadata.get(key)
    if value is None:
        return ""
    return ", ".join(part.strip() for part in str(value).splitlines() if part.strip()).strip()


def _set_text(node: etree._Element, value: str) -> None:
    node.text = value
    if value[:1].isspace() or value[-1:].isspace():
        node.set(f"{{{XML_NS}}}space", "preserve")


def _replace_text_nodes(nodes: list[etree._Element], old: str, new: str) -> int:
    """Replace text even when Word split a placeholder across multiple runs."""

    if not old:
        return 0
    count = 0
    cursor = 0
    pattern = re.compile(re.escape(old), re.IGNORECASE)
    while True:
        texts = [node.text or "" for node in nodes]
        flat = "".join(texts)
        match = pattern.search(flat, cursor)
        if match is None:
            return count
        start, end = match.span()
        replacement = new
        if replacement and end < len(flat) and flat[end].isalpha() and not replacement[-1:].isspace():
            replacement += " "
        offset = 0
        start_node = end_node = None
        start_offset = end_offset = 0
        for index, value in enumerate(texts):
            next_offset = offset + len(value)
            if start_node is None and offset <= start < next_offset:
                start_node = index
                start_offset = start - offset
            if offset < end <= next_offset:
                end_node = index
                end_offset = end - offset
                break
            offset = next_offset
        if start_node is None or end_node is None:
            return count

        if start_node == end_node:
            value = texts[start_node]
            _set_text(nodes[start_node], value[:start_offset] + replacement + value[end_offset:])
        else:
            _set_text(nodes[start_node], texts[start_node][:start_offset] + replacement)
            for index in range(start_node + 1, end_node):
                _set_text(nodes[index], "")
            _set_text(nodes[end_node], texts[end_node][end_offset:])
        cursor = start + len(replacement)
        count += 1


def _replace_in_paragraph(paragraph: etree._Element, old: str, new: str) -> int:
    return _replace_text_nodes(paragraph.xpath(".//w:t", namespaces=NS), old, new)


def _replace_all(root: etree._Element, old: str, new: str) -> None:
    for paragraph in root.xpath(".//w:p", namespaces=NS):
        _replace_in_paragraph(paragraph, old, new)


def fill_study_report_template(
    template_bytes: bytes,
    metadata: dict[str, Any],
    results: dict[str, Any],
) -> bytes:
    """Replace fields while retaining the supplied report's Word package/layout."""

    director = _value(metadata, "study_director") or _value(metadata, "prepared_by")
    replacements = {
        "STUDY DIRECTOR NAME*": _value(metadata, "study_director_name") or director,
        "QUALITY ASSURANCE UNIT*": _value(metadata, "quality_assurance_unit"),
        "TEST FACILITY MAN*": _value(metadata, "test_facility_manager"),
        "STUDY SPONSOR*": _value(metadata, "sponsor_name"),
        "SPONSOR REPRESENTATIVE*": _value(metadata, "sponsor_contact"),
        "ADDRESS*": _value(metadata, "sponsor_address"),
        "STUDY NUMBER*": _value(metadata, "study_number"),
        "STUDY DIRECTOR EMAIL*": _value(metadata, "study_director_email"),
        "STUDY DIRECTOR*": director,
        "MONITORING SCIENTIST*": _value(metadata, "monitoring_scientist"),
        "STUDY CO-ORDINATOR*": _value(metadata, "study_coordinator"),
        "STUDY PERSONNEL*": _value(metadata, "study_personnel"),
        "STUDY INITIATION DATE*": _value(metadata, "study_initiation_date"),
        "EXPERIMENTAL STARTING DATE*": _value(metadata, "experimental_start_date"),
        "EXPERIMENTAL COMPLETION DATE*": _value(metadata, "experimental_completion_date"),
        "DRAFT REPORT SUBMISSION DATE*": _value(metadata, "draft_report_submission_date"),
        "TEST ITEM CODE BY TEST FACILITY*": _value(metadata, "test_item_code") or _value(metadata, "test_item_id"),
        "DETAILS TO BE CONSIDERED*": _value(metadata, "test_item_extraction_details"),
        "QUANTITY SENT*": _value(metadata, "quantity_sent"),
        "BATCH NO*": _value(metadata, "batch_number"),
        "LOT NO*": _value(metadata, "lot_number") or _value(metadata, "batch_lot"),
        "MANUFACTURED/SUPPLIED*": _value(metadata, "manufacturer_name"),
        "DATE OF MANUFACTURE*": _value(metadata, "manufacture_date"),
        "DATE OF EXPIRY*": _value(metadata, "expiry_date"),
        "STORAGE CONDITION*": _value(metadata, "storage_condition"),
        "TEST ITEM*": _value(metadata, "test_item_name"),
        "R1*": str(results.get("replicate_1", "")).strip(),
        "R2*": str(results.get("replicate_2", "")).strip(),
        "R3*": str(results.get("replicate_3", "")).strip(),
        "GRADE*": str(results.get("grade", "")).strip(),
        "RESULT*": str(results.get("result", "")).strip(),
    }
    # Longer keys must be handled before shorter overlapping fields.
    replacements = dict(sorted(replacements.items(), key=lambda item: len(item[0]), reverse=True))

    output = BytesIO()
    with ZipFile(BytesIO(template_bytes), "r") as source, ZipFile(output, "w", ZIP_DEFLATED) as target:
        for item in source.infolist():
            content = source.read(item.filename)
            if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                try:
                    root = etree.fromstring(content)
                except etree.XMLSyntaxError:
                    target.writestr(item, content)
                    continue
                for old, new in replacements.items():
                    _replace_all(root, old, new)
                _replace_all(root, "xx March 2026", _value(metadata, "study_completion_date"))
                _replace_all(root, "Xx xxxx 2026", _value(metadata, "study_completion_date"))
                _replace_all(root, "1 mL per device", _value(metadata, "test_item_weight"))
                content = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            target.writestr(item, content)
    return output.getvalue()
