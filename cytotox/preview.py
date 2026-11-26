"""Render generated Word files for a page-accurate in-app preview."""

from __future__ import annotations

import subprocess
import tempfile
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .template_filler import apply_dca_toc_formatting
from .template_input import libreoffice_executable


PREVIEW_RENDERER_VERSION = "5"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}


def _apply_toc_tab_compatibility(document_bytes: bytes) -> bytes:
    """Work around LibreOffice dropping TOC1 tab stops during PDF export.

    The direct paragraph settings are applied only to an in-memory preview copy.
    The generated DCA DOCX receives matching TOC formatting during filling.
    """

    try:
        with ZipFile(BytesIO(document_bytes), "r") as source:
            names = set(source.namelist())
            if not {"word/document.xml", "word/styles.xml"}.issubset(names):
                return document_bytes
            styles = etree.fromstring(source.read("word/styles.xml"))
            toc_style = styles.xpath(
                ".//w:style[@w:styleId='TOC1']", namespaces=NS
            )
            if not toc_style:
                return document_bytes
            style_properties = toc_style[0].find(f"{W}pPr")
            if style_properties is None:
                return document_bytes
            style_indent = style_properties.find(f"{W}ind")
            indent_position = (
                style_indent.get(f"{W}start") or style_indent.get(f"{W}left")
                if style_indent is not None
                else None
            )
            style_tabs = style_properties.findall(f"{W}tabs/{W}tab")
            left_positions = {
                tab.get(f"{W}pos")
                for tab in style_tabs
                if tab.get(f"{W}val") == "clear"
                and tab.get(f"{W}pos") == indent_position
            }
            right_tabs = [
                tab
                for tab in style_tabs
                if tab.get(f"{W}val") == "right"
                and tab.get(f"{W}leader") == "dot"
                and tab.get(f"{W}pos", "").isdigit()
            ]
            if not left_positions or not right_tabs:
                return document_bytes
            left_position = next(iter(left_positions))
            right_position = max(
                (tab.get(f"{W}pos") for tab in right_tabs), key=int
            )

            document = etree.fromstring(source.read("word/document.xml"))
            instruction = " ".join(
                document.xpath(".//w:instrText/text()", namespaces=NS)
            )
            if "Table,6,Figure,7,Appendix,8,Annexure,9" not in instruction:
                return document_bytes
            contents = document.xpath(
                ".//w:sdtContent[.//w:instrText[contains(., 'TOC')]]",
                namespaces=NS,
            )
            if contents:
                paragraphs = [
                    paragraph
                    for content in contents
                    for paragraph in content.xpath(".//w:p", namespaces=NS)
                    if paragraph.xpath(
                        "string(./w:pPr/w:pStyle/@w:val)", namespaces=NS
                    )
                    == "TOC1"
                ]
            elif document.xpath(".//w:instrText[contains(., 'TOC')]", namespaces=NS):
                paragraphs = document.xpath(
                    ".//w:p[./w:pPr/w:pStyle/@w:val='TOC1']", namespaces=NS
                )
            else:
                paragraphs = []
            if not paragraphs:
                return document_bytes

            for paragraph in paragraphs:
                properties = paragraph.find(f"{W}pPr")
                if properties is None:
                    properties = etree.Element(f"{W}pPr")
                    paragraph.insert(0, properties)

                existing_tabs = properties.find(f"{W}tabs")
                if existing_tabs is not None:
                    properties.remove(existing_tabs)
                tabs = etree.Element(f"{W}tabs")
                left = etree.SubElement(tabs, f"{W}tab")
                left.set(f"{W}val", "left")
                left.set(f"{W}pos", left_position)
                right = etree.SubElement(tabs, f"{W}tab")
                right.set(f"{W}val", "right")
                right.set(f"{W}pos", right_position)
                right.set(f"{W}leader", "dot")

                style = properties.find(f"{W}pStyle")
                tabs_index = list(properties).index(style) + 1 if style is not None else 0
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

            output = BytesIO()
            with ZipFile(output, "w", ZIP_DEFLATED) as destination:
                for info in source.infolist():
                    data = source.read(info.filename)
                    if info.filename == "word/document.xml":
                        data = etree.tostring(
                            document,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone=True,
                        )
                    destination.writestr(info, data)
            return output.getvalue()
    except Exception:
        # A compatibility adjustment must never prevent the normal preview.
        return document_bytes


def render_docx_pdf(document_bytes: bytes) -> bytes:
    """Render a generated DOCX to PDF without modifying the Word file."""

    executable = libreoffice_executable()
    if executable is None:
        raise RuntimeError(
            "LibreOffice is required to render the page-accurate preview. "
            "Install LibreOffice, then restart the app."
        )

    with tempfile.TemporaryDirectory(prefix="cytotox-preview-") as temp_dir:
        work_dir = Path(temp_dir)
        source = work_dir / "generated-study-document.docx"
        output_dir = work_dir / "pdf"
        profile_dir = work_dir / "profile"
        output_dir.mkdir()
        profile_dir.mkdir()
        preview_bytes = _apply_toc_tab_compatibility(
            apply_dca_toc_formatting(document_bytes)
        )
        source.write_bytes(preview_bytes)

        result = subprocess.run(
            [
                executable,
                f"-env:UserInstallation={profile_dir.as_uri()}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(output_dir),
                str(source),
            ],
            capture_output=True,
            text=True,
            timeout=90,
            check=False,
        )
        rendered_pdf = output_dir / "generated-study-document.pdf"
        if result.returncode != 0 or not rendered_pdf.is_file():
            detail = (result.stderr or result.stdout).strip()
            message = "LibreOffice could not render the Word document preview."
            if detail:
                message = f"{message} {detail[:300]}"
            raise RuntimeError(message)
        return rendered_pdf.read_bytes()

