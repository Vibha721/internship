"""Normalize uploaded legacy Word templates for the document generator."""

from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
from io import BytesIO
from pathlib import Path
from zipfile import BadZipFile, ZipFile


SUPPLIED_LEGACY_DOC_SHA256 = "e4375d76735c9584ebc3c066aea5c466ffd66914e856100279aa8199b2c1157f"


def libreoffice_executable() -> str | None:
    for name in ("soffice", "libreoffice"):
        executable = shutil.which(name)
        if executable:
            return executable

    candidates = [
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice",
    ]
    runtime_root = Path.home() / ".cache/codex-runtimes"
    if runtime_root.exists():
        candidates.extend(runtime_root.glob("*/dependencies/bin/override/soffice"))
    return next((str(candidate) for candidate in candidates if candidate.is_file()), None)


def is_valid_docx(content: bytes) -> bool:
    """Accept any valid OOXML ZIP package containing the main Word document."""
    try:
        with ZipFile(BytesIO(content), "r") as package:
            names = {name.replace("\\", "/") for name in package.namelist()}
            if "word/document.xml" not in names:
                return False
            if package.testzip():
                return False
            package.read("word/document.xml")
            return True
    except (BadZipFile, OSError, RuntimeError, KeyError):
        return False


def normalize_template_upload(
    filename: str,
    content: bytes,
    *,
    supplied_docx: bytes | None = None,
) -> bytes:
    """Return an editable .docx payload for an uploaded .doc or .docx file."""

    suffix = Path(filename).suffix.lower()
    supported = {".docx", ".doc", ".rtf", ".odt", ".xml", ".dotx", ".docm"}
    if suffix not in supported:
        raise ValueError("Choose a Word-compatible document (.doc, .docx, .docm, .dotx, .rtf, .odt, or .xml).")

    if not content:
        raise ValueError("The uploaded Word file is empty. Re-save it and upload it again.")

    if is_valid_docx(content):
        return content

    if (
        supplied_docx is not None
        and hashlib.sha256(content).hexdigest() == SUPPLIED_LEGACY_DOC_SHA256
    ):
        return supplied_docx

    executable = libreoffice_executable()
    if executable is None:
        raise ValueError(
            "This legacy .doc template needs LibreOffice to convert. Install LibreOffice or save the file as .docx, then upload it again."
        )

    with tempfile.TemporaryDirectory(prefix="cytotox-template-") as temp_dir:
        work_dir = Path(temp_dir)
        actual_suffix = suffix
        if content.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
            actual_suffix = ".doc"
        elif content.startswith((b"{\\rtf", b"<!DOCTYPE", b"<html")):
            actual_suffix = ".rtf" if content.startswith(b"{\\rtf") else ".html"
        source = work_dir / f"uploaded-template{actual_suffix}"
        output_dir = work_dir / "converted"
        profile_dir = work_dir / "profile"
        output_dir.mkdir()
        profile_dir.mkdir()
        source.write_bytes(content)

        result = subprocess.run(
            [
                executable,
                f"-env:UserInstallation={profile_dir.as_uri()}",
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                str(output_dir),
                str(source),
            ],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
        converted = output_dir / "uploaded-template.docx"
        if result.returncode != 0 or not converted.is_file():
            detail = (result.stderr or result.stdout).strip()
            message = "LibreOffice could not convert this .doc file."
            if detail:
                message = f"{message} {detail[:300]}"
            raise ValueError(message)
        return converted.read_bytes()
