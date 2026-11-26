"""Assemble the PyInstaller output into a transferable offline zip."""

from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import stat
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


ROOT = Path(__file__).resolve().parent.parent
APP_NAME = "CytotoxicityStudyApp"
PLATFORMS = {
    "linux-x86_64": ("Start Cytotoxicity App.sh", "linux"),
    "windows-x86_64": ("Start Cytotoxicity App.bat", "windows"),
    "macos-arm64": ("Start Cytotoxicity App.command", "macos"),
    "macos-x86_64": ("Start Cytotoxicity App.command", "macos"),
}

README = """Cytotoxicity Study Document Generator — offline app

This package includes its Python runtime and application dependencies. It does
not need internet access or a Python installation on the target laptop.

Start the app:
- Windows: double-click "Start Cytotoxicity App.bat".
- macOS: double-click "Start Cytotoxicity App.command". The app is unsigned;
  macOS may require Control-click > Open the first time.
- Linux: run `./"Start Cytotoxicity App.sh"` from a terminal in the extracted
  folder, or run `chmod +x "Start Cytotoxicity App.sh"` first if needed.

Keep the terminal window open while using the app. It opens in your browser at
a local 127.0.0.1 address. The server listens only on this laptop. To stop it,
close the terminal window or press Ctrl+C there.

Choose the zip built for the target computer:
- linux-x86_64: 64-bit Intel/AMD Linux
- windows-x86_64: 64-bit Intel/AMD Windows
- macos-arm64: Apple silicon Macs (M-series)
- macos-x86_64: Intel Macs

The supplied Word templates and editable DOCX generation are included.
Page-accurate PDF previews, legacy .doc conversion, and uploaded templates that
need LibreOffice conversion require LibreOffice to be installed separately.
Internet is not required for document generation.

To transfer this app to a laptop that is already offline, download the matching
zip on a connected computer and copy it over using a USB drive or local network.
"""

WINDOWS_LAUNCHER = r"""@echo off
setlocal
"%~dp0CytotoxicityStudyApp\CytotoxicityStudyApp.exe"
if errorlevel 1 pause
"""
POSIX_LAUNCHER = """#!/bin/sh
set -eu
SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
exec "$SCRIPT_DIR/CytotoxicityStudyApp/CytotoxicityStudyApp"
"""


def _write_zip_file(archive: ZipFile, file_path: Path, archive_path: str) -> None:
    info = ZipInfo(archive_path)
    info.compress_type = ZIP_DEFLATED
    info.external_attr = (file_path.stat().st_mode & 0xFFFF) << 16
    archive.writestr(info, file_path.read_bytes())


def package(platform: str) -> Path:
    launcher_name, os_family = PLATFORMS[platform]
    app_dir = ROOT / "dist" / APP_NAME
    if not app_dir.is_dir():
        raise FileNotFoundError(f"PyInstaller output not found: {app_dir}")
    executable = app_dir / (f"{APP_NAME}.exe" if os_family == "windows" else APP_NAME)
    if not executable.is_file():
        raise FileNotFoundError(f"Packaged app executable not found: {executable}")

    dist_dir = ROOT / "dist"
    archive_path = dist_dir / f"cytotoxicity-offline-{platform}.zip"
    package_root_name = f"cytotoxicity-offline-{platform}"

    with tempfile.TemporaryDirectory(prefix="cytotoxicity-package-", dir=dist_dir) as temporary:
        staging_dir = Path(temporary) / package_root_name
        shutil.copytree(app_dir, staging_dir / APP_NAME)
        launcher_path = staging_dir / launcher_name
        launcher_path.write_text(
            WINDOWS_LAUNCHER if os_family == "windows" else POSIX_LAUNCHER,
            encoding="utf-8",
            newline="",
        )
        if os_family != "windows":
            launcher_path.chmod(launcher_path.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
        (staging_dir / "OFFLINE_README.txt").write_text(README, encoding="utf-8")

        with ZipFile(archive_path, "w", ZIP_DEFLATED) as archive:
            for file_path in sorted(staging_dir.rglob("*")):
                if file_path.is_file():
                    relative_path = file_path.relative_to(staging_dir.parent).as_posix()
                    _write_zip_file(archive, file_path, relative_path)
    return archive_path


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--platform", choices=sorted(PLATFORMS), required=True)
    arguments = parser.parse_args()
    print(package(arguments.platform))


if __name__ == "__main__":
    main()
