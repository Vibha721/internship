"""Start a packaged app and verify its local health endpoint."""

from __future__ import annotations

import argparse
from pathlib import Path
import re
import subprocess
import sys
import threading
import time
from typing import TextIO
from urllib.error import URLError
from urllib.request import urlopen


def _read_output(output: list[str]) -> str:
    return "".join(output)


def _capture_output(pipe: TextIO, output: list[str]) -> None:
    for line in pipe:
        output.append(line)


def smoke_test(app_dir: Path) -> None:
    executable = app_dir / ("CytotoxicityStudyApp.exe" if sys.platform == "win32" else "CytotoxicityStudyApp")
    if not executable.is_file():
        raise FileNotFoundError(f"Packaged executable not found: {executable}")

    process = subprocess.Popen(
        [str(executable)],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    if process.stdout is None:
        raise RuntimeError("Could not capture packaged app startup output.")

    output: list[str] = []
    reader = threading.Thread(
        target=_capture_output,
        args=(process.stdout, output),
        daemon=True,
    )
    reader.start()
    try:
        deadline = time.monotonic() + 60
        while time.monotonic() < deadline:
            if process.poll() is not None:
                raise RuntimeError(
                    f"Packaged app exited with status {process.returncode}.\n{_read_output(output)}"
                )

            match = re.search(
                r"(?:Uvicorn server started on|URL: http://)127\.0\.0\.1:(\d+)",
                _read_output(output),
            )
            if match:
                url = f"http://127.0.0.1:{match.group(1)}/_stcore/health"
                try:
                    with urlopen(url, timeout=2) as response:
                        if response.read().decode("utf-8").strip() == "ok":
                            print(f"Packaged app health check passed at {url}.")
                            return
                except URLError:
                    pass
            time.sleep(0.25)
        raise TimeoutError(
            f"Packaged app did not become healthy within 60 seconds.\n{_read_output(output)}"
        )
    finally:
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait()
        reader.join(timeout=1)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--app-dir", type=Path, required=True)
    arguments = parser.parse_args()
    smoke_test(arguments.app_dir)


if __name__ == "__main__":
    main()
