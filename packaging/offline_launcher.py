"""Launch the bundled Streamlit app on this computer only."""

from __future__ import annotations

import os
from pathlib import Path
import socket
import sys

# Streamlit loads app.py dynamically, so import its dependencies for PyInstaller.
import cytotox.excel_io
import cytotox.metadata
import cytotox.preview
import cytotox.report_excel
import cytotox.report_generator
import cytotox.report_template_filler
import cytotox.template_filler
import cytotox.template_input
import pandas
import streamlit
import streamlit_pdf
from streamlit.web import cli


def main() -> None:
    bundle_dir = Path(sys._MEIPASS)
    app_path = bundle_dir / "app.py"
    with socket.socket() as server_socket:
        server_socket.bind(("127.0.0.1", 0))
        port = server_socket.getsockname()[1]
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--global.developmentMode=false",
        "--server.address=127.0.0.1",
        f"--server.port={port}",
        "--server.headless=false",
        "--server.showEmailPrompt=false",
        "--browser.gatherUsageStats=false",
    ]
    cli.main()


if __name__ == "__main__":
    main()
