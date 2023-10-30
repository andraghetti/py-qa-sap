"""QASAP dashboard."""
import os
from pathlib import Path
import sys
import pkg_resources

from streamlit.web import cli as stcli

PACKAGE_PATH = os.path.dirname(pkg_resources.resource_filename(__name__, ""))
STREAMLIT_PORT = 8501
STREAMLIT_ENTRYPOINT = Path(PACKAGE_PATH) / "qasap" / "streamlit" / "Main.py"


def make_streamlit_config() -> dict:
    return {
        "server.headless": os.getenv("STREAMLIT_SERVER_HEADLESS", "false"),
        "server.port": os.getenv("STREAMLIT_SERVER_PORT", str(STREAMLIT_PORT)),
        "server.enableStaticServing": "true",
        "theme.base": "dark",
        "theme.primaryColor": "#4a5ea8",
        "theme.backgroundColor": "#181730",
        "theme.secondaryBackgroundColor": "#1b1b38",
        "theme.textColor": "#ffffff",
        "server.address": os.getenv("STREAMLIT_SERVER_ADDRESS", "127.0.0.1"),
        "browser.gatherUsageStats": "false",
    }


def start_dashboard(develop: bool):
    config = make_streamlit_config()
    st_args = [arg for key, value in config.items() for arg in (f"--{key}", value)]
    sys.argv = ["streamlit", "run", str(STREAMLIT_ENTRYPOINT), *st_args]

    if develop:
        os.environ["PYTHONPATH"] = f'{os.environ.get("PYTHONPATH", "")}:{PACKAGE_PATH}'
    sys.exit(stcli.main())
