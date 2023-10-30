import logging

from rich import print
import rich_click as click
from rich.console import Console
from rich.table import Table

from qasap._version import __version__
from qasap.dashboard import start_dashboard

click.rich_click.USE_RICH_MARKUP = True


@click.group()
def qasap():
    logging.basicConfig(level=logging.INFO)
    pass


@qasap.command()
def version():
    """Print version and exit."""
    print(f"Version: {__version__}")


@qasap.command()
@click.option(
    "-d",
    "--develop",
    type=bool,
    is_flag=True,
    help="Enables development mode on the dashboard, allowing to update the python package.",
    default=False,
)
def dashboard(develop: bool):
    """
    Start dashboard.
    """
    start_dashboard(develop=develop)


if __name__ == "__main__":
    qasap()
