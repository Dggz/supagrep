"""Main module of application"""

from logging import getLogger
from typing import Any

from supergrep.config import settings
from supergrep.parseargs import parse_args

logger = getLogger(__name__)


def show_banner() -> None:
    """Display banner with version and copyright info and log build date"""
    print(settings.title)


def main(argv: Any) -> Any:
    """Console application entry point

    :param argv: sys.argv
    :return:
    """
    show_banner()
    func, input_files, output_file = parse_args(argv[1:])
    if not output_file.endswith('.xlsx'):
        logger.error('output filename should end with .xlsx')
        return 1
    func(input_files, output_file)
