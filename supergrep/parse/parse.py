"""Celerra command entry point"""

from openpyxl import Workbook

from supergrep.parsing import load_raw_content
from supergrep.utils import get_relevant_content

__all__ = ('main',)


def main(input_files: str, output_file: str) -> None:
    """Celerra Entry point

    :param input_files:
    :param output_file:
    """
    workbook = Workbook()

    raw_content_patterns = (
        '*cmd_outputs/hostname',
    )

    raw_content = load_raw_content(tuple(input_files), raw_content_patterns)

    import ipdb; ipdb.set_trace()
    system_details_content = get_relevant_content(
        raw_content, (raw_content_patterns[0], ), '*' * 20 + '\n')


    workbook.save(output_file)
