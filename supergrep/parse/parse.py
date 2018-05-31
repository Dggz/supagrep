"""Celerra command entry point"""
import json
import xmltodict

from typing import Any

from openpyxl import Workbook

from supergrep.parsing import load_raw_content, run_parser_over
from supergrep.utils import get_relevant_content

__all__ = ('main',)


def main(templates: Any, input_files: str, output_file: str) -> None:
    """Celerra Entry point

    :param templates:
    :param input_files:
    :param output_file:
    """
    workbook = Workbook()

    raw_content_patterns = (
        '*cmd_outputs/hostname',
    )
    # import ipdb; ipdb.set_trace()
    with open(templates, 'r') as temp:
        temp_data = xmltodict.parse(''.join(temp.readlines()))['sgr']
    import ipdb; ipdb.set_trace()
    with open(input_files[0], 'r') as inp:
        data = ''.join(inp.readlines())


    extracted_data = run_parser_over(data, temp_data['template'])
    fname = temp_data['fname']
    sheet = temp_data['sheet']

    import ipdb; ipdb.set_trace()
    # raw_content = load_raw_content(tuple(input_files), raw_content_patterns)

    # import ipdb; ipdb.set_trace()
    # system_details_content = get_relevant_content(
    #     raw_content, (raw_content_patterns[0], ), '*' * 20 + '\n')

    workbook.save(output_file)
