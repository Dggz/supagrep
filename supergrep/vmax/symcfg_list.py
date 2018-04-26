"""Symcfg_list parser"""

import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

SYMCFG_TMPL = textwrap.dedent("""\
    Value Required SymmID (\d+)
    Value Attachment (\w+)
    Value Model (\S+)
    Value McodeVersion (\d+)
    Value CacheSize (\d+)
    Value NumPhysDevices (\d+)
    Value NumSymmDevices (\d+)

    Start
      ^\s*${SymmID}\s+${Attachment}\s+${Model}\s+${McodeVersion}\s+${CacheSize}\s+${NumPhysDevices}\s+${NumSymmDevices} -> Record Start
""")


def process(workbook: Any, content: str) -> None:
    """Process symcfg_list worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'symcfg_list'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(SYMCFG_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)

    symcfg_list_out = run_parser_over(content, SYMCFG_TMPL)
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, symcfg_list_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'SymcfgListTable',
        'symcfg_list',
        final_col,
        final_row)
