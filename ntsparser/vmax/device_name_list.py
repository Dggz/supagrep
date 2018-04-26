"""device_name_list parser"""

import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output

DEVNM_TMPL = textwrap.dedent("""\
    Value Filldown,Required SymmetrixID (\d+)
    Value Required Sym (.....|....)
    Value Config (\w+)
    Value Attr ((.+)|(\s+))
    Value DeviceName (\w+)

    Start
      ^\s*Symmetrix ID:\s*${SymmetrixID} 
      ^\s*(-+)\s+(-+) -> SymStart

    SymStart
      ^\s*${Sym}\s+${Config}\s+${Attr}\s+${DeviceName}\s*$$ -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process device_name_list worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'device_name_list'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(DEVNM_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)
    device_name_list_out = run_parser_over(content, DEVNM_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, device_name_list_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            if chr(col_n) != 'B':
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'DevNameListTable',
        'device_name_list',
        final_col,
        final_row)
