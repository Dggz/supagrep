"""list_WWN parser"""

import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

LSTWWN_TMPL = textwrap.dedent("""\
    Value Filldown,Required SymmetrixID (\d+)
    Value Required Sym (.....|....)
    Value Physical (\S+\s\S+|\S+)
    Value Config (.+)
    Value Attr (\(.\)|\s*)
    Value WWN (\d+)

    Start
      ^\s*Symmetrix ID:\s*${SymmetrixID} 
      ^\s*(-+)\s+(-+) -> SymStart

    SymStart
      ^\s*${Sym}\s+${Physical}\s+${Config}\s+\s+${Attr}\s+${WWN}\s*$$ -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process list_WWN worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'list_WWN'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(LSTWWN_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)
    list_wwn_out = run_parser_over(content, LSTWWN_TMPL)
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, list_wwn_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            if chr(col_n) not in ('B', 'F'):
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'ListWWNTable',
        'list_WWN',
        final_col,
        final_row)
