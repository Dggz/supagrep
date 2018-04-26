"""Access_view parser"""

import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

ACSVW_TMPL = textwrap.dedent("""\
    Value Filldown,Required SymmetrixID (\d+)
    Value Required MaskingViewName (\S+)
    Value InitiatorGroup (\S+) 
    Value PortGroup (\S+)
    Value StorageGroup (\S+)

    Start
      ^\s*Symmetrix ID\s+:\s*${SymmetrixID} 
      ^\s*(-+)\s+(-+)\s+(-+)\s+(-+) -> ACSVWLine

    ACSVWLine
      ^\s*${MaskingViewName}\s+${InitiatorGroup}\s+${PortGroup}\s+${StorageGroup} -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process access_view worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'access_view'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(ACSVW_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)

    access_view_out = run_parser_over(content, ACSVW_TMPL)
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, access_view_out), 2):
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
        'AccessViewTable',
        'access_view',
        final_col,
        final_row)
