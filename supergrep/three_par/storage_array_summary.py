"""Storage Array Summary (3Par) Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

SHOWSYS_TMPL = textwrap.dedent("""\
    Value Name (\S+)
    Value Model (\S+\s\S+|\S+)
    Value Serial (\S+)
    Value Nodes (\S+)
    Value TotalCapMB (\S+)
    Value AllocCapMB (\S+)
    Value FreeCapMB (\S+)
    Value FailedCapMB (\S+)
    Value Location (.+)
    Value ReleaseVersion (.+)
    
    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+${Model}\s+${Serial}\s+${Nodes}\s+(\S+)\s+${TotalCapMB}\s+${AllocCapMB}\s+${FreeCapMB}\s+${FailedCapMB}
      ^\s*(\*+) -> SysDLine

    SysDLine
      ^\s*Location\s*:\s*${Location}
      ^\s*(\*+) -> ReleaseLine

    ReleaseLine
      ^\s*Release version\s+${ReleaseVersion} -> Record Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Storage Array Summary (3Par) worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Storage Array Summary')

    headers = get_parser_header(SHOWSYS_TMPL)
    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    showsys_out = run_parser_over(content, SHOWSYS_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, showsys_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(column_format(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'StorageArraySummaryTable',
        'Storage Array Summary',
        final_col,
        final_row)
