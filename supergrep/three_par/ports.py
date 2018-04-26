"""Ports (3Par) Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

SHOWPORT_TMPL = textwrap.dedent("""\
    Value Filldown Name (\S+)
    Value Filldown Serial (\S+)
    Value Required NSP (\S+)
    Value Mode (\S+)
    Value State (\S+)
    Value NodeWWN (\S+)
    Value PortWWNOrHWAddr (\S+)
    Value Type (\S+)
    Value Protocol (\S+)
    Value Label (\S+)
    Value Partner (\S+)
    Value FailoverState (\S+)

    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+(\S+\s\S+|\S+)\s+${Serial}\s+
      ^\s*(\*+) -> PortStart

    PortStart
      ^\s*N:S:P\s+Mode\s+State\s+----Node_WWN----\s+-Port_WWN/HW_Addr-\s+Type\s+Protocol\s+Label\s+Partner\s+FailoverState -> PortLines

    PortLines
      ^\s*${NSP}\s+${Mode}\s+${State}\s+${NodeWWN}\s+${PortWWNOrHWAddr}\s+${Type}\s+${Protocol}\s+${Label}\s+${Partner}\s+${FailoverState} -> Record PortLines
      ^\s*(-+) -> Start
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Ports (3Par) worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Ports')

    headers = get_parser_header(SHOWPORT_TMPL)
    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_ports_out = run_parser_over(content, SHOWPORT_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_ports_out), 2):
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
        'PortsTable',
        'Ports',
        final_col,
        final_row)
