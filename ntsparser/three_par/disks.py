"""Disks (3Par) Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output

SHOWPD_TMPL = textwrap.dedent("""\
    Value Filldown Name (\S+)
    Value Filldown Serial (\S+)
    Value Required Id (\S+)
    Value CagePos (\S+)
    Value Type (\S+)
    Value RPM (\S+)
    Value State (\S+)
    Value Total (\S+)
    Value Free (\S+)
    Value A (\S+)
    Value B (\S+)
    Value CapGB (\S+)
    
    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+(\S+\s\S+|\S+)\s+${Serial}\s+
      ^\s*(\*+) -> DisksStart

    DisksStart
      ^\s*Id\s+CagePos\s+Type\s+RPM\s+State\s+Total\s+Free\s+A\s+B\s+(Cap|Capacity)\(GB\) -> DisksLines
      
    DisksLines
      ^\s*${Id}\s+${CagePos}\s+${Type}\s+${RPM}\s+${State}\s+${Total}\s+${Free}\s+${A}\s+${B}\s+${CapGB} -> Record DisksLines
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Disks (3Par) worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Disks')

    headers = get_parser_header(SHOWPD_TMPL)
    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    headers[7], headers[8], headers[11] = 'Total(MB)', 'Free(MB)', 'Cap(GB)'
    build_header(worksheet, headers)

    show_pd_out = run_parser_over(content, SHOWPD_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_pd_out), 2):
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
        'DisksTable',
        'Disks',
        final_col,
        final_row)
