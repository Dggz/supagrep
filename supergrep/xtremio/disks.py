"""Disks Sheet (XtremIO)"""
import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map
from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

SHOW_SSDS_TMPL = textwrap.dedent("""\
    Value Required ClusterName (\S+)
    Value BrickName (\S+)
    Value Slot (\S+)
    Value Product (\S+)
    Value PartNumber (\S+)
    Value SSDSize (\S+)
    Value DPGName (\S+)

    Start
      ^\s*Name\s+Index\s+Cluster-Name\s+Index\s+Brick-Name\s+Index\s+Slot #\s+Product-Model -> DisksLine

    DisksLine
      ^\s*Name\s+Index\s+Cluster-Name\s+Index\s+Brick-Name\s+Index\s+Slot #\s+Product-Model -> DisksLine
      ^\s*(\S+)\s+(\S+)\s+${ClusterName}\s+(\S+)\s+${BrickName}\s+(\S+)\s+${Slot}\s+${Product}\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+${PartNumber}\s+(\S+)\s+${SSDSize}\s+${DPGName}\s+(.+) -> Record 
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Disks worksheet (XtremIO)

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Disks')

    headers = get_parser_header(SHOW_SSDS_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_ssds_out = run_parser_over(content, SHOW_SSDS_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_ssds_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            if isinstance(col_value, str):
                cell.value = str.strip(col_value)
            else:
                cell.alignment = Alignment(wrapText=True)
                cell.value = '\n'.join(col_value)
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
