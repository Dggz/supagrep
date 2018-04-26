"""Data Protection Groups Sheet (XtremIO)"""
import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

SHOW_DATA_PROTECTION_TMPL = textwrap.dedent("""\
    Value Name (\S+)
    Value ClusterName (\S+)
    Value NumSSDs (\d+)
    Value UsefulSSDSPace (\S+)
    Value UserSpace (\S+)
    Value UserSpaceInUse (\S+)
    Value BrickName (\S+)

    Start
      ^\s*Name\s+Index\s+Cluster-Name\s+Index\s+State\s+Num-of-SSDs -> DataProtectionLine

    DataProtectionLine
      ^\s*${Name}\s+(\S+)\s+${ClusterName}\s+(\S+)\s+(\S+)\s+${NumSSDs}\s+${UsefulSSDSPace}\s+${UserSpace}\s+${UserSpaceInUse}\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+${BrickName}\s+(.+) -> Record
""")


def process(workbook: Any, content: str) -> None:
    """Process Data Protection Groups worksheet (XtremIO)

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Data Protection Groups')

    headers = get_parser_header(SHOW_DATA_PROTECTION_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_targets_out = run_parser_over(content, SHOW_DATA_PROTECTION_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_targets_out), 2):
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
        'DataProtectionGroupsTable',
        'Data Protection Groups',
        final_col,
        final_row)
