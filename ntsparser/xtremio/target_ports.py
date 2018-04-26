"""Target Ports Sheet (XtremIO)"""
import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map
from openpyxl.styles import Alignment

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output

SHOW_TARGETS_TMPL = textwrap.dedent("""\
    Value Name (\S+)
    Value ClusterName (\S+)
    Value StorageControllerName (\S+)
    Value PortSpeed (\S+)
    Value PortType (\S+)
    Value PortAddress (\S+)
    Value PortState (\S+)

    Start
      ^\s*Name\s+Index\s+Cluster-Name\s+Index\s+Port-Type\s+Port-Address -> TargetLine

    TargetLine
      ^\s*Name\s+Index\s+Cluster-Name\s+Index\s+Port-Type\s+Port-Address -> TargetLine
      ^\s*${Name}\s+(\S+)\s+${ClusterName}\s+(\S+)\s+${PortType}\s+${PortAddress}\s+\s+(\s{17}\s+)\s+${PortSpeed}\s+${PortState}\s+(\S+)\s+${StorageControllerName}\s+(.+) -> Record
      ^\s*${Name}\s+(\S+)\s+${ClusterName}\s+(\S+)\s+${PortType}\s+${PortAddress}\s+(\S+)\s+${PortSpeed}\s+${PortState}\s+(\S+)\s+${StorageControllerName}\s+(.+) -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Target Ports worksheet (XtremIO)

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Target Ports')

    headers = get_parser_header(SHOW_TARGETS_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_targets_out = run_parser_over(content, SHOW_TARGETS_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_targets_out), 2):
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
        'TargetPortsTable',
        'Target Ports',
        final_col,
        final_row)
