"""Nodes (3Par) Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output


SHOWNODE_TMPL = textwrap.dedent("""\
    Value Filldown Name (\S+)
    Value Filldown Serial (\S+)
    Value Required Node (\S+)
    Value NodeName (\S+)
    Value State (\S+)
    Value Master (\S+)
    Value InCluster (\S+)
    Value Service_LED (\S+)
    Value LED (\S+)
    Value ControlMem (\S+)
    Value DataMem (\S+)
    Value Available (\S+)

    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+(\S+\s\S+|\S+)\s+${Serial}\s+
      ^\s*(\*+) -> NodeStart

    NodeStart
      ^\s*Node --Name--- -State- Master InCluster -Service_LED- ---LED--- -> NodeLines

    NodeLines
      ^\s*${Node}\s+${NodeName}\s+${State}\s+${Master}\s+${InCluster}\s+${Service_LED}\s+${LED}\s+${ControlMem}\s+${DataMem}\s+${Available} -> Record NodeLines
      ^\s*(-+) -> Start
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Nodes (3Par) worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Nodes')

    headers = get_parser_header(SHOWNODE_TMPL)
    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_nodes_out = run_parser_over(content, SHOWNODE_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_nodes_out), 2):
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
        'NodesTable',
        'Nodes',
        final_col,
        final_row)
