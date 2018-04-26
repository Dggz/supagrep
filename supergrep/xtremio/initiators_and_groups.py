"""Initiators and Groups Sheet (XtremIO)"""
import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import concat, map

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output, multiple_join

SHOW_INITIATORS_TMPL = textwrap.dedent("""\
    Value Required IGName (\S+)
    Value ClusterName (\S+)
    Value InitiatorName (\S+)
    Value PortType (\S+)
    Value PortAddress (\S+)
    Value InitiatorOS (\S+)

    Start
      ^\s*Initiator-Name\s+Index\s+Cluster-Name\s+Index\s+Port-Type\s+Port-Address\s+ -> InitiatorLine

    InitiatorLine
      ^\s*${InitiatorName}\s+(\S+)\s+${ClusterName}\s+(\S+)\s+${PortType}\s+${PortAddress}\s+${IGName}\s+(.+)\s+${InitiatorOS}\s*$$ -> Record
      ^\s*(\*+) -> Start
""")

SHOW_INITIATOR_GROUPS_TMPL = textwrap.dedent("""\
    Value Required IGName (\S+)
    Value ClusterName (\S+)
    Value NumOfInitiators (\d+)
    Value NumOfVols (\d+)

    Start
      ^\s*IG-Name\s+Index\s+Cluster-Name\s+Index\s+Num-of-Initiators -> GroupsLine

    GroupsLine
      ^\s*${IGName}\s+(\S+)\s+${ClusterName}\s+(\S+)\s+${NumOfInitiators}\s+${NumOfVols}\s+ -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Initiators and Groups worksheet (XtremIO)

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Initiators and Groups')

    headers = list(concat([
        get_parser_header(SHOW_INITIATORS_TMPL),
        get_parser_header(SHOW_INITIATOR_GROUPS_TMPL)[2:],
    ]))
    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_initiators_out = run_parser_over(content, SHOW_INITIATORS_TMPL)
    show_initiator_groups_out = run_parser_over(content, SHOW_INITIATOR_GROUPS_TMPL)

    common_columns = (0, 1)

    rows = multiple_join(
        common_columns, [show_initiators_out, show_initiator_groups_out])

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            set_cell_to_number(cell)
            style_value_cell(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'InitiatorsGroupsTable',
        'Initiators and Groups',
        final_col,
        final_row)
