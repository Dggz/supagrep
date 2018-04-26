"""Perf_REQUESTS parser"""

import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output

REQUESTS_TMPL = textwrap.dedent("""\
    Value Filldown,Required SymmetrixID (\d+)
    Value Filldown Time (\S+)
    Value Required SymDev (.....|....)
    Value Physical (\S+\s\S+|\S+)
    Value IOSecRead (\S+)
    Value IOSecWrite (\S+)
    Value KBSecRead (\S+)
    Value KBSecWrite (\S+)
    Value PercentHitsRead (\S+)
    Value PercentHitsWrite (\S+)
    Value PercentSeqRead (\S+)
    Value PercentSeqWrite (\S+)
    Value NumWPTracks (\S+)

    Start
      ^\s*${SymmetrixID}_(\S+)_type_requests.txt -> RequestsStart

    RequestsStart
      ^\s*(\S+)\s+Dev -> RequestsLines
      ^\s*DEVICE -> RequestsLines

    RequestsLines
      ^\s*Total\s+ -> Start
      ^\s*${Time}\s+${SymDev}\s+\(${Physical}\s*\)\s+${IOSecRead}\s+${IOSecWrite}\s+${KBSecRead}\s+${KBSecWrite}\s+${PercentHitsRead}\s+${PercentHitsWrite}\s+${PercentSeqRead}\s+${PercentSeqWrite}\s+${NumWPTracks}\s*$$ -> Record
      ^\s*${SymDev}\s+\(${Physical}\s*\)\s+${IOSecRead}\s+${IOSecWrite}\s+${KBSecRead}\s+${KBSecWrite}\s+${PercentHitsRead}\s+${PercentHitsWrite}\s+${PercentSeqRead}\s+${PercentSeqWrite}\s+${NumWPTracks}\s*$$ -> Record
      ^\s*${Time}\s+${SymDev}\s+${Physical}\s+${IOSecRead}\s+${IOSecWrite}\s+${KBSecRead}\s+${KBSecWrite}\s+${PercentHitsRead}\s+${PercentHitsWrite}\s+${PercentSeqRead}\s+${PercentSeqWrite}\s+${NumWPTracks}\s*$$ -> Record
      ^\s*${SymDev}\s+${Physical}\s+${IOSecRead}\s+${IOSecWrite}\s+${KBSecRead}\s+${KBSecWrite}\s+${PercentHitsRead}\s+${PercentHitsWrite}\s+${PercentSeqRead}\s+${PercentSeqWrite}\s+${NumWPTracks}\s*$$ -> Record
""")


def process(workbook: Any, content: str) -> None:
    """Process Perf_REQUESTS worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'Perf_REQUESTS'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(REQUESTS_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)
    list_wwn_out = run_parser_over(content, REQUESTS_TMPL)
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, list_wwn_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            if chr(col_n) not in ('C',):
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'PerfREQUESTSTable',
        'Perf_REQUESTS',
        final_col,
        final_row)
