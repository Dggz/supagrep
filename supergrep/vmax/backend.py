"""Perf_BACKEND parser"""
import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output


BACKEND_TMPL = textwrap.dedent("""\
    Value Filldown,Required SymmetrixID (\d+)
    Value Filldown Time (\S+)
    Value Required SymDev (.....|....)
    Value Physical (\S+\s\S+|\S+)
    Value IOSecRead (\S+)
    Value IOSecWrite (\S+)
    Value KBSecRead (\S+)
    Value KBSecWrite (\S+)
    Value PrefetchTracks (\S+)
    Value TracksUsed (\S+)

    Start
      ^\s*${SymmetrixID}_(\S+)_type_backend.txt -> BackendStart

    BackendStart
      ^\s*(\S+)\s+Dev -> BackendLines
      ^\s*DEVICE -> BackendLines
      
    BackendLines
      ^\s*${Time}\s+${SymDev}\s+\(${Physical}\s*\)\s+${IOSecRead}\s+${IOSecWrite}\s+${KBSecRead}\s+${KBSecWrite}\s+${PrefetchTracks}\s+${TracksUsed}\s*$$ -> Record
      ^\s*${SymDev}\s+\(${Physical}\s*\)\s+${IOSecRead}\s+${IOSecWrite}\s+${KBSecRead}\s+${KBSecWrite}\s+${PrefetchTracks}\s+${TracksUsed}\s*$$ -> Record
      ^\s*${Time}\s+${SymDev}\s+${Physical}\s+${IOSecRead}\s+${IOSecWrite}\s+${KBSecRead}\s+${KBSecWrite}\s+${PrefetchTracks}\s+${TracksUsed}\s*$$ -> Record
      ^\s*${SymDev}\s+${Physical}\s+${IOSecRead}\s+${IOSecWrite}\s+${KBSecRead}\s+${KBSecWrite}\s+${PrefetchTracks}\s+${TracksUsed}\s*$$ -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Perf_BACKEND worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'Perf_BACKEND'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(BACKEND_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)
    list_wwn_out = run_parser_over(content, BACKEND_TMPL)
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, list_wwn_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            if chr(col_n) not in ('C', ):
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'PerfBACKENDTable',
        'Perf_BACKEND',
        final_col,
        final_row)
