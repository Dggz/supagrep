"""Pools Sheet"""

import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

POOLS_TMPL = textwrap.dedent("""\
    Value Filldown,Required SystemName (\w+)
    Value Required PoolName (\S+)
    Value HardSizeGB (\d+)
    Value HardVolsGB (\d+)
    Value HardEmptyGB (\d+)
    Value SoftVolsGB (\d+)
    Value SoftEmptyGB (\d+)
    Value SnapSizeGB (\d+)
    Value HardSnapsGB (\d+)
    Value PerfClassName (\(.\)|\s+)
    Value Compress (\S+)
    
    Start
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+cod_list\s+-f\s+all -> BackendLines

    BackendLines
      ^\s*system_name\s+ ${SystemName} -> System

    System
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+pool_list\s+-f\s+all -> PoolsStart
      
    PoolsStart
      ^\s*Name\s+Size (GB)\s+Soft Vols (GB)\s+Snap Size (GB)\s+Soft Empty (GB)\s+Hard Size (GB)\s+Hard Vols (GB)\s+Locked\s+Hard Snaps (GB)\s+Hard Empty (GB)\s+Perf Class Name\s+Domain\s+Create Compressed Volumes
      ^\s*${PoolName}\s+${HardSizeGB}\s+${SoftVolsGB}\s+${SnapSizeGB}\s+${SoftEmptyGB}\s+(\S+)\s+${HardVolsGB}\s+(\S+)\s+${HardSnapsGB}\s+${HardEmptyGB}\s+${PerfClassName}\s+${Compress} -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Pools worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'Pools'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(POOLS_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)

    pools_out = run_parser_over(content, POOLS_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, pools_out), 2):
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
        'PoolsTable',
        'Pools',
        final_col,
        final_row)
