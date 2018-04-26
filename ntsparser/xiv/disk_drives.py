"""Disk Drives Sheet"""

import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import concat, map

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output, multiple_join

SYSTEM_NAME_TMPL = textwrap.dedent("""\
    Value Filldown,Required SystemName (\w+)
    Value Required ComponentId (\d+:\S+:\d+:\d+)
    Value Filldown DiskType (\S+)
    
    Start
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+cod_list\s+-f\s+all -> BackendLines
      
    BackendLines
      ^\s*system_name\s+ ${SystemName} -> System

    System
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u xiv_development\s+-p\s+conf_get path=misc.status -> StatusLine
      
    StatusLine
      ^\s*${DiskType}_disk\S* -> CompLine
    
    CompLine
      ^\s*component_id\s*=\s*"${ComponentId} -> Record
      ^\s*(\*+) -> Start
""")

DISK_TMPL = textwrap.dedent("""\
    Value Filldown,Required SystemName (\w+)
    Value Required ComponentId (\S+)
    Value DiskSize ([\d\w_]+)
    Value Status (\S+)
    Value Vendor (\S+)

    Start
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+cod_list\s+-f\s+all -> BackendLines
      
    BackendLines
      ^\s*system_name\s+ ${SystemName} -> System
      
    System
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+disk_list\s+-f\s+all -> DiskLine

    DiskLine
      ^\s*Component ID\s+Status\s+Currently Functioning\s+Capacity\s+Target Status\s+Vendor -> DiskLine
      ^\s*${ComponentId}\s+${Status}\s+(\S+)\s+${DiskSize}\s+${Vendor}\s+(.+) -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Disk Drivers worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'Disk Drives'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = list(concat([
        get_parser_header(SYSTEM_NAME_TMPL),
        get_parser_header(DISK_TMPL)[2:],
    ]))
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)

    system_drivers_out = run_parser_over(content, SYSTEM_NAME_TMPL)
    disk_drivers_out = run_parser_over(content, DISK_TMPL)

    common_columns = (0, 1)
    rows = multiple_join(
        common_columns, [system_drivers_out, disk_drivers_out])

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
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
        'DiskDrivesTable',
        'Disk Drives',
        final_col,
        final_row)
