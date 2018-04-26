"""Storage Controllers Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import concat, map

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output, multiple_join

STORAGE_CONTROLLERS_TMPL = textwrap.dedent("""\
    Value Required,Filldown SystemName (\w+)
    Value ConsumedCapacity (\d+)
    Value MachineModel (\d+)
    Value MachineSerialNumber (\d+)
    Value MachineType (\d+)
    Value SystemId (\d+)
   
    Start
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+cod_list\s+-f\s+all -> BackendLines

    BackendLines
      ^\s*consumed_capacity\s+ ${ConsumedCapacity}
      ^\s*machine_model\s+ ${MachineModel}
      ^\s*machine_serial_number\s+ ${MachineSerialNumber}
      ^\s*machine_type\s+ ${MachineType} 
      ^\s*system_id\s+ ${SystemId}
      ^\s*system_name\s+ ${SystemName} -> Start
""")

STORAGE_VERSION_TMPL = textwrap.dedent("""\
    Value Required SystemName (\w+)
    Value ArrayFirmware (\S+)

    Start
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+cod_list\s+-f\s+all -> BackendLines

    BackendLines
      ^\s*system_name\s+ ${SystemName} -> ControllersStart

    ControllersStart
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+version_get -> VersionStart

    VersionStart
      ^\s*---------\s+ -> VersionLine

    VersionLine
      ^\s*${ArrayFirmware}\s*$$ -> Start
""")

STORAGE_CAPACITY_TMPL = textwrap.dedent("""\
    Value Required SystemName (\w+)
    Value SystemState (\w+)
    Value RedundancyStatus (\w+\s+\w+)
    Value SystemCapacityGB (\d+)
    Value RawCapacityGB (\d+)
    Value UsedCapacityGB (\d+)
    Value ConsumedCapacityGB (\d+)
    Value FreeCapacityGB (\d+)
    
    Start
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+cod_list\s+-f\s+all -> BackendLines

    BackendLines
      ^\s*system_name\s+ ${SystemName} -> StatusLine

    StatusLine
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+conf_get\s+path=system -> SystemLine 
       
    SystemLine
      ^\s*capacity: -> CapacityLine

    CapacityLine
      ^\s*soft\s*=\s*"${SystemCapacityGB}
      ^\s*free_soft\s*=\s*"${FreeCapacityGB}
      ^\s*raw_capacity\s*=\s*"${RawCapacityGB}
      ^\s*used_capacity_soft\s*=\s*"${UsedCapacityGB}
      ^\s*consumed_capacity\s*=\s*"${ConsumedCapacityGB} -> SystemName 

    SystemName
      ^\s*system_state: -> System
    
    System
      ^\s*system_state\s*=\s*"${SystemState}
      ^\s*redundancy_status\s*=\s*"${RedundancyStatus} ->  Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Storage Controllers worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Storage Controllers')

    headers = list(concat([
        get_parser_header(STORAGE_CONTROLLERS_TMPL),
        get_parser_header(STORAGE_VERSION_TMPL)[1:],
        get_parser_header(STORAGE_CAPACITY_TMPL)[1:]
    ]))

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)
    storage_controllers_out = run_parser_over(content, STORAGE_CONTROLLERS_TMPL)
    storage_version_out = run_parser_over(content, STORAGE_VERSION_TMPL)
    storage_capacity_out = run_parser_over(content, STORAGE_CAPACITY_TMPL)

    common_columns = (0,)
    rows = multiple_join(
        common_columns,
        [storage_controllers_out,
         storage_version_out,
         storage_capacity_out])

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
        'StorageControllersTable',
        'Storage Controllers',
        final_col,
        final_row)
