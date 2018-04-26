"""Disk Groups Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from ntsparser.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output

DISK_GROUPS_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required ID (.+)
    Value StorageProfiles (.+)
    Value RaidType (.+)
    Value LogicalCapacity (.+)
    Value NumSpindles (.+)
    Value NumLuns (.+)
    Value NumDiskVolumes (.+)
    Value SpindleType (.+)
    Value Bus (.+)
    Value VirtuallyProvisioned (.+)
    Value RawCapacity (.+)
    Value UsedCapacity (.+)
    Value FreeCapacity (.+)
    Value PercentFull (.+)
    Value Hidden (.+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> NASStart

    NASStart
      ^\s*Output from:\s+/nas/bin/nas_storage\s+-info\s+-all\s+-o\s+sync=no -> DiskStart

    DiskStart
      ^\s*Disk Groups -> DiskLines
      
    DiskLines
      ^\s*id\s*=\s*${ID}
      ^\s*storage profiles\s*=\s*${StorageProfiles}
      ^\s*raid_type\s*=\s*${RaidType}
      ^\s*logical_capacity\s*=\s*${LogicalCapacity}
      ^\s*num_spindles\s*=\s*${NumSpindles}
      ^\s*num_luns\s*=\s*${NumLuns}
      ^\s*num_disk_volumes\s*=\s*${NumDiskVolumes}
      ^\s*spindle_type\s*=\s*${SpindleType}
      ^\s*bus\s*=\s*${Bus}
      ^\s*virtually_provisioned\s*=\s*${VirtuallyProvisioned}
      ^\s*raw_capacity\s*=\s*${RawCapacity}
      ^\s*used_capacity\s*=\s*${UsedCapacity}
      ^\s*free_capacity\s*=\s*${FreeCapacity}
      ^\s*percent_full_threshold\s*=\s*${PercentFull}
      ^\s*hidden\s*=\s*${Hidden} -> Record
      ^\s*Spindles\s*$$ -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Disk Groups worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Disk Groups')

    headers = get_parser_header(DISK_GROUPS_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    disk_groups_out = run_parser_over(content, DISK_GROUPS_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, disk_groups_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(column_format(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            if chr(col_n) not in ('B', ):
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'DiskGroupsTable',
        'Disk Groups',
        final_col,
        final_row)
