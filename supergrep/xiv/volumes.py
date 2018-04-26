"""Volumes sheet"""
import textwrap
from collections import namedtuple
from typing import Any, Iterable

import xmltodict

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output, search_tag_value, \
    ordered_jsons, flatten_dict

SYSTEM_NAME_TMPL = textwrap.dedent("""\
    Value Filldown,Required SystemName (\w+)

    Start
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+cod_list\s+-f\s+all -> BackendLines

    BackendLines
      ^\s*system_name\s+ ${SystemName} -> Start
""")


def process(workbook: Any, contents: Iterable) -> None:
    """Process Volumes Status worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Volumes'
    worksheet = workbook.get_sheet_by_name(worksheet_name)
    headers = get_parser_header(SYSTEM_NAME_TMPL)

    headers += [
        'VolumeId', 'VolumeName', 'Size', 'Size_MiB', 'Capacity',
        'Serial', 'cg_Name', 'sg_Name', 'SgSnapshotOf', 'Locked',
        'SnapshotTime', 'SnapshotTimeOnMaster', 'PoolName',
        'UsedCapacity_GB', 'LockedByPool', 'SnapshotInternalRole',
        'Mirrored', 'Compressed', 'Ratio', 'Saving', 'Online',
        'MetadataMismatch'
    ]

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)
    headers = [
        'id/@value', 'name/@value', 'size/@value', 'size_MiB/@value',
        'capacity/@value', 'serial/@value', 'cg_name/@value',
        'sg_name/@value', 'sg_snapshot_of/@value', 'locked/@value',
        'snapshot_time/@value', 'snapshot_time_on_master/@value',
        'pool_name/@value', 'used_capacity/@value',
        'locked_by_pool/@value', 'snapshot_internal_role/@value',
        'mirrored/@value', 'compressed/@value', 'ratio/@value',
        'saving/@value', 'online/@value', 'metadata_mismatch/@value'
    ]

    rows = []  # type: list
    for sys_content, content in contents:
        system_name = run_parser_over(sys_content, SYSTEM_NAME_TMPL)[0]
        volumes_content = '\n'.join(content.split('\n')[1:])
        doc = xmltodict.parse(volumes_content)
        command_details = search_tag_value(doc, 'volume')
        flat_data = [flatten_dict(data) for data in command_details]
        volumes = ordered_jsons(flat_data, headers)
        rows += [system_name + row for row in volumes]

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            if chr(col_n) != 'B':
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'VolumesTable',
        'Volumes',
        final_col,
        final_row)
