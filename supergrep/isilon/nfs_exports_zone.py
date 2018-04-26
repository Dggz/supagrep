"""NFS Exports by Zone sheet"""
from contextlib import suppress
from typing import Any, Iterable

import xmltodict
from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from supergrep.isilon.utils import (
    collected_alias_data, collected_data, process_nfs_zones)
from supergrep.utils import sheet_process_output, search_tag_value


def process(workbook: Any, contents: list) -> None:
    """Process NFS Exports by Zone worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'NFS Exports by Zone'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = [
        'hostname', 'encoding', 'map_retry', 'security_flavors',
        'write_unstable_reply', 'write_filesync_reply', 'readdirplus_prefetch',
        'block_size', 'write_transfer_size', 'id', 'description',
        'max_file_size', 'paths', 'write_unstable_action', 'zone',
        'name_max_size', 'can_set_time', 'read_transfer_multiple',
        'return_32bit_file_ids', 'write_transfer_multiple', 'all_dirs',
        'setattr_asynchronous', 'map_failure/primary_group',
        'map_failure/enabled', 'map_failure/secondary_groups',
        'map_failure/user', 'link_max', 'write_datasync_reply', 'no_truncate',
        'time_delta', 'snapshot', 'read_only', 'map_lookup_uid',
        'chown_restricted', 'write_datasync_action', 'read_transfer_size',
        'map_full', 'read_transfer_max_size', 'map_root/primary_group',
        'map_root/enabled', 'map_root/secondary_groups', 'map_root/user',
        'map_non_root/primary_group', 'map_non_root/enabled',
        'map_non_root/secondary_groups', 'map_non_root/user', 'symlinks',
        'commit_asynchronous', 'write_filesync_action', 'case_insensitive',
        'readdirplus', 'write_transfer_max_size', 'directory_transfer_size',
        'case_preserving', 'read_only_clients', 'clients'
    ]
    build_header(worksheet, headers)

    rows = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        nfs_zone, nfs_alias_zone = [], []  # type: Iterable, Iterable
        host = component_details['hostname']
        for entry in command_details:
            with suppress(TypeError):
                nfs_zone_alias_content = collected_alias_data(
                    entry, 'cmd', 'isi nfs exports list by zone **format?json')
                nfs_zone_content = collected_data(
                    entry, 'cmd', 'isi nfs exports list by zone **format?json')

                nfs_alias_zone = process_nfs_zones(
                    nfs_zone_alias_content, headers[1:]) \
                    if nfs_zone_alias_content else nfs_alias_zone
                nfs_zone = process_nfs_zones(
                    ''.join(nfs_zone_content), headers[1:]) \
                    if nfs_zone_content else nfs_zone

        nfs_zone = list(nfs_zone) + list(nfs_alias_zone)
        rows += [[host] + row for row in nfs_zone]

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(rows, 2):
        for col_n, col_value in \
                enumerate(row_tuple, ord('A')):
            cell = worksheet['{}{}'.format(column_format(col_n), row_n)]
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
        'NFSExportsByZoneTable',
        'NFS Exports by Zone',
        final_col,
        final_row)
