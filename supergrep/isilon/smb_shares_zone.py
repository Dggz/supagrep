"""SMB Shares by Zone sheet"""
from contextlib import suppress
from typing import Any, Iterable

import xmltodict
from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from supergrep.isilon.utils import (
    collected_alias_data, collected_data, process_smb_zones)
from supergrep.utils import sheet_process_output, search_tag_value


def process(workbook: Any, contents: list) -> None:
    """Process SMB Shares by Zone worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'SMB Shares by Zone'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = [
        'hostname', 'file_create_mode', 'ntfs_acl_support',
        'file_filtering_enabled', 'directory_create_mask',
        'inheritable_path_acl', 'ca_write_integrity', 'file_create_mask',
        'path', 'access_based_enumeration_root_only',
        'allow_variable_expansion', 'id', 'hide_dot_files',
        'create_permissions', 'strict_locking', 'mangle_map', 'zid',
        'file_filter_type', 'browsable', 'description', 'auto_create_directory',
        'continuously_available', 'directory_create_mode',
        'allow_delete_readonly', 'ca_timeout', 'access_based_enumeration',
        'allow_execute_always', 'csc_policy', 'permissions/permission_type',
        'permissions/trustee/id', 'permissions/permission', 'name',
        'change_notify', 'mangle_byte_start', 'impersonate_user',
        'impersonate_guest', 'oplocks', 'strict_flush', 'strict_ca_lockout'
    ]
    build_header(worksheet, headers)

    rows = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        smb_zone, smb_alias_zone = [], []  # type: Iterable, Iterable
        host = component_details['hostname']
        for entry in command_details:
            with suppress(TypeError):
                smb_zone_alias_content = collected_alias_data(
                    entry, 'cmd', 'isi smb shares list by zone **format?json')
                smb_zone_content = collected_data(
                    entry, 'cmd', 'isi smb shares list by zone **format?json')

                smb_alias_zone = process_smb_zones(
                    smb_zone_alias_content, headers[1:]) \
                    if smb_zone_alias_content else smb_zone
                smb_zone = process_smb_zones(
                    ''.join(smb_zone_content), headers[1:]) \
                    if smb_zone_content else smb_zone
        smb_zone = list(smb_zone) + list(smb_alias_zone)
        rows += [[host] + row for row in smb_zone]

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
            final_col = col_n if col_n > final_col else final_col
        final_row = row_n

    sheet_process_output(
        worksheet,
        'SMBSharesByZoneTable',
        'SMB Shares by Zone',
        final_col,
        final_row)
