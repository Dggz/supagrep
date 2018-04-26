"""Sync Policies sheet"""
import json
from collections import namedtuple
from typing import Any, Iterable

import xmltodict

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.isilon.utils import collected_data
from supergrep.utils import sheet_process_output, search_tag_value, \
    ordered_jsons


def process(workbook: Any, contents: list) -> None:
    """Process Sync Policies worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Sync Policies'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = [
        'hostname', 'name', 'schedule', 'source_root_path', 'enabled',
        'target_path', 'last_success', 'action', 'id', 'target_host'
    ]

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    rows = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        sync_policies = []  # type: Iterable
        host = component_details['hostname']
        for entry in command_details:
            policies_content = collected_data(
                entry, 'cmd', 'isi sync policies list --format?json')
            sync_policies = ordered_jsons(
                json.loads(policies_content), headers[1:]) \
                if policies_content else sync_policies
        rows += [[host] + row for row in sync_policies]

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
        'SyncPoliciesTable',
        'Sync Policies',
        final_col,
        final_row)
