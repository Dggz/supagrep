"""Drive List sheet"""
from contextlib import suppress
# noinspection TaskProblemsInspection
from typing import Any, Iterable

import xmltodict
from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from supergrep.isilon.utils import (
    collected_data,
    process_drives)
from supergrep.utils import sheet_process_output, search_tag_value


def process(workbook: Any, contents: list) -> None:
    """Process Drive List worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Drive List'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = [
        'hostname', 'cluster', 'lnum', 'baynum', 'lnn', 'purpose_description',
        'blocks', 'serial', 'wwn', 'logical_block_length', 'ui_state',
        'physical_block_length', 'firmware/desired_firmware',
        'firmware/current_firmware', 'id', 'media_type', 'interface_type',
        'handle', 'devname', 'chassis', 'purpose', 'y_loc', 'x_loc', 'present',
        'locnstr', 'model'
    ]
    build_header(worksheet, headers)

    rows, errors = [], 0  # type: list, int
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        drive_list = []  # type: Iterable
        host = component_details['hostname']
        for entry in command_details:
            with suppress(TypeError):
                drives_content = collected_data(
                    entry, 'cmd',
                    'isi_for_array isi devices list --format?json')
                drive_list, local_errors = process_drives(
                    drives_content,
                    headers[2:]) if drives_content else [drive_list, 0]
                errors += local_errors
        rows += [[host] + row for row in drive_list]

    if errors != 0:
        print('{} bad jsons found in {}, '
              'some data will not be found in the output!'
              .format(errors, worksheet_name))
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
        'DriveListTable',
        'Drive List',
        final_col,
        final_row)
