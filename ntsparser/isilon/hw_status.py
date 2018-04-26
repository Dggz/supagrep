"""HW Status sheet"""
from collections import namedtuple
from typing import Any, Iterable

import xmltodict

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.isilon.utils import collected_data, hw_split
from ntsparser.utils import sheet_process_output, search_tag_value


def process(workbook: Any, contents: list) -> None:
    """Process HW Status worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'HW Status'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = ['Hostname', 'Cluster', 'Component', 'Status']

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    rows = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        hw_statuses = []  # type: Iterable
        host = component_details['hostname']
        for entry in command_details:
            status_content = collected_data(
                entry, 'cmd', 'isi_for_array isi_hw_status')
            hw_statuses = hw_split(status_content) \
                if status_content else hw_statuses
        rows += [[host] + row for row in hw_statuses]

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
        'HWStatusTable',
        'HW Status',
        final_col,
        final_row)
