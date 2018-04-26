"""NFS Exports List Summary sheet"""
import csv
from collections import namedtuple
from typing import Any

import xmltodict

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.isilon.utils import collected_data, squash
from ntsparser.utils import sheet_process_output, search_tag_value


def process(workbook: Any, contents: list) -> list:
    """Process NFS Exports List worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'NFS Exports List'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = ['Hostname', 'ID', 'Zone', 'Paths', 'Description']

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    rows = []  # type: Any
    bad_rows = 0
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        exports = []  # type: list
        host = component_details['hostname']
        for entry in command_details:
            exports_content = collected_data(
                entry, 'cmd', 'isi nfs exports list --format?csv -a -z')
            exports = list(csv.reader(exports_content.split('\n'))
                           if exports_content else exports)
        bad_rows += len(list(filter(
            lambda x: len(x) != 4 and len(x) > 1, exports)))
        rows += [
            [host] + squash(row, 2, -1)
            for row in filter(lambda x: len(x) == 4, exports)
        ]

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
        'NFSExportsListTable',
        'NFS Exports List',
        final_col,
        final_row)

    if bad_rows:
        print('{} bad rows in nfs exports csv!'.format(bad_rows))

    return [[row[0], row[3].split('/')[-1], row[3], 'NFS']  for row in rows]
