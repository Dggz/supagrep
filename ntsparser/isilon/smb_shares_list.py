"""SMB Shares List Summary sheet"""
import csv
from collections import namedtuple
from typing import Any

import xmltodict

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.isilon.utils import collected_data, squash
from ntsparser.utils import sheet_process_output, search_tag_value


def process(workbook: Any, contents: list) -> list:
    """Process SMB Shares List worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'SMB Shares List'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = ['Hostname', 'ShareName', 'Path']

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    rows = []  # type: Any
    bad_rows = 0
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        shares = []  # type: list
        host = component_details['hostname']
        for entry in command_details:
            shares_content = collected_data(
                entry, 'cmd', 'isi smb shares list --format?csv -a -z')
            shares = list(csv.reader(shares_content.split('\n'))
                          if shares_content else shares)
        bad_rows += len(list(filter(
            lambda x: len(x) != 2 and len(x) > 1, shares)))
        rows += [
            [host] + squash(row, 0, -1)
            for row in filter(lambda x: len(x) == 2, shares)
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
        'SMBSharesListTable',
        'SMB Shares List',
        final_col,
        final_row)

    if bad_rows:
        print('{} bad rows in smb shares csv!'.format(bad_rows))

    return [[row[0], row[1].split('/')[-1], row[1], 'SMB'] for row in rows]
