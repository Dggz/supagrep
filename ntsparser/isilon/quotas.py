"""Quotas sheet"""
import json
from collections import namedtuple
from typing import Any, Iterable

import xmltodict

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.isilon.utils import collected_data, quotas_json
from ntsparser.utils import sheet_process_output, search_tag_value


def process(workbook: Any, contents: list) -> None:
    """Process Quotas worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Quotas'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = [
        'Hostname', 'Path', 'Type', 'AppliesTo', 'Hard', 'Soft', 'Advisory',
        'PhysicalUsed', 'LogicalUsed', 'InodesCount'
    ]

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    rows = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        quotas = []  # type: Iterable
        host = component_details['hostname']
        for entry in command_details:
            quotas_content = collected_data(
                entry, 'cmd', 'isi quota quotas list --format?json')
            quotas = quotas_json(json.loads(quotas_content)) \
                if quotas_content else quotas
        rows += [[host] + row for row in quotas]

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
        'QuotasTable',
        'Quotas',
        final_col,
        final_row)
