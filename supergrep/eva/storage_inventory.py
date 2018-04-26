"""Storage Inventory sheet"""
from collections import namedtuple
from typing import Any

import xmltodict

from supergrep.formatting import build_header
from supergrep.utils import (
    ordered_jsons, sheet_process_output, search_tag_value, write_excel)


def process(workbook: Any, contents: list) -> None:
    """Process Storage Inventory worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Storage Inventory'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    system_excel_header = [
        'SystemObjectName', 'SystemType', 'FirmwareVersion',
        'TotalStorageSpace', 'AvailableStorageSpace', 'UsedStorageSpace'
    ]

    SystemTuple = namedtuple('SystemTuple', system_excel_header)

    build_header(worksheet, system_excel_header)

    system_header = [
        'objectname', 'systemtype', 'firmwareversion',
        'totalstoragespace', 'availablestoragespace', 'usedstoragespace'
    ]

    system_data = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        system_data += list(
            ordered_jsons(search_tag_value(doc, 'object'), system_header))

    final_col, final_row = write_excel(system_data, worksheet, SystemTuple, 'A')
    sheet_process_output(
        worksheet,
        'SystemTable',
        'Storage Inventory',
        final_col,
        final_row)
