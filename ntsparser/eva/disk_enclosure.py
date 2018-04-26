"""Disk Enclosure sheet"""
from collections import namedtuple
from typing import Any

import xmltodict

from ntsparser.eva.utils import merge_dicts
from ntsparser.formatting import build_header
from ntsparser.utils import (
    sheet_process_output, search_tag_value, ordered_jsons, flatten_dict,
    write_excel)


def process(workbook: Any, contents: list) -> None:
    """Process Controller worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Disk Enclosure'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    excel_header = [
        'Name', 'Type', 'DiskSlotType', 'Transport',
        'ProductId', 'ProductNumber', 'DiskSlot'
    ]

    RowTuple = namedtuple('RowTuple', excel_header)
    build_header(worksheet, excel_header)

    header = [
        'objectname', 'objecttype', 'diskslottype', 'transport',
        'productid', 'productnum', 'diskslot/name'
    ]

    disk_data = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)

        raw_disk_data = [flatten_dict(det_dict)
                         for det_dict in search_tag_value(doc, 'object')]
        for main_dict in raw_disk_data:
            entry = list(ordered_jsons([main_dict], ['diskslot']))
            if entry:
                main_dict['diskslot'] = merge_dicts(main_dict['diskslot'])
                main_dict = flatten_dict(main_dict)
                disk_data += ordered_jsons([main_dict], header)

    final_col, final_row = write_excel(disk_data, worksheet, RowTuple, 'A')
    sheet_process_output(
        worksheet,
        'DiskEnclosureTable',
        'DiskEnclosure',
        final_col,
        final_row)
