"""Disks sheet"""
from collections import namedtuple
from typing import Any

import xmltodict

from supergrep.formatting import build_header
from supergrep.utils import (
    flatten_dict, ordered_jsons, sheet_process_output, search_tag_value,
    write_excel)


def process(workbook: Any, contents: list) -> None:
    """Process Disks worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Disks'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    disks_excel_header = [
        'DiskName', 'OperationalState', 'EnclosureDiskBays', 'DiskBayNumber',
        'ShelfNumber', 'DiskGroupName', 'DiskType', 'ModelNumber',
        'FormattedCapacity', 'Occupancy'
    ]

    HostTuple = namedtuple('HostTuple', disks_excel_header)

    build_header(worksheet, disks_excel_header)

    disks_header = [
        'diskname', 'operationalstate', 'EnclosureDiskBays', 'diskbaynumber',
        'shelfnumber', 'diskgroupname', 'disktype', 'modelnumber',
        'formattedcapacity', 'occupancy'
    ]

    disk_data = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)

        raw_disk_data = [flatten_dict(det_dict)
                         for det_dict in search_tag_value(doc, 'object')]
        disk_data += ordered_jsons(raw_disk_data, disks_header)

    final_col, final_row = write_excel(disk_data, worksheet, HostTuple, 'A')
    sheet_process_output(
        worksheet,
        'DisksTable',
        'Disks',
        final_col,
        final_row)
