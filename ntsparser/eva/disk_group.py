"""Disk Group sheet"""
from collections import namedtuple
from typing import Any

import xmltodict

from ntsparser.formatting import build_header
from ntsparser.utils import (
    sheet_process_output, search_tag_value, ordered_jsons, write_excel)


def process(workbook: Any, contents: list) -> None:
    """Process Disk Group worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Disk Group'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    disk_excel_header = [
        'DiskGroupName', 'TotalDisks', 'DiskDriveType',
        'RequestedSparePolicy', 'CurrentSparePolicy', 'TotalStorageSpaceGB',
        'UsedStorageSpaceGB'
    ]
    object_excel_header = ['ObjectName', 'TotalUngroupedDisks', 'XMLCapacityGB']

    DiskTuple = namedtuple('DiskTuple', disk_excel_header)
    ObjectTuple = namedtuple('ObjectTuple', object_excel_header)

    build_header(worksheet, disk_excel_header)
    build_header(worksheet, object_excel_header, 'I')

    disk_header = [
        'diskgroupname', 'totaldisks', 'diskdrivetype',
        'requestedsparepolicy', 'currentsparepolicy', 'totalstoragespacegb',
        'usedstoragespacegb'
    ]
    object_header = ['objectname', 'totalungroupeddisks', 'xmlcapacitygb']

    disk_data, object_data = [], []  # type: list, list
    for content in contents:
        doc = xmltodict.parse(content)
        disk_data += list(
            ordered_jsons(search_tag_value(doc, 'object'), disk_header))

        object_data += list(
            ordered_jsons(search_tag_value(doc, 'object'), object_header))

    final_col, final_row = write_excel(disk_data, worksheet, DiskTuple, 'A')
    sheet_process_output(
        worksheet,
        'DiskTable',
        'Disk Group',
        final_col,
        final_row)

    final_col, final_row = write_excel(object_data, worksheet, ObjectTuple, 'I')
    sheet_process_output(
        worksheet,
        'ObjectTable',
        'Disk Group',
        final_col,
        final_row,
        start_col=ord('I'))
