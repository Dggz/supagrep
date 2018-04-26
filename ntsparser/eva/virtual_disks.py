"""Virtual Disks sheet"""
from collections import namedtuple
from typing import Any

import xmltodict

from ntsparser.formatting import build_header
from ntsparser.utils import (
    flatten_dict, ordered_jsons, sheet_process_output, search_tag_value,
    write_excel)


def process(workbook: Any, contents: list) -> None:
    """Process Virtual Disks worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Virtual Disks'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    family_excel_header = ['FamilyName', 'OperationalState', 'TotalSnapshots']
    host_excel_header = [
        'FamilyName', 'AllocatedCapacity', 'HostName', 'HostOSMode'
    ]

    FamilyTuple = namedtuple('FamilyTuple', family_excel_header)
    HostTuple = namedtuple('HostTuple', host_excel_header)

    build_header(worksheet, family_excel_header)
    build_header(worksheet, host_excel_header, 'E')

    family_header = [
        'familyname', 'operationalstate', 'totalsnapshots'
    ]
    host_header = [
        'familyname', 'allocatedcapacity', 'presentation/hostname',
        'presentation/hostosmode'
    ]

    family_data, host_data = [], []  # type: list, list
    for content in contents:
        doc = xmltodict.parse(content)

        family_data += list(
            ordered_jsons(search_tag_value(doc, 'object'), family_header))

        raw_host_data = [flatten_dict(det_dict)
                         for det_dict in search_tag_value(doc, 'object')]
        host_data += ordered_jsons(raw_host_data, host_header)

    final_col, final_row = write_excel(family_data, worksheet, FamilyTuple, 'A')
    sheet_process_output(
        worksheet,
        'FamilyTable',
        'Virtual Disks',
        final_col,
        final_row)

    final_col, final_row = write_excel(host_data, worksheet, HostTuple, 'E')
    sheet_process_output(
        worksheet,
        'HostTable',
        'Virtual Disks',
        final_col,
        final_row,
        start_col=ord('E'))
