"""Hosts sheet"""
from collections import namedtuple
from typing import Any

import xmltodict

from supergrep.formatting import build_header
from supergrep.utils import (
    flatten_dict, ordered_jsons, sheet_process_output, search_tag_value,
    write_excel)


def process(workbook: Any, contents: list) -> None:
    """Process Hosts worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Hosts'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    host_excel_header = [
        'HostName', 'OperationalState', 'OSMode', 'HostType', 'VirtualDiskName'
    ]

    HostTuple = namedtuple('HostTuple', host_excel_header)

    build_header(worksheet, host_excel_header)

    host_header = [
        'hostname', 'operationalstate', 'osmode', 'hosttype',
        'presentation/virtualdiskname'
    ]

    disk_data = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)

        raw_disk_data = [flatten_dict(det_dict)
                         for det_dict in search_tag_value(doc, 'object')]
        disk_data += ordered_jsons(raw_disk_data, host_header)

    final_col, final_row = write_excel(disk_data, worksheet, HostTuple, 'A')
    sheet_process_output(
        worksheet,
        'HostsTable',
        'Hosts',
        final_col,
        final_row)
