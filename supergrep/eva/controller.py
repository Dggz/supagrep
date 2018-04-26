"""Controller sheet"""
from collections import namedtuple
from typing import Any

import xmltodict

from supergrep.eva.utils import merge_dicts
from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.utils import (
    flatten_dict, ordered_jsons, sheet_process_output, search_tag_value)


def process(workbook: Any, contents: list) -> None:
    """Process Controller worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Controller'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    excel_header = [
        'controllername', 'datablocksize', 'modelnumber', 'productnumber',
        'serialnumber', 'firmwareversion', 'cachecondition', 'readcapacity',
        'writecapacity', 'mirrorcapacity', 'portname', 'topology',
        'hostportaddress', 'switchtype'
    ]
    build_header(worksheet, excel_header)
    RowTuple = namedtuple('RowTuple', excel_header)

    header = [
        'controllername', 'datablocksize', 'modelnumber', 'productnumber',
        'serialnumber', 'firmwareversion', 'cachememory/cachecondition',
        'cachememory/readcapacity', 'cachememory/writecapacity',
        'cachememory/mirrorcapacity', 'hostports/hostport/portname',
        'hostports/hostport/topology', 'hostports/hostport/hostportaddress',
        'deviceports/deviceport/switchtype'
    ]

    rows = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        raw_data = [flatten_dict(det_dict)
                    for det_dict in search_tag_value(doc, 'object')]
        for main_dict in raw_data:
            entry = list(ordered_jsons([main_dict], header[:6]))
            if entry:
                main_dict['hostports/hostport'] = merge_dicts(
                    main_dict['hostports/hostport'])

                main_dict['deviceports/deviceport'] = merge_dicts(
                    main_dict['deviceports/deviceport'])

                main_dict = flatten_dict(main_dict)
                rows += ordered_jsons([main_dict], header)

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
        'ControllerTable',
        'Controller',
        final_col,
        final_row)
