"""DisksPivot Sheet"""
from collections import namedtuple
from operator import itemgetter

from cytoolz.curried import map, groupby
from openpyxl import Workbook
from openpyxl.styles import Font

from ntsparser.formatting import (
    build_header, style_value_cell, set_cell_to_number)
from ntsparser.utils import sheet_process_output


def process(workbook: Workbook, content: list) -> None:
    """Process DisksPivot worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'DisksPivot'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    type_header = ['ArrayName', 'DiskCount', 'SumOfCapacityTB']
    speed_header = ['SpeedArrayName', 'DiskCount']

    TypeTuple = namedtuple(
        'TypeTuple', type_header)
    SpeedTuple = namedtuple(
        'SpeedTuple', speed_header)

    build_header(worksheet, type_header)
    build_header(worksheet, speed_header, 'E')

    type_array_groups = groupby(itemgetter(-4, 0), content)
    type_groups = groupby(itemgetter(0), type_array_groups)

    speed_array_groups = groupby(itemgetter(-3, 0), content)
    speed_groups = groupby(itemgetter(0), speed_array_groups)

    type_rows, grand_total = [], [0, 0]  # type: list, list
    for row_type in type_groups:
        type_sum, array_rows = [0, 0], []
        for array in type_groups[row_type]:
            array_disks = list(zip(*type_array_groups[array]))

            row = [array[1], len(array_disks), sum(map(float, array_disks[7]))]

            type_sum = [x + y for x, y in zip(type_sum, row[1:])]
            array_rows.append(map(str, row))
        grand_total = [x + y for x, y in zip(grand_total, type_sum)]
        type_rows += [[row_type] + list(map(str, type_sum)), *array_rows]

    type_rows += [['Grand Total'] + grand_total]
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(TypeTuple._make, type_rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            col_value = str(col_value) \
                if not isinstance(col_value, str) else col_value
            cell.value = col_value
            style_value_cell(cell)
            if row_tuple.ArrayName \
                    in [sg for sg in type_groups.keys()] + ['Grand Total']:
                cell.font = Font(bold=True, size=11)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'TypeTable',
        'DisksPivot',
        final_col,
        final_row)

    speed_rows, speed_total = [], [0, 0]  # type: list, list
    for speed in speed_groups:
        speed_sum, array_rows = [0, 0], []
        for array in speed_groups[speed]:
            array_disks = list(zip(*speed_array_groups[array]))

            row = [array[1], len(array_disks)]

            speed_sum = [x + y for x, y in zip(speed_sum, row[1:])]
            array_rows.append(map(str, row))
        speed_total = [x + y for x, y in zip(speed_total, speed_sum)]
        speed_rows += [[speed] + list(map(str, speed_sum)), *array_rows]

    speed_rows += [['Grand Total'] + speed_total]
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(SpeedTuple._make, speed_rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('E')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            col_value = str(col_value) \
                if not isinstance(col_value, str) else col_value
            cell.value = col_value
            style_value_cell(cell)
            if row_tuple.SpeedArrayName \
                    in [pg for pg in speed_groups.keys()] + ['Grand Total']:
                cell.font = Font(bold=True, size=11)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'DiskSpeedTable',
        'DisksPivot',
        final_col,
        final_row,
        start_col=ord('E'))
