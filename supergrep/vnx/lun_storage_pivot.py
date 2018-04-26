"""LUN-Storage_Pivot Sheet"""
from collections import namedtuple
from operator import itemgetter

from cytoolz.curried import map, groupby
from openpyxl import Workbook
from openpyxl.styles import Font

from supergrep.formatting import (
    build_header, style_value_cell, set_cell_to_number)
from supergrep.utils import sheet_process_output


def process(workbook: Workbook, content: list) -> None:
    """Process LUNsPivot worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'LUN-Storage_Pivot'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    state_header = ['ArrayName', 'LUNCount', 'SumOfLUNCapacityTB']

    StateTuple = namedtuple(
        'StateTuple', state_header)

    build_header(worksheet, state_header)

    array_storage_groups = groupby(itemgetter(0, 3), content)
    array_groups = groupby(itemgetter(0), array_storage_groups)

    state_rows, grand_total = [], [0, 0]  # type: list, list
    for array in array_groups:
        lun_count, lun_capacity, storage_array_rows = 0, 0, []
        for array_group in array_groups[array]:
            array_luns = list(zip(*array_storage_groups[array_group]))

            row = [
                array_group[1],
                len(array_storage_groups[array_group]),
                sum(map(float, array_luns[12]))
            ]

            lun_count += row[1]
            lun_capacity += row[2]
            storage_array_rows.append(map(str, row))

        grand_total[0], grand_total[1] = grand_total[0] + lun_count, \
            grand_total[1] + lun_capacity
        state_rows += [
            [array, str(lun_count), str(lun_capacity)],
            *storage_array_rows
        ]

    state_rows.append(['Grand Total', *grand_total])
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(StateTuple._make, state_rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            col_value = str(col_value) \
                if not isinstance(col_value, str) else col_value
            cell.value = col_value
            style_value_cell(cell)
            if row_tuple.ArrayName \
                    in [sg for sg in array_groups.keys()] + ['Grand Total']:
                cell.font = Font(bold=True, size=11)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'LUNStorageTable',
        'LUN-Storage_Pivot',
        final_col,
        final_row)
