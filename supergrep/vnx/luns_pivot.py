"""LUNsPivot Sheet"""
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
    worksheet_name = 'LUNsPivot'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    state_header = ['ArrayName', 'SumOfLUNCapacityTB']
    private_header = ['PrivateArrayName', 'SumOfLUNCapacityTB', 'Count']

    StateTuple = namedtuple(
        'StateTuple', state_header)
    PrivateTuple = namedtuple(
        'PrivateTuple', private_header)

    build_header(worksheet, state_header)
    build_header(worksheet, private_header, 'D')

    state_array_groups = groupby(itemgetter(7, 0), content)
    state_groups = groupby(itemgetter(0), state_array_groups)

    private_array_groups = groupby(itemgetter(15, 0), content)
    private_groups = groupby(itemgetter(0), private_array_groups)

    state_rows, grand_total = [], 0  # type: list, float
    for state in state_groups:
        state_sum, array_rows = 0, []
        for array in state_groups[state]:
            array_luns = list(zip(*state_array_groups[array]))

            row = [array[1], sum(map(float, array_luns[12]))]

            state_sum += row[1]
            array_rows.append(map(str, row))
        grand_total += state_sum
        state_rows += [[state] + [str(state_sum)], *array_rows]

    state_rows += [['Grand Total', str(grand_total)]]
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
                    in [sg for sg in state_groups.keys()] + ['Grand Total']:
                cell.font = Font(bold=True, size=11)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'BoundTable',
        'LUNsPivot',
        final_col,
        final_row)

    private_rows, private_total = [], [0, 0]  # type: list, list
    for private in private_groups:
        private_sum, array_rows = [0, 0], []
        for array in private_groups[private]:
            array_luns = list(zip(*private_array_groups[array]))

            row = [array[1], sum(map(float, array_luns[12])), len(array_luns)]

            private_sum = [x + y for x, y in zip(private_sum, row[1:])]
            array_rows.append(map(str, row))
        private_total = [x + y for x, y in zip(private_total, private_sum)]
        private_rows += [[private] + list(map(str, private_sum)), *array_rows]

    private_rows += [['Grand Total'] + private_total]
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(PrivateTuple._make, private_rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('D')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            col_value = str(col_value) \
                if not isinstance(col_value, str) else col_value
            cell.value = col_value
            style_value_cell(cell)
            if row_tuple.PrivateArrayName \
                    in [pg for pg in private_groups.keys()] + ['Grand Total']:
                cell.font = Font(bold=True, size=11)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'PrivateTable',
        'LUNsPivot',
        final_col,
        final_row,
        start_col=ord('D'))
