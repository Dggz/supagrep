"""StorageArrayPivot Sheet"""
from collections import namedtuple
from operator import itemgetter

from cytoolz.curried import map, groupby
from openpyxl import Workbook
from openpyxl.styles import Font

from supergrep.formatting import \
    (build_header, style_value_cell, set_cell_to_number)
from supergrep.utils import sheet_process_output, write_excel


def process(workbook: Workbook, content: tuple) -> None:
    """Process StorageArrayPivot worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'StorageArrayPivot'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    speed_header = [
        'ArrayName', 'SumRegisteredInitiators',
        'SumLoggedInInitiators', 'SumNotLoggedInInitiators'
    ]
    model_header = ['ArrayModel', 'Count']
    revision_header = ['Revision', 'Count']

    SpeedTuple = namedtuple(
        'RowTuple', speed_header)  # pylint: disable=invalid-name
    ModelTuple = namedtuple(
        'RowTuple', model_header)  # pylint: disable=invalid-name
    RevisionTuple = namedtuple(
        'RowTuple', model_header)  # pylint: disable=invalid-name

    build_header(worksheet, speed_header)
    build_header(worksheet, model_header, 'F')
    build_header(worksheet, revision_header, 'I')

    speed_array_groups = groupby(itemgetter(6, 0), content[0])
    speed_groups = groupby(itemgetter(0), speed_array_groups)

    speed_rows, grand_total = [], [0, 0, 0]  # type: list, list
    for speed in speed_groups:
        total_initiators = [0, 0, 0]
        array_rows = []
        for array in speed_groups[speed]:
            array_initiators = list(zip(*speed_array_groups[array]))

            row = [
                array[1], sum(map(int, array_initiators[3])),
                sum(map(int, array_initiators[4])),
                sum(map(int, array_initiators[5]))
            ]

            total_initiators = [x + y
                                for x, y in zip(total_initiators, row[1:])]
            array_rows.append(map(str, row))
        grand_total = [x + y for x, y in zip(grand_total, total_initiators)]
        speed_rows += [[speed] + list(map(str, total_initiators)), *array_rows]

    speed_rows += [['Grand Total'] + grand_total]
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(SpeedTuple._make, speed_rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            col_value = str(col_value) \
                if not isinstance(col_value, str) else col_value
            cell.value = col_value
            style_value_cell(cell)
            if row_tuple.ArrayName \
                    in [sg for sg in speed_groups.keys()] + ['Grand Total']:
                cell.font = Font(bold=True, size=11)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'StorageArrayPivotTable',
        'StorageArrayPivot',
        final_col,
        final_row)

    model_rows = [(key, len(val)) for key, val in content[1].items()]
    model_rows.append(('Total', sum([row[1] for row in model_rows])))
    final_col, final_row = write_excel(model_rows, worksheet, ModelTuple, 'F')
    sheet_process_output(
        worksheet,
        'ModelTable',
        'StorageArrayPivot',
        final_col,
        final_row,
        start_col=ord('F'))

    revision_rows = [(key, len(val)) for key, val in content[2].items()]
    revision_rows.append(('Total', sum([row[1] for row in revision_rows])))
    final_col, final_row = write_excel(
        revision_rows, worksheet, RevisionTuple, 'I')
    sheet_process_output(
        worksheet,
        'RevisionTable',
        'StorageArrayPivot',
        final_col,
        final_row,
        start_col=ord('I'))
