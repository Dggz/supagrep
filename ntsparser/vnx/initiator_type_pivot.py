"""InitiatorType Sheet"""
from collections import namedtuple
from operator import itemgetter

from cytoolz.curried import map, groupby
from openpyxl import Workbook
from openpyxl.styles import Font

from ntsparser.formatting import build_header, style_value_cell, \
    set_cell_to_number
from ntsparser.utils import sheet_process_output


# noinspection TaskProblemsInspection
def process(workbook: Workbook, content: list) -> None:
    """Process InitiatorType worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'InitiatorType'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    speed_header = ['ArrayName', 'SumRegisteredInitiators']

    RowTuple = namedtuple(
        'RowTuple', speed_header)  # pylint: disable=invalid-name

    build_header(worksheet, speed_header)

    type_array_groups = groupby(itemgetter(-2, 0), content)
    type_groups = groupby(itemgetter(0), type_array_groups)

    initiator_rows = []  # type: list
    for array_type in type_groups:
        total_initiators = 0
        array_rows = []
        for array in type_groups[array_type]:
            array_initiators = list(zip(*type_array_groups[array]))

            row = [array[1], sum(map(int, array_initiators[3]))]
            total_initiators += row[1]
            array_rows.append(map(str, row))
        initiator_rows += [[array_type] + [str(total_initiators)], *array_rows]

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, initiator_rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            col_value = str(col_value) \
                if not isinstance(col_value, str) else col_value
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            if row_tuple.ArrayName in type_groups.keys():
                cell.font = Font(bold=True, size=11)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'InitiatorTypeTable',
        'InitiatorType',
        final_col,
        final_row)
