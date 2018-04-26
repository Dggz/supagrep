"""Performance Summary Sheet (XtremIO)"""
from typing import Any

from ntsparser.formatting import build_header, set_cell_to_number, style_value_cell
from ntsparser.utils import sheet_process_output


def process(workbook: Any, content: list) -> None:
    """Process Performance Summary worksheet (XtremIO)

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Performance Summary')

    header = ['Column', 'Avg', '95th', 'Max']
    build_header(worksheet, header)

    final_col, final_row = 0, 0
    content += [('Logical capacity exposed:', '', '', '')]
    for row_n, row_tuple in enumerate(content, 2):
        for col_n, col_value in \
                enumerate(row_tuple, ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            set_cell_to_number(cell)
            style_value_cell(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'PerformanceSummaryTable',
        'Performance Summary',
        final_col,
        final_row)
