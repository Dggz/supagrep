"""Performance Output Sheet (XtremIO)"""
import csv
from typing import Any

from supergrep.formatting import build_header, set_cell_to_number, style_value_cell
from supergrep.utils import sheet_process_output, percentile
from supergrep.xtremio.utils import compute_row, store_summary


def process(workbook: Any, content: str) -> list:
    """Process Performance Output worksheet (XtremIO)

    :param workbook:
    :param content:
    :return:
    """
    worksheet = workbook.get_sheet_by_name('Performance Output')

    perf_csv = csv.reader(content.split('\n'))
    header = next(perf_csv)
    header += [
        'Per Interval Read % by throughput', 'Per interval total throughput',
        'Per interval write io size', 'per interval read io size',
        'Per interval total io size (KB/s)'
    ]
    build_header(worksheet, header)

    # we check for '' in each row to remove blank rows that are in the file
    # perf_csv is a list of lists containing all the rows of the file
    rows = filter(
        lambda x: '' not in x and len(x) > 1 and x != header[:8], perf_csv)

    summary_data = ([], [], [], [], [], [], [], [], [])  # type: tuple

    final_col, final_row, start_date, end_date = 0, 0, '', ''
    for row_n, row_tuple in enumerate(rows, 2):
        start_date = row_tuple[0] if start_date == '' else start_date
        xls_row = compute_row(row_tuple)
        summary_data = store_summary(summary_data, xls_row)
        for col_n, col_value in \
                enumerate(xls_row, ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            set_cell_to_number(cell)
            style_value_cell(cell)
            final_col = col_n
        final_row = row_n
        end_date = row_tuple[0]

    perf_summary = list()
    for summary_column, data_origin in zip(
            summary_data, header[1:5] + header[8:]):
        perf_summary.append((
            data_origin, str(sum(summary_column) / len(summary_column)),
            str(percentile(summary_column, 0.95)), str(max(summary_column))))

    perf_summary += [
        ('Perf Collection Beginning Date:', start_date, '', ''),
        ('Perf Collection End Date:', end_date, '', '')
    ]

    sheet_process_output(
        worksheet,
        'PerformanceOutputTable',
        'Performance Output',
        final_col,
        final_row)

    return perf_summary
