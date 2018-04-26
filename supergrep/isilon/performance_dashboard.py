"""Performance Dashboard Sheet"""
from collections import defaultdict
from operator import itemgetter
from typing import Any

from cytoolz.curried import groupby

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.isilon.utils import perf_dashboard
from supergrep.utils import sheet_process_output


def process(workbook: Any, content: tuple) -> None:
    """Process Performance Dashboard worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Performance Dashboard')

    header = [
        'Cluster', 'Top Level Directories', 'No of Directories', 'No of Files',
        'Throughput Avg', 'Operations Avg', 'Ops 95th Percentile',
        'Latency Avg', 'Total', '< 128k', '< 128k Percent', '> 128k',
        '> 128k Percent'
    ]
    build_header(worksheet, header)

    data = [groupby(itemgetter(0), cont) for cont in content]

    data_dict = defaultdict(tuple)  # type: dict
    for pos, _ in enumerate(data):
        for key in data[pos]:
            data_dict[key] = (*data_dict[key], data[pos][key])

    rows = []  # type: list
    for cluster in data_dict:
        rows.append(
            [cluster] + list(map(str, perf_dashboard(data_dict[cluster]))))

    rows.append(['Total'] + list(map(str, perf_dashboard(content))))

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(rows, 2):
        for col_n, col_value in \
                enumerate(row_tuple, ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            set_cell_to_number(cell, '0.00')
            style_value_cell(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'PerformanceDashboardTable',
        'Performance Dashboard',
        final_col,
        final_row)
