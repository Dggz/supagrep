"""Ops Sheet"""
import csv
from typing import Any

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.isilon.utils import perf_cluster_name
from ntsparser.utils import sheet_process_output


def process(workbook: Any, content: list) -> list:
    """Process Ops worksheet

    :param workbook:
    :param content:
    :return:
    """
    worksheet = workbook.get_sheet_by_name('Ops')

    rows, header = [], []  # type: list, list
    for csv_file in content:
        cluster = perf_cluster_name(csv_file)
        ops_csv = csv.reader(csv_file.split('\n')[1:])
        header = ['Cluster'] + next(ops_csv) if header == [] else header
        rows += [[cluster] + csv_row
                 for csv_row in ops_csv if csv_row not in ([], header[1:])]

    if header:
        header[3], header[5] = header[3] + ' 2', header[5] + ' 3',
        header[7], header[9] = header[7] + ' 4', header[9] + ' 5',
        header[11], header[13] = header[11] + ' 6', header[13] + ' 7',
        header[15], header[-2] = header[15] + ' 8', header[-2] + ' 9',
    build_header(worksheet, header)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(rows, 2):
        for col_n, col_value in enumerate(row_tuple, ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            set_cell_to_number(cell)
            style_value_cell(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'OpsTable',
        'Ops',
        final_col,
        final_row)

    return rows
