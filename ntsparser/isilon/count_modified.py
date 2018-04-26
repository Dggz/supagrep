"""Count Modified Sheet"""
import csv
from typing import Any

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.isilon.utils import perf_cluster_name
from ntsparser.utils import sheet_process_output


def process(workbook: Any, content: list) -> None:
    """Process Count Modified worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Count Modified')

    rows, header = [], []  # type: list, list
    for csv_file in content:
        cluster = perf_cluster_name(csv_file)
        count_mod_csv = csv.reader(csv_file.split('\n')[1:])
        header = ['Cluster'] + next(count_mod_csv) if header == [] else header
        rows += [[cluster] + csv_row
                 for csv_row in count_mod_csv
                 if csv_row not in ([], header[1:])]

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
        'CountModifiedTable',
        'Count Modified',
        final_col,
        final_row)
