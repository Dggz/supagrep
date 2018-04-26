"""Features Sheet"""
import csv
from typing import Any

from supergrep.formatting import build_header, set_cell_to_number, \
    style_value_cell
from supergrep.ibmds.utils import get_rows
from supergrep.utils import sheet_process_output


def process(workbook: Any, content: list) -> None:
    """Process Features worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Features')

    header = [
        'SystemName', 'Name', 'Enabled', 'Capacity (TiB)',
        'Storage Type', 'Key Code'
    ]
    build_header(worksheet, header)

    rows = []  # type: list
    for csv_file in content:
        storage_csv = csv.reader(csv_file.split('\n'))
        system_name = get_rows(storage_csv, ['Name'])
        features_rows = get_rows(storage_csv, header[1:])
        rows += [system_name[0] + feat_row for feat_row in features_rows]

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(rows, 2):
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
        'FeaturesTable',
        'Features',
        final_col,
        final_row)
