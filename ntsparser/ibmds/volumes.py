"""Volumes Sheet"""
import csv
from typing import Any

from ntsparser.formatting import build_header, set_cell_to_number, \
    style_value_cell
from ntsparser.ibmds.utils import get_rows
from ntsparser.utils import sheet_process_output


def process(workbook: Any, content: list) -> None:
    """Process Volumes worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Volumes')

    header = [
        'SystemName', 'Name', 'State', 'Storage Type', 'Total Capacity (GiB)',
        'Allocated Capacity (GiB)', 'Pool', 'Node', 'LUN ID',
        'Thin-provisioning', 'Allocation Method', 'LSS', 'Address Group',
        'MTM', 'Data Type', 'VOLSER', 'GUID', 'Number of Hosts', 'Scope',
        'Performance Policy', 'Migrating Capacity'
    ]
    build_header(worksheet, header)

    rows = []  # type: list
    for csv_file in content:
        storage_csv = csv.reader(csv_file.split('\n'))
        system_name = get_rows(storage_csv, ['Name'])
        volumes_rows = get_rows(storage_csv, header[1:])
        rows += [system_name[0] + feat_row for feat_row in volumes_rows]

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
        'VolumesTable',
        'Volumes',
        final_col,
        final_row)
