"""File System by Protocol Summary sheet"""
from collections import namedtuple
from operator import itemgetter
from typing import Any

from cytoolz.curried import groupby

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.isilon.utils import get_multiprotocol
from ntsparser.utils import sheet_process_output


def process(workbook: Any, nfs_rows: list, smb_rows: list) -> None:
    """Process File System by Protocol Sheet

    :param workbook:
    :param nfs_rows:
    :param smb_rows:
    """
    worksheet_name = 'File System by Protocol'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = ['Hostname', 'FileSystem', 'Path', 'Type']

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    key_columns = (0, 1)
    rows = groupby(itemgetter(*key_columns), nfs_rows + smb_rows)
    rows = [get_multiprotocol(rows[i]) for i in rows]

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'FileSystemProtocolTable',
        'File System by Protocol',
        final_col,
        final_row)
