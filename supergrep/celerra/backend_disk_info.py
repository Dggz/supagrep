"""Backend Disk Info Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output


BACKEND_DISK_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required ID (.+)
    Value Capacity (.+)
    Value UsedCapacity (.+)
    Value DiskGroup (.+)
    Value Hidden (.+)
    Value Type (.+)
    Value Vendor (.+)
    Value State (.+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> NASStart

    NASStart
      ^\s*Output from:\s+/nas/bin/nas_storage\s+-info\s+-all\s+-o\s+sync=no -> SpindleStart

    SpindleStart
      ^\s*Spindles\s*$$ -> SpindleLines

    SpindleLines
      ^\s*id\s+=\s+${ID}
      ^\s*capacity\s+=\s+${Capacity}
      ^\s*used_capacity\s+=\s+${UsedCapacity}
      ^\s*disk_group\s+=\s+${DiskGroup}
      ^\s*hidden\s+=\s+${Hidden}
      ^\s*type\s+=\s+${Type}
      ^\s*vendor\s+=\s+${Vendor}
      ^\s*state\s+=\s+${State} -> Record SpindleLines
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Backend Disk Info worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Backend Disk Info')

    headers = get_parser_header(BACKEND_DISK_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    backend_storage_out = run_parser_over(content, BACKEND_DISK_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, backend_storage_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(column_format(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'BackendDiskInfoTable',
        'Backend Disk Info',
        final_col,
        final_row)
