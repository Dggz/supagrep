"""fs_dedupe Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from ntsparser.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output

FS_DEDUPE_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required ID (\d+)
    Value Name (\S+)
    Value State (\S+)
    Value Status (\S+|\S+\s\(\S+\s+\S+\)|\s+)
    Value TimeOfLastScan (\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+)
    Value OriginalDataSize (\S+\s\S+)
    Value Usage (\S+)
    Value SpaceSaved (\S+\s\S+\s\S+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> DedupeStart

    DedupeStart
      ^\s*Output from:\s+/nas/bin/fs_dedupe\s+-list -> DedupeHeader

    DedupeHeader
      ^\s*id\s+name\s+state\s+status\s+time_of_last_scan\s+original_data_size\s+usage\s+space_saved -> DedupeLines

    DedupeLines
      ^\s*${ID}\s+${Name}\s+${State}\s+${Status}\s+${TimeOfLastScan}\s+${OriginalDataSize}\s+${Usage}\s+${SpaceSaved}\s*$$ -> Record DedupeLines
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process fs_dedupe worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('fs_dedupe')

    headers = get_parser_header(FS_DEDUPE_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    fs_dedupe_out = run_parser_over(content, FS_DEDUPE_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, fs_dedupe_out), 2):
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
        'fsDedupeTable',
        'fs_dedupe',
        final_col,
        final_row)
