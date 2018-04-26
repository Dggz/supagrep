"""Dskgrp_summary parser"""

import textwrap
from collections import namedtuple
from typing import Any

from openpyxl.comments import Comment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

DSKRGP_TMPL = textwrap.dedent("""\
    Value Filldown,Required SymmetrixID (\d+)
    Value Required GroupNumber (\d+)
    Value GroupName ([\d\w_]+) 
    Value DiskCount (\d+)
    Value DiskFlags (..)
    Value DiskSpeed (\d+)
    Value DiskSize (\d+)
    Value HyperSizeMB (\w+)
    Value CapacityMB (\d+)
    Value FreeCapacityMB (\d+)

    Start
      ^\s*Symmetrix ID:\s*${SymmetrixID} -> DSKGRPLine
      
    DSKGRPLine
      ^\s*${GroupNumber}\s+${GroupName}\s+${DiskCount}\s+${DiskFlags}\s+${DiskSpeed}\s+${DiskSize}\s+${HyperSizeMB}\s+${CapacityMB}\s+${FreeCapacityMB} -> Record
      ^\s*Total -> Start
""")


legend = textwrap.dedent("""
    Legend:
        Disk (L)ocation:
            I = Internal, X = External, - = N/A
        (T)echnology:
            S = SATA, F = Fibre Channel, E = Enterprise Flash Drive, - = N/A
""")


def process(workbook: Any, content: str) -> None:
    """Process dskgrp_summary worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'dskgrp_summary'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(DSKRGP_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    headers[5], headers[6], headers[8] = \
        'diskspeed(RPM)', 'disksize(MB)', 'totalcapacity(MB)'
    build_header(worksheet, headers)
    worksheet['E1'].comment = Comment(legend, '')

    dskgrp_summary_out = run_parser_over(content, DSKRGP_TMPL)
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, dskgrp_summary_out), 2):
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
        'DskgrpSummaryTable',
        'dskgrp_summary',
        final_col,
        final_row)
