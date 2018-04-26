"""CPG (3Par) Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output


SHOWCPG_TMPL = textwrap.dedent("""\
    Value Filldown Name (\S+)
    Value Filldown Serial (\S+)
    Value Required ID (\S+)
    Value CPGName (\S+)
    Value Domain (\S+)
    Value WarnPercent (\S+)
    Value VVs (\S+)
    Value TPVVs (\S+)
    Value Usr (\S+)
    Value Snp (\S+)
    Value UserTotal (\S+)
    Value UserUsed (\S+)
    Value SnpTotal (\S+)
    Value SnpUsed (\S+)
    Value AdmTotal (\S+)
    Value AdmUsed (\S+)

    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+(\S+\s\S+|\S+)\s+${Serial}\s+
      ^\s*(\*+) -> CPGStart

    CPGStart
      ^\s*Id\s+Name\s+Domain\s+Warn%\s+VVs\s+TPVVs -> CPGLines

    CPGLines
      ^\s*${ID}\s+${CPGName}\s+${Domain}\s+${WarnPercent}\s+${VVs}\s+${TPVVs}\s+(\S+)*\s+${Usr}\s+${Snp}\s+${UserTotal}\s+${UserUsed}\s+${SnpTotal}\s+${SnpUsed}\s+${AdmTotal}\s+${AdmUsed} -> Record CPGLines
      ^\s*(-+) -> Start
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process CPG (3Par) worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('CPG')

    headers = get_parser_header(SHOWCPG_TMPL)
    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_cpg_out = run_parser_over(content, SHOWCPG_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_cpg_out), 2):
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
        'CPGTable',
        'CPG',
        final_col,
        final_row)
