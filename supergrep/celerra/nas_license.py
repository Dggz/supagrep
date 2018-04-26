"""NAS_License Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

NAS_LICENSE_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required Key (\S+)
    Value Status (\S+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> LicenseStart
      
    LicenseStart
      ^\s*key\s+status\s+value -> LicenseLines

    LicenseLines
      ^\s*${Key}\s+${Status}\s+(.+) -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process NAS_License worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('NAS_License')

    headers = get_parser_header(NAS_LICENSE_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    nas_license_out = run_parser_over(content, NAS_LICENSE_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, nas_license_out), 2):
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
        'NASLicenseTable',
        'NAS_License',
        final_col,
        final_row)
