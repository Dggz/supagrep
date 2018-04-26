"""License (3Par) Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output

SHOW_FEATURES_TMPL = textwrap.dedent("""\
    Value Filldown Name (\S+)
    Value Filldown Serial (\S+)
    Value Required EnabledLicenseFeature (.+)

    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+(\S+\s\S+|\S+)\s+${Serial}\s+
      ^\s*(\*+) -> FeaturesStart

    FeaturesStart
      ^\s*License features currently enabled: -> FeaturesLines

    FeaturesLines
      ^\s*License features enabled on a trial -> Start
      ^\s*${EnabledLicenseFeature} -> Record FeaturesLines
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process License (3Par) worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('License')

    headers = get_parser_header(SHOW_FEATURES_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_features = run_parser_over(content, SHOW_FEATURES_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_features), 2):
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
        'LicenseTable',
        'License',
        final_col,
        final_row)
