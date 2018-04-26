"""License Summary Summary sheet"""
import textwrap
from collections import namedtuple
from typing import Any

import xmltodict

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.isilon.utils import collected_data
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output, search_tag_value

LICENSE_TMPL = textwrap.dedent("""\
    Value Name (.+)
    Value Status (\S+)
    Value Expiration (\S+)

    Start
      ^\s*Name\s+Status\s+Expiration -> LicenseLines
      
    LicenseLines
      ^\s*${Name}\s+${Status}\s+${Expiration}\s*$$ -> Record
""")


def process(workbook: Any, contents: list) -> None:
    """Process License Summary worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'License Summary'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = ['Hostname'] + get_parser_header(LICENSE_TMPL)

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    rows = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        licenses = []  # type: list
        host = component_details['hostname']
        for entry in command_details:
            license_content = collected_data(
                entry, 'cmd', 'isi license*[licenses]*list')
            licenses = run_parser_over(
                license_content, LICENSE_TMPL)\
                if license_content else licenses
        rows += [[host] + row for row in licenses]

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
        'LicenseSummaryTable',
        'License Summary',
        final_col,
        final_row)
