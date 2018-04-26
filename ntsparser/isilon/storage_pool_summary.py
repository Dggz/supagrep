"""Storage Pool Summary sheet"""
import textwrap
from collections import namedtuple
from typing import Any

import xmltodict

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.isilon.utils import collected_data
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output, search_tag_value

POOL_TMPL = textwrap.dedent("""\
    Value Name (.+)
    Value Nodes (.+)
    Value RequestedProtection (.+)
    Value Required Type (nodepool)
    Value HDDUsed (.+)
    Value HDDTotal (.+)
    Value HDDPercentUsed (.+)
    Value SSDUsed (.+)
    Value SSDTotal (.+)
    Value SSDPercentUsed (.+)

    Start
      ^\s*Name:\s+${Name}
      ^\s*Nodes:\s+${Nodes}
      ^\s*Requested Protection:\s+${RequestedProtection}
      ^\s*Type:\s+${Type}
      ^\s*HDD Used:\s+${HDDUsed}
      ^\s*HDD Total:\s+${HDDTotal}
      ^\s*HDD % Used:\s+${HDDPercentUsed}
      ^\s*SSD Used:\s+${SSDUsed}
      ^\s*SSD Total:\s+${SSDTotal}
      ^\s*SSD % Used:\s+${SSDPercentUsed} -> Record
""")


def process(workbook: Any, contents: list) -> None:
    """Process Storage Pool Summary worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Storage Pool Summary'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = ['Hostname'] + get_parser_header(POOL_TMPL)

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    rows = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        pool = []  # type: list
        host = component_details['hostname']
        for entry in command_details:
            pool_content = collected_data(
                entry, 'cmd', 'isi storagepool list --format?list')
            pool = run_parser_over(
                pool_content, POOL_TMPL) if pool_content else pool
        rows += [[host] + row for row in pool]

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
        'StoragePoolSummaryTable',
        'Storage Pool Summary',
        final_col,
        final_row)
