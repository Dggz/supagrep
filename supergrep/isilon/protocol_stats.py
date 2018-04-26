"""Protocol Stats sheet"""
import textwrap
from collections import namedtuple
from typing import Any, Iterable

import xmltodict

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.isilon.utils import collected_data
from supergrep.parsing import run_parser_over, get_parser_header
from supergrep.utils import sheet_process_output, search_tag_value

PROTOCOL_TMPL = textwrap.dedent("""\
    Value Ops (\S+)
    Value In (\S+)
    Value Out (\S+)
    Value TimeAvg (\S+)
    Value TimeStdDev (\S+)
    Value Node (\S+)
    Value Proto (\S+)
    Value Class (\S+)
    Value Op (\S+)

    Start
      ^\s*Ops\s+In\s+Out\s+TimeAvg\s+TimeStdDev\s+Node\s+Proto\s+Class\s+Op -> ProtocolLines 
      
    ProtocolLines
      ^\s*${Ops}\s+${In}\s+${Out}\s+${TimeAvg}\s+${TimeStdDev}\s+${Node}\s+${Proto}\s+${Class}\s+${Op} -> Record
""")


def process(workbook: Any, contents: list) -> None:
    """Process Protocol Stats worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Protocol Stats'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = ['Hostname'] + get_parser_header(PROTOCOL_TMPL)

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    rows = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        protocols = []  # type: Iterable
        host = component_details['hostname']
        for entry in command_details:
            protocol_content = collected_data(
                entry, 'cmd', 'isi statistics protocol')
            protocols = run_parser_over(protocol_content, PROTOCOL_TMPL) \
                if protocol_content else protocols
        rows += [[host] + row for row in protocols]

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
        'ProtocolStatsTable',
        'Protocol Stats',
        final_col,
        final_row)
