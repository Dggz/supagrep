"""Hosts (3Par) Sheet"""
import textwrap
from collections import namedtuple
from contextlib import suppress
from operator import itemgetter
from typing import Any

from cytoolz import concat, groupby

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output

SHOWHOST_TMPL = textwrap.dedent("""\
    Value Filldown Name (\S+)
    Value Filldown Serial (\S+)
    Value Required ID (\S+)
    Value HostName (\S+)
    Value WWNiSCSIName (\S+)
    Value Domain (\S+)
    Value Persona (\S+)
    Value Port (\S+)
    Value IP_addr (\S+)

    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+(\S+\s\S+|\S+)\s+${Serial}\s+
      ^\s*(\*+) -> HostStart

    HostStart
      ^\s*Id\s+Name\s+Domain\s+Persona\s+-WWN/iSCSI_Name-\s+Port\s+IP_addr -> HostLines

    HostLines
      ^\s*${ID}\s+${HostName}\s+${Domain}\s+${Persona}\s+${WWNiSCSIName}\s+${Port}\s+${IP_addr} -> Record HostLines
      ^\s*(\*+) -> Start
""")


SHOWHOST_LINES_TMPL = textwrap.dedent("""\
    Value Filldown Name (\S+)
    Value Filldown Serial (\S+)
    Value Required ID (\S+)
    Value HostName (\S+)
    Value Location (.+)
    Value OS (.+)
    Value Model (.+)

    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+(\S+\s\S+|\S+)\s+${Serial}\s+
      ^\s*(\*+) -> HostStart

    HostStart
      ^\s*(-*) Host (\S+) (-*) -> HostLines

    HostLines
      ^\s*Name\s*:\s*${HostName}
      ^\s*Id\s*:\s*${ID}
      ^\s*Location\s*:\s*${Location}
      ^\s*OS\s*:\s*${OS}
      ^\s*Model\s*:\s*${Model} -> Record
      ^\s*(\*+) -> Record Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Hosts (3Par) worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Hosts')

    headers = list(concat([
        get_parser_header(SHOWHOST_TMPL),
        get_parser_header(SHOWHOST_LINES_TMPL)[4:],
    ]))

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_hosts_out = groupby(
        itemgetter(0, 1, 2, 3, 4), run_parser_over(content, SHOWHOST_TMPL))
    show_hosts_lines_out = groupby(
        itemgetter(0, 1, 2, 3), run_parser_over(content, SHOWHOST_LINES_TMPL))

    rows = []
    for idfier in show_hosts_out:
        with suppress(KeyError):
            for host_line, details_line in \
                    zip(show_hosts_out[idfier], show_hosts_lines_out[idfier[:-1]]):
                rows.append(host_line + details_line[4:])

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
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
        'HostsTable',
        'Hosts',
        final_col,
        final_row)
