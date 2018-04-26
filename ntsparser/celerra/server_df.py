"""server_df Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from ntsparser.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output


SERVER_DF_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required,Filldown Server (\S+)
    Value Required Filesystem (\S+)
    Value Kbytes (\S+)
    Value Used (\S+)
    Value Avail (\S+)
    Value Capacity (\S+)
    Value MountedOn (\S+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> DFCheck
      
    DFCheck
      ^\s*(\*+) -> DFStart

    DFStart
      ^\s*Output from:\s+/nas/bin/server_df\s+ -> DMLine
      ^\s*(\*+) -> Start
      
    DMLine
      ^\s*${Server}\s+:\s*$$ -> DFHeader 

    DFHeader
      ^\s*Filesystem\s+kbytes\s+used\s+avail\s+capacity\s+Mounted on -> DFLines
      ^\s*Error -> DFStart
      ^\s*(\*+) -> Start
      
    DFLines
      ^\s*${Filesystem}\s*$$
      ^\s*${Kbytes}\s+${Used}\s+${Avail}\s+${Capacity}\s+${MountedOn}\s*$$ -> Record DFLines
      ^\s*${Filesystem}\s+${Kbytes}\s+${Used}\s+${Avail}\s+${Capacity}\s+${MountedOn}\s*$$ -> Record DFLines
      ^\s*Output from:\s+/nas/bin/server_df\s+ -> DMLine
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process server_df worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('server_df')

    headers = get_parser_header(SERVER_DF_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    server_df_out = run_parser_over(content, SERVER_DF_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, server_df_out), 2):
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
        'serverDfTable',
        'server_df',
        final_col,
        final_row)
