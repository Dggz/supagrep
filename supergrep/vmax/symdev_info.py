"""symdev_info parser"""

import textwrap
from collections import namedtuple
from typing import Any

from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

SYMDEV_TMPL = textwrap.dedent("""\
    Value Filldown,Required SymmetrixID (\d+)
    Value Filldown,Required InitiatorGroupName (.+)
    Value LastUpdate (.+)
    Value GroupLastUpdate (.+)
    Value List HostInitiators (.+)
    Value List Alias (.+)
    Value List IG (.+)
    Value Filldown,Required Symdev (\S+)
    Value List,Required DirPort (\S+)
    Value List PyshicalDeviceName (\S+\s\S+|\S+)
    Value List HostLUN (\S+)
    Value List Attr (\(.\)|\s*)
    Value Capacity (\d+)
    Value List MaskingViewName (.+)
    
    Start
      ^\s*Symmetrix ID\s+:\s*${SymmetrixID}
      ^\s*Initiator Group Name\s*:\s*${InitiatorGroupName}  
      ^\s*Last update time\s*:\s*${LastUpdate}      
      ^\s*Group last update time\s*:\s*${GroupLastUpdate} -> HostStart
      
    HostStart
      ^\s+WWN\s*:\s*${HostInitiators}
      ^\s+\[alias\s*:\s*${Alias}\]
      ^\s+IG\s*:\s*${IG}
      ^\s+Sym\s*Host\s+Cap -> SymDevLine
      ^\s*(\*+) -> Start
      
    SymDevLine
      ^\s*${Symdev}\s+${DirPort}\s+${PyshicalDeviceName}\s+${HostLUN}\s+${Attr}\s+${Capacity}\s+${MaskingViewName}
      ^\s*${DirPort}\s+${PyshicalDeviceName}\s+${HostLUN}\s+${Attr}\s+${MaskingViewName} -> Record
      ^\s*Total Capacity\s* -> Record Start
""")


def process(workbook: Any, content: str) -> None:
    """Process symdev_info worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'symdev_info'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(SYMDEV_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    headers[8], headers[-2] = 'Dir:Port', 'Capacity(MB)'
    build_header(worksheet, headers)
    symdev_info_out = run_parser_over(content, SYMDEV_TMPL)
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, symdev_info_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            if isinstance(col_value, str):
                cell.value = str.strip(col_value)
            else:
                cell.alignment = Alignment(wrapText=True)
                cell.value = '\n'.join(col_value)
            style_value_cell(cell)
            if chr(col_n) not in ('H', ):
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'SymDevTable',
        'symdev_info',
        final_col,
        final_row)
