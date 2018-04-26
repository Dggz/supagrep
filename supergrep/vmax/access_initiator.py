"""Access_initiator parser"""

import textwrap
from collections import namedtuple
from typing import Any

from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

ACSINIT_TMPL = textwrap.dedent("""\
    Value Required SymmetrixID (\d+)
    Value Required InitiatorGroup (\S+)
    Value LastUpdate (.+) 
    Value GroupLastUpdate (.+)
    Value List HostInitiators ((WWN\s*:\s*.+)|(alias:\s*.+)|(IG\s*:.+))
    Value MaskingViewName (\w+)
    Value ParentInitiatorGroup (\w+\S+)

    Start
      ^\s*Symmetrix ID\s+:\s*${SymmetrixID} -> AccessLine
      
    AccessLine
      ^\s*Initiator Group Name\s*:\s*${InitiatorGroup}
      ^\s*Last update time\s*:\s*${LastUpdate}
      ^\s*Group last update time\s*:\s*${GroupLastUpdate}
      ^\s*Host Initiators\s* -> HostLine
      
    HostLine
      ^\s*${HostInitiators}
      ^\s*\[${HostInitiators}\]
      ^\s*Masking View Names -> MaskingLine
      
    MaskingLine
      ^\s*${MaskingViewName} -> ParentStart
      
    ParentStart
      ^\s*Parent Initiator Groups -> ParentLine
      
    ParentLine
      ^\s*${ParentInitiatorGroup} -> Record Start
""")


def process(workbook: Any, content: str) -> None:
    """Process access_initiator worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'access_initiator'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(ACSINIT_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)

    access_initiator_out = run_parser_over(content, ACSINIT_TMPL)
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, access_initiator_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            if isinstance(col_value, str):
                cell.value = str.strip(col_value)
            else:
                cell.alignment = Alignment(wrapText=True)
                cell.value = '\n'.join(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'AccessInitiatorTable',
        'access_initiator',
        final_col,
        final_row)
