"""System Details Sheet"""
import textwrap
from collections import namedtuple, Counter
from typing import Any

from openpyxl.styles import Alignment

from ntsparser.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output


SYSTEM_DETAILS_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Model (.+)
    Value ModelType (.+)
    Value SerialNo (.+)
    Value Required Node (.+)
    Value SPMemory (.+)
    Value Cabinet (.+)
    Value List DET (\S+)
    Value List DETCounts (.+)
    Value Filldown ServerVersion (.+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine
      ^\s*Output from:\s+/nas/sbin/navicli\s+-h\s+(\S+)\s+getagent -> AgentLine
      ^\s*^\s*Output from:\s+/nas/sbin/navicli\s+-h\s+(\S+)\s+getcrus -> CrusLines
      
    HostLine
      ^\s*${Hostname}\s*$$ -> ServerLine
      
    ServerLine
      ^\s*(\S+)\s+:\s+Product:\s+(.+)Version:\s+${ServerVersion} -> CrusStart
      
    CrusStart
      ^\s*^\s*Output from:\s+/nas/sbin/navicli\s+-h\s+(\S+)\s+getcrus -> CrusLines
      
    CrusLines
      ^\s*${DET}\s+Bus\s+(\d+)\s+Enclosure\s+(\d+)
      ^\s*Time on CS -> AgentStart
      
    AgentStart
      ^\s*Output from:\s+/nas/sbin/navicli\s+-h\s+(\S+)\s+getagent -> AgentLine
      
    AgentLine
      ^\s*Node:\s*${Node}
      ^\s*Model:\s*${Model}
      ^\s*Model Type:\s*${ModelType}
      ^\s*SP Memory:\s*${SPMemory}
      ^\s*Serial No:\s*${SerialNo}
      ^\s*Cabinet:\s*${Cabinet} -> Record Start
""")


def process(workbook: Any, content: str) -> None:
    """Process System Details worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('System Details')

    headers = get_parser_header(SYSTEM_DETAILS_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    system_details_out = run_parser_over(content, SYSTEM_DETAILS_TMPL)

    det_count = slice(-3, -1)
    for system_entry in system_details_out:
        det_counts = Counter(system_entry[-3])
        system_entry[det_count] = \
            [det for det in det_counts], \
            [str(count) for count in det_counts.values()]

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, system_details_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(column_format(col_n), row_n)]
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
        'SystemDetailsTable',
        'System Details',
        final_col,
        final_row)
