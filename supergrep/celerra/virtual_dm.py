"""VIRTUAL DM Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

NAS_VIRTUAL_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required ID (\S+)
    Value ACL (\S+)
    Value Server (\S+)
    Value MountedFS (\S+|\s+)
    Value RootFS (\S+)
    Value Name (\S+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> NasStart

    NasStart
      ^\s*Output from:\s+/nas/bin/nas_server\s+-list\s+-all -> NasHeader

    NasHeader
      ^\s*id\s+acl\s+server -> NasLines

    NasLines
      ^\s*${ID}\s+${ACL}\s+${Server}\s+${MountedFS}\s+${RootFS}\s+${Name}\s*$$ -> Record NasLines
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process VIRTUAL DM worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('VIRTUAL DM')

    headers = get_parser_header(NAS_VIRTUAL_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    nas_virtual_out = run_parser_over(content, NAS_VIRTUAL_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, nas_virtual_out), 2):
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
        'VIRTUALDMTable',
        'VIRTUAL DM',
        final_col,
        final_row)
