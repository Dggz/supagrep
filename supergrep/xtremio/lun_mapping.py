"""Lun Mapping Sheet (XtremIO)"""
import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map
from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output


LUN_MAPPING_TMPL = textwrap.dedent("""\
    Value ClusterName (\S+)
    Value VolumeName (\S+|\S+\s\S+(\s\S+)*)
    Value IGName (\S+)
    Value LUN (\S+)
    Value MappingIndex (\S+)

    Start
      ^\s*Cluster-Name\s+Index\s+Volume-Name\s+Index\s+IG-Name\s+Index\s+TG-Name -> LunMappingLine

    LunMappingLine
      ^\s*${ClusterName}\s+(\S+)\s+${VolumeName}\s+(\d+)\s+${IGName}\s+(\S+)\s+(\S+)\s+(\S+)\s+${LUN}\s+${MappingIndex}\s* -> Record
""")


def process(workbook: Any, content: str) -> None:
    """Process Lun Mapping worksheet (XtremIO)

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Lun Mapping')

    headers = get_parser_header(LUN_MAPPING_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_lun_mapping_out = run_parser_over(content, LUN_MAPPING_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_lun_mapping_out), 2):
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
        'LunMappingTable',
        'Lun Mapping',
        final_col,
        final_row)
