"""Volumes Sheet (XtremIO)"""
import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map
from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output


SHOW_VOLUMES_TMPL = textwrap.dedent("""\
    Value VolumeName (\S+|\S+\s\S+(\s\S+)*)
    Value Index (\d+)
    Value ClusterName (\S+)
    Value VolSize (\d+\.\d+[PTGMK]|(\d+[PTGMK])|(\d+))
    Value LBSize (\d+)
    Value VSGSpace (\d+\.\d+[PTGMK]|(\d+[PTGMK])|(\d+))
    Value CreatedFromVolume (\S+|\S+\s\S+|\s+)
    Value TotalWrites (\d+\.\d+[PTGMK]|(\d+[PTGMK])|(\d+))
    Value TotalReads (\d+\.\d+[PTGMK]|(\d+[PTGMK])|(\d+))
    Value VolumeType (\w+)
    Value VolumeAccess (\S+)
    Value Tags (\S+\s\S+|\s*)

    Start
      ^\s*Volume-Name\s+Index\s+Cluster-Name\s+Index\s+Vol-Size\s+LB-Size\s+VSG-Space-In-Use -> VolumesLine

    VolumesLine
      ^\s*${VolumeName}\s+${Index}\s+${ClusterName}\s+(\d+)\s+${VolSize}\s+${LBSize}\s+${VSGSpace}\s+(\d+)\s+${CreatedFromVolume}\s+(\d+)\s+(\S+)\s+(\S+)\s+(\S+)\s+${TotalWrites}\s+${TotalReads}\s+(\S+|\s+)\s+(\S+)\s+(\S+)\s+${VolumeType}\s+(\S+\s+\S+)\s+${VolumeAccess}\s*$$ -> Record
      ^\s*${VolumeName}\s+${Index}\s+${ClusterName}\s+(\d+)\s+${VolSize}\s+${LBSize}\s+${VSGSpace}\s+(\d+)\s+${CreatedFromVolume}\s+(\d+)\s+(\S+)\s+(\S+)\s+(\S+)\s+${TotalWrites}\s+${TotalReads}\s+(\S+|\s+)\s+(\S+)\s+(\S+)\s+${VolumeType}\s+(\S+\s+\S+)\s+${VolumeAccess}\s+${Tags}\s*$$ -> Record
""")


def process(workbook: Any, content: str) -> None:
    """Process Volumes worksheet (XtremIO)

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Volumes')

    headers = get_parser_header(SHOW_VOLUMES_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_volumes_out = run_parser_over(content, SHOW_VOLUMES_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_volumes_out), 2):
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
        'VolumesTable',
        'Volumes',
        final_col,
        final_row)
