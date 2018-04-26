"""Volume Performance Sheet (XtremIO)"""
import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map
from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output


SHOW_VOLUME_PERFORMANCE_TMPL = textwrap.dedent("""\
    Value VolumeName (\S+|\S+\s\S+)
    Value ClusterName (\S+)
    Value WriteIOPS (\d+)
    Value ReadIOPS (\d+)
    Value IOPS (\d+)
    Value TotalWriteIOs (\d+)
    Value TotalReadIOs (\d+)

    Start
      ^\s*Volume-Name\s+Index\s+Cluster-Name\s+Index\s+Write-BW\(MB/s\) -> VolumePerformanceLine

    VolumePerformanceLine
      ^\s*${VolumeName}\s+(\S+)\s+${ClusterName}\s+(\d+)\s+(\S+)\s+${WriteIOPS}\s+(\S+)\s+${ReadIOPS}\s+(\S+)\s+${IOPS}\s+${TotalWriteIOs}\s+${TotalReadIOs}\s* -> Record
""")


def process(workbook: Any, content: str) -> None:
    """Process Volume Performance worksheet (XtremIO)

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Volume Performance')

    headers = get_parser_header(SHOW_VOLUME_PERFORMANCE_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_volume_performance_out = run_parser_over(content, SHOW_VOLUME_PERFORMANCE_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, show_volume_performance_out), 2):
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
        'VolumePerformanceTable',
        'Volume Performance',
        final_col,
        final_row)
