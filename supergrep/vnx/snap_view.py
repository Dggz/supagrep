"""SnapView Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output
from supergrep.vnx.utils import take_array_names, check_empty_arrays

SNAP_VIEW_TMPL = textwrap.dedent("""\
    Value Required,Filldown ArrayName (\S+)
    Value SPName (.+)
    Value TotalNumberOfLUNs (.+)
    Value NumberOfUnallocatedLUNs (.+)
    Value UnallocatedLUNs (.+)
    Value TotalSizeGB (.+)
    Value UnallocatedSizeGB (.+)
    Value UsedLUNPoolGB (.+)
    Value PercentUsedOfLUNPool (.+)
    Value ChunkSizeBlocks (.+)
    Value SnapViewLUNName (.+)
    Value SnapViewLUNID (.+)
    Value TargetLUN (.+)
    Value State (.+)
    Value Comments (.+)

    Start
      ^\s*#spcollect\s+${ArrayName}_SP(.+)\s*$$ -> ReservedStart
      
    ReservedStart
      ^\S+NavisecCli.exe\s+-np\s+reserved\s+-lunpool\s+-list\s*$$ -> ReservedLine

    ReservedLine
      ^\s*Name of the SP:\s*${SPName}\s*$$
      ^\s*Total Number of LUNs in Pool:\s*${TotalNumberOfLUNs}\s*$$
      ^\s*Number of Unallocated LUNs in Pool:\s*${NumberOfUnallocatedLUNs}\s*$$
      ^\s*Unallocated LUNs:\s*${UnallocatedLUNs}\s*$$
      ^\s*Total size in GB:\s*${TotalSizeGB}\s*$$
      ^\s*Unallocated size in GB:\s*${UnallocatedSizeGB}\s*$$
      ^\s*Used LUN Pool in GB:\s*${UsedLUNPoolGB}\s*$$
      ^\s*% Used of LUN Pool:\s*${PercentUsedOfLUNPool}\s*$$
      ^\s*Chunk size in disk blocks:\s*${ChunkSizeBlocks}\s* -> SnapStart
      
    SnapStart
      ^\S+NavisecCli.exe\s+-np\s+snapview\s+-listsnapshots\s*$$ -> CheckDashes

    CheckDashes
      ^\s*(-+) -> SnapLine
    
    SnapLine
      ^\s*SnapView logical unit name:\s*${SnapViewLUNName}\s*$$
      ^\s*SnapView logical unit ID:\s*${SnapViewLUNID}\s*$$
      ^\s*Target Logical Unit:\s*${TargetLUN}\s*$$
      ^\s*State:\s*${State} -> Record SnapLine
      ^\s*(\*+) -> Record Start
""")


def process(workbook: Any, content: str, array_names: dict) -> None:
    """Process SnapView worksheet

    :param workbook:
    :param content:
    :param array_names:
    """
    worksheet = workbook.get_sheet_by_name('SnapView')

    headers = get_parser_header(SNAP_VIEW_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    snap_view_out = check_empty_arrays(take_array_names(
        array_names, run_parser_over(content, SNAP_VIEW_TMPL)))

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, snap_view_out), 2):
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
        'SnapViewTable',
        'SnapView',
        final_col,
        final_row)
