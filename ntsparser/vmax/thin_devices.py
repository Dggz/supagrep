"""thin_devices parser"""

import textwrap
from collections import namedtuple
from typing import Any

from cytoolz.curried import map
from openpyxl.comments import Comment
from openpyxl.styles import Alignment

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output

legend = textwrap.dedent("""
    Legend:
     Flags:  (E)mulation : A = AS400, F = FBA, 8 = CKD3380, 9 = CKD3390
             (S)hared Tracks : S = Shared Tracks Present, . = No Shared Tracks
             (P)ersistent Allocs : A = All, S = Some, . = None
             S(T)atus    : B = Bound, I = Binding, U = Unbinding, A = Allocating,
                           D = Deallocating, R = Reclaiming, C = Compressing,
                           N = Uncompressing, F = FreeingAll, . = Unbound
""")


THNDEV_TMPL = textwrap.dedent("""\
    Value Filldown,Required SymmetrixID (\d+)
    Value Filldown,Required Sym (.....|....)
    Value Required BoundPoolName (\w+|-)
    Value FlagsESPT (....)
    Value Filldown TotalGB (\d+\.\d+|-)
    Value Filldown PoolSubPercent (\d+|-)
    Value PoolAllocatedGB (\S+)
    Value PoolAllocatedPercent (\d+)
    Value Filldown TotalWrittenGB (\d+\.\d+|-)
    Value Filldown TotalWrittenPercent (\d+|-)
    Value Filldown ExclusiveAllocatedGB (\S+|-)
    Value Filldown CompressionSizeGB (\d+\.\d+)
    Value CompressionRatioPercent (\S+)
    
    Start
      ^\s*Symmetrix ID:\s*${SymmetrixID}
      ^\s*(-+)\s+(-+) -> SymStart
      
    SymStart
      ^\s*${Sym}\s+${BoundPoolName}\s+${FlagsESPT}\s+${TotalGB}\s+${PoolSubPercent}\s+${PoolAllocatedGB}\s+${PoolAllocatedPercent}\s+${TotalWrittenGB}\s+${TotalWrittenPercent}\s+${CompressionSizeGB}\s+${CompressionRatioPercent}\s*$$ -> Record 
      ^\s*${BoundPoolName}\s+(-\.--)\s+\s+-\s+\s+-\s+\s+${PoolAllocatedGB}\s+${PoolAllocatedPercent}\s+-\s+-\s+${CompressionSizeGB}\s+${CompressionRatioPercent}\s*$$ -> Record
      ^\s*${Sym}\s+${BoundPoolName}\s+${FlagsESPT}\s+${TotalGB}\s+${PoolSubPercent}\s+${PoolAllocatedGB}\s+${PoolAllocatedPercent}\s+${ExclusiveAllocatedGB}\s+${CompressionRatioPercent}\s*$$ -> Record 
      ^\s*${BoundPoolName}\s+(-\.--)\s+\s+-\s+\s+-\s+\s+${PoolAllocatedGB}\s+${PoolAllocatedPercent}\s+-\s+-\s+${CompressionRatioPercent}\s*$$ -> Record
      ^\s*Total -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process thin_devices worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'thin_devices'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(THNDEV_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)
    worksheet['D1'].comment = Comment(legend, '')

    thin_devices_out = run_parser_over(content, THNDEV_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, thin_devices_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            if isinstance(col_value, str):
                cell.value = str.strip(col_value)
            else:
                cell.alignment = Alignment(wrapText=True)
                cell.value = '\n'.join(col_value)
            style_value_cell(cell)
            if chr(col_n) != 'B':
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'ThinDevTable',
        'thin_devices',
        final_col,
        final_row)
