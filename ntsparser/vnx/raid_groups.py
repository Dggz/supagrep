"""RAID-Groups sheet"""
import textwrap
from collections import namedtuple
from operator import itemgetter
from typing import Any

from cytoolz.curried import unique
from openpyxl.styles import Alignment

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output
from ntsparser.vnx.utils import capacity_conversion, check_empty_arrays

GETRG_TMPL = textwrap.dedent("""\
    Value Required,Filldown ArrayName (\S+)
    Value RaidGroupID (.+)
    Value RaidGroupType (.+)
    Value List Disks (Bus.+)                                            
    Value List Luns (.+)
    Value MaxNumberOfDisks (.+)
    Value MaxNumberOfLuns (.+)
    Value RawCapacityBlocks (.+)
    Value LogicalCapacityBlocks (.+)
    Value RawCapacityTB (.+)
    Value FreeCapacityBlocksNonContiguous (.+)
    Value FreeContiguousGroupOfUnboundSegments (.+)
    Value LegalRAIDTypes (.+)
    Value Comments (.+)

    Start
      ^\S+NavisecCli.exe\s+-np\s+arrayname -> ArrayNameLine

    ArrayNameLine
      ^\s*Array Name:\s+${ArrayName} -> RAIDStart

    RAIDStart
      ^\S+NavisecCli.exe\s+-np\s+getrg\s*$$ -> CheckDashes
      
    CheckDashes
      ^\s*(-+) -> CheckEnd
      
    CheckEnd
      ^\s*(-+) -> Record Start
      ^\s*RaidGroup ID:\s+${RaidGroupID} -> RAIDLine
      
    RAIDLine
      ^\s*RaidGroup ID:\s+${RaidGroupID}
      ^\s*RaidGroup Type:\s+${RaidGroupType}
      ^\s*List of disks:\s+${Disks}
      ^\s+${Disks}
      ^\s*List of luns:\s+${Luns}
      ^\s*Max Number of disks:\s+${MaxNumberOfDisks}
      ^\s*Max Number of luns:\s+${MaxNumberOfLuns}
      ^\s*Raw Capacity \(Blocks\):\s+${RawCapacityBlocks}
      ^\s*Logical Capacity \(Blocks\):\s+${LogicalCapacityBlocks}
      ^\s*Free Capacity \(Blocks,non-contiguous\):\s+${FreeCapacityBlocksNonContiguous}
      ^\s*Free contiguous group of unbound segments:\s+${FreeContiguousGroupOfUnboundSegments}
      ^\s*Legal RAID types:\s+${LegalRAIDTypes} -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process RAID-Groups worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'RAID-Groups'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(GETRG_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)

    cmd_getrg_out = run_parser_over(content, GETRG_TMPL)
    cmd_getrg_out = check_empty_arrays(
        list(unique(cmd_getrg_out, key=itemgetter(0, 1))))

    for row in cmd_getrg_out:
        row[9] = capacity_conversion(row[8], conversion_factor=2147483648)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, cmd_getrg_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            if isinstance(col_value, str):
                cell.value = str.strip(col_value)
            else:
                cell.alignment = Alignment(wrapText=True)
                cell.value = str.strip('\n'.join(col_value))
            style_value_cell(cell)
            if chr(col_n) == 'J':
                set_cell_to_number(cell, '0.00000')
            else:
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'RaidGroupsTable',
        'RAID-Groups',
        final_col,
        final_row)
