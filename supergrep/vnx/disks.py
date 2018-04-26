"""Disks sheet"""
import textwrap
from collections import namedtuple
from operator import itemgetter
from typing import Any

from cytoolz.curried import unique
from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output
from supergrep.vnx.utils import capacity_conversion, check_empty_arrays

GETDISK_TMPL = textwrap.dedent("""\
    Value Filldown ArrayName (\S+)
    Value Bus (\d+)
    Value Enclosure (\d+)
    Value Disk (\d+)
    Value VendorId (.+)
    Value ProductId (.+)
    Value Capacity (.+)
    Value CapacityTB (.+)
    Value Lun (.+)
    Value Type (.+)
    Value State (.+)
    Value HotSpare (.+)
    Value Private (.+)
    Value NumberOfLuns (.+)
    Value RaidGroupID (.+)
    Value DriveType (.+)
    Value CurrentSpeed (.+)
    Value MaximumSpeed (.+)
    Value Comment (----)

    Start
      ^\S+NavisecCli.exe\s+-np\s+arrayname -> Record ArrayNameLine

    ArrayNameLine
      ^\s*Array Name:\s+${ArrayName} -> DisksStart

    DisksStart
      ^\S+NavisecCli.exe\s+-np\s+getdisk\s*$$ -> CheckDashes
      
    CheckDashes
      ^\s*(-+) -> CheckEnd
      
    CheckEnd
      ^\s*(-+)\s*$$ -> Record Start
      ^\s*(\*+) -> Start
      ^\s*Bus\s+${Bus}\s+Enclosure\s+${Enclosure}\s+Disk\s+${Disk} -> LinesStart
      
    LinesStart
      ^\s*Vendor Id:\s+${VendorId} -> FullLines
      ^\s*State:\s+${State} -> Record CheckEnd
      
    FullLines
      ^\s*Product Id:\s+${ProductId}
      ^\s*Product Revision:\s
      ^\s*Lun:+${Lun}*$$
      ^\s*Type:\s+${Type}
      ^\s*State:\s+${State}
      ^\s*Hot Spare:\s+${HotSpare}
      ^\s*Prct Rebuilt:\s+
      ^\s*Capacity:\s+${Capacity}
      ^\s*Private:\s+${Private}
      ^\s*Number of Luns:\s+${NumberOfLuns}
      ^\s*Raid Group ID:\s+${RaidGroupID}
      ^\s*Drive Type:\s+${DriveType}
      ^\s*Current Speed:\s+${CurrentSpeed}
      ^\s*Maximum Speed:\s+${MaximumSpeed} -> Record CheckEnd
      ^\s*(\*+) -> Start
""")


# noinspection TaskProblemsInspection
def process(workbook: Any, content: str) -> list:
    """Process Disks worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'Disks'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(GETDISK_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)

    cmd_disks_out = run_parser_over(content, GETDISK_TMPL)
    cmd_disks_out = check_empty_arrays(
        list(unique(cmd_disks_out, key=itemgetter(0, 1, 2, 3))))

    for row in cmd_disks_out:
        row[7] = capacity_conversion(row[6])

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, cmd_disks_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            if isinstance(col_value, str):
                cell.value = str.strip(col_value)
            else:
                cell.alignment = Alignment(wrapText=True)
                cell.value = str.strip('\n'.join(col_value))
            style_value_cell(cell)
            if chr(col_n) == 'H':
                set_cell_to_number(cell, '0.00000')
            else:
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'DisksTable',
        'Disks',
        final_col,
        final_row)

    return cmd_disks_out
