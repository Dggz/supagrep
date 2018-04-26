"""LUNs sheet"""
import textwrap
from collections import namedtuple
from operator import itemgetter
from typing import Any

from cytoolz.curried import unique

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output
from supergrep.vnx.utils import (
    capacity_conversion, check_empty_arrays, get_luns)

GETLUN_TMPL = textwrap.dedent("""\
    Value Filldown ArrayName (\S+)
    Value ArrayLogicalUnitNumber (\S+)
    Value Name (.+)
    Value StorageGroup (.+)
    Value HostLogicalUnitNumber (.+)
    Value RAIDType (.+)
    Value RAIDGroupID (.+)
    Value State (.+)
    Value ElementSize (.+) 
    Value CurrentOwner (.+)
    Value DefaultOwner (.+)
    Value LUNCapacityMB (.+)
    Value LUNCapacityTB (.+)
    Value LUNCapacityBlocks (.+)
    Value UID (.+)
    Value IsPrivate (.+)
    Value SnapshotsList (.+)
    Value MirrorViewName (.+)
    Value Comments (.+)

    Start
      ^\S+NavisecCli.exe\s+-np\s+arrayname -> ArrayNameLine

    ArrayNameLine
      ^\s*Array Name:\s+${ArrayName} -> LunStart 

    LunStart 
      ^\S+NavisecCli.exe\s+-np\s+getlun$$ -> LunLine

    LunLine
      ^\s*LOGICAL UNIT NUMBER\s+${ArrayLogicalUnitNumber}
      ^\s*Name\s+${Name}
      ^\s*RAID Type:\s+${RAIDType}
      ^\s*RAIDGroup ID:\s+${RAIDGroupID}
      ^\s*State:\s+${State}
      ^\s*Element Size:\s+${ElementSize}
      ^\s*Current owner:\s+${CurrentOwner}
      ^\s*Default Owner:\s+${DefaultOwner}
      ^\s*LUN Capacity\(Megabytes\):\s+${LUNCapacityMB}
      ^\s*LUN Capacity\(Blocks\):\s+${LUNCapacityBlocks}
      ^\s*UID:\s+${UID}
      ^\s*Is Private:\s+${IsPrivate}
      ^\s*Snapshots List:\s+${SnapshotsList}
      ^\s*MirrorView Name if any:\s+${MirrorViewName} -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str, sg_data: list) -> list:
    """Process LUNs worksheet

    :param workbook:
    :param content:
    :param sg_data:
    """
    worksheet_name = 'LUNs'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(GETLUN_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)

    cmd_getlun_out = run_parser_over(content, GETLUN_TMPL)
    cmd_getlun_out = check_empty_arrays(
        list(unique(cmd_getlun_out, key=itemgetter(0, 1))))

    expanded_luns = [[*entry[:-1], get_luns(entry[-1])] for entry in sg_data]

    sg_dict = {}  # type: dict
    for entry in expanded_luns:
        for lun in entry[-1]:
            sg_dict[entry[0], lun[1]] = entry[1], lun[0]

    for row in cmd_getlun_out:
        if (row[0], row[1]) in sg_dict.keys():
            row[3], row[4] = sg_dict[row[0], row[1]][0], \
                             sg_dict[row[0], row[1]][1]
        else:
            row[3], row[4] = ('No Storage Group Found', '')
        row[12] = capacity_conversion(row[11])

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, cmd_getlun_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            if chr(col_n) == 'K':
                set_cell_to_number(cell, '0.00000')
            else:
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'LUNSTable',
        'LUNs',
        final_col,
        final_row)

    return cmd_getlun_out
