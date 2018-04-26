"""MirrorView-A Sheet"""
import textwrap
from collections import namedtuple
from operator import itemgetter
from typing import Any

from cytoolz.curried import unique

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output
from ntsparser.vnx.utils import take_array_names, check_empty_arrays

MIRROR_VIEW_A_TMPL = textwrap.dedent("""\
    Value Required ArrayName (\S+)
    Value CanMirrorBeCreated (.+)
    Value MaxNumberOfRemoteImages (.+)
    Value LUNsMirroredPrimary (.+)
    Value LUNsMirroredSecondary (.+)
    Value RemoteSystemsEnabled (.+)
    Value MaxNumberOfMirrors (.+)
    Value MaxNumberOfGroups (.+)
    Value MaxNumberOfMirrorsPerGroup (.+)
    Value Comments (.+)

    Start
      ^\s*#spcollect\s+${ArrayName}_SP(.+) -> MirrorStart

    MirrorStart
      ^\S+NavisecCli.exe\s+-np\s+mirror\s+-async\s+-info\s*$$ -> CheckDashes
      
    CheckDashes
      ^\s*(-+) -> MirrorLine

    MirrorLine
      ^\s*(-+) -> Record Start
      ^\s*Can a mirror be created on this system:\s*${CanMirrorBeCreated}
      ^\s*Maximum number of Remote Images:\s*${MaxNumberOfRemoteImages}
      ^\s*Logical Units that are mirrored in Primary Images:\s*${LUNsMirroredPrimary}
      ^\s*Logical Units that are mirrored in Secondary Images:\s*${LUNsMirroredSecondary}
      ^\s*Remote systems that are enabled for mirroring:\s*${RemoteSystemsEnabled}
      ^\s*Maximum number of possible Mirrors:\s*${MaxNumberOfMirrors} -> MirrorListStart

    MirrorListStart
      ^\S+NavisecCli.exe\s+-np\s+mirror\s+-async\s+-listgroups\s*$$ -> CheckListDashes
    
    CheckListDashes
      ^\s*(-+) -> MirrorListLine

    MirrorListLine
      ^\s*(-+) -> Record Start
      ^\s*Maximum Number of Groups Allowed:${MaxNumberOfGroups}
      ^\s*Maximum Number of Mirrors per Group:${MaxNumberOfMirrorsPerGroup} -> Record Start
""")


def process(workbook: Any, content: str, array_names: dict) -> None:
    """Process MirrorView-A worksheet

    :param workbook:
    :param content:
    :param array_names:
    """
    worksheet = workbook.get_sheet_by_name('MirrorView-A')

    headers = get_parser_header(MIRROR_VIEW_A_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    cmd_mirror_view_a_out = check_empty_arrays(take_array_names(
        array_names, run_parser_over(content, MIRROR_VIEW_A_TMPL)))

    rows = unique(cmd_mirror_view_a_out, key=itemgetter(0))
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
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
        'MirrorViewATable',
        'MirrorView-A',
        final_col,
        final_row)
