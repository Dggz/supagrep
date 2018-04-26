"""SnapClones Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output
from ntsparser.vnx.utils import take_array_names, check_empty_arrays

SNAP_CLONES_TMPL = textwrap.dedent("""\
    Value Required,Filldown ArrayName (\S+)
    Value Name (.+)
    Value CloneGroupUid (.+)
    Value InSync (.+)
    Value Description (.+)  
    Value QuiesceThreshold (.+)
    Value SourceMediaFailure (.+)
    Value IsControllingSP (.+)
    Value SourceLUNSize (.+)
    Value CloneCount (.+)
    Value Sources (.+)
    Value Clones (.+)  
    Value CloneID (.+)
    Value CloneState (.+)
    Value CloneCondition (.+)
    Value AvailableForIO (.+)
    Value CloneMediaFailure (.+)
    Value IsDirty (.+)
    Value PercentSynced (.+)
    Value RecoveryPolicy (.+)
    Value SyncRate (.+)
    Value CloneLUNs (.+)
    Value UseProtectedRestore (.+)
    Value IsFractured (.+)
    Value Comments (.+)

    Start
      ^\s*#spcollect\s+${ArrayName}_SP(.+) -> CloneStart
      
    CloneStart
      ^\s*Can't run  snapview -listclonegroup\s*$$ -> Record Start
      ^\S+NavisecCli.exe\s+-np\s+snapview\s+-listclonegroup\s*$$ -> CloneSeparator
      
    CloneSeparator 
      ^\s*(-+) -> CheckEnd

    CheckEnd
      ^\s*(-+) -> Record Start
      ^\s*Name:\s*${Name} -> CloneLine

    CloneLine
      ^\s*(-+) -> Record Start
      ^\s*Name:\s*${Name}
      ^\s*CloneGroupUid:\s*${CloneGroupUid}
      ^\s*InSync:\s*${InSync}
      ^\s*Description:\s*${Description}
      ^\s*QuiesceThreshold:\s*${QuiesceThreshold}
      ^\s*SourceMediaFailure:\s*${SourceMediaFailure}
      ^\s*IsControllingSP:\s*${IsControllingSP}
      ^\s*SourceLUNSize:\s*${SourceLUNSize}
      ^\s*CloneCount:\s*${CloneCount}
      ^\s*Sources:\s*${Sources}
      ^\s*Clones:\s*${Clones}
      ^\s*CloneID:\s*${CloneID}
      ^\s*CloneState:\s*${CloneState}
      ^\s*CloneCondition:\s*${CloneCondition}
      ^\s*AvailableForIO:\s*${AvailableForIO}
      ^\s*CloneMediaFailure:\s*${CloneMediaFailure}
      ^\s*IsDirty:\s*${IsDirty}
      ^\s*PercentSynced:\s*${PercentSynced}
      ^\s*RecoveryPolicy:\s*${RecoveryPolicy}
      ^\s*SyncRate:\s*${SyncRate}
      ^\s*CloneLUNs:\s*${CloneLUNs}
      ^\s*UseProtectedRestore:\s*${UseProtectedRestore}
      ^\s*IsFractured:\s*${IsFractured} -> Record CloneLine
""")


def process(workbook: Any, content: str, array_names: dict) -> None:
    """Process SnapClones worksheet

    :param workbook:
    :param content:
    :param array_names:
    """
    worksheet = workbook.get_sheet_by_name('SnapClones')

    headers = get_parser_header(SNAP_CLONES_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    snap_clones_out = check_empty_arrays(take_array_names(
        array_names, run_parser_over(content, SNAP_CLONES_TMPL)))

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, snap_clones_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            if chr(col_n) != 'M':
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'SnapClonesTable',
        'SnapClones',
        final_col,
        final_row)
