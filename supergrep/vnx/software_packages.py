"""Software Packages sheet"""
import textwrap
from collections import namedtuple
from operator import itemgetter
from typing import Any

from cytoolz.curried import unique

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output
from supergrep.vnx.utils import check_empty_arrays

NDU_TMPL = textwrap.dedent("""\
    Value Filldown ArrayName (\S+)
    Value SoftwarePackage (.+) 
    Value Revision (.+) 
    Value CommitRequired (.+) 
    Value RevertPossible (.+) 
    Value ActiveState (.+)
    Value IsInstallationCompleted (.+)
    Value IsSystemSoftware (.+)
    Value Comments (.+)
    
    Start
      ^\S+NavisecCli.exe\s+-np\s+arrayname -> Record ArrayNameLine

    ArrayNameLine
      ^\s*Array Name:\s+${ArrayName} -> SoftwarePackageStart
      
    SoftwarePackageStart
      ^\S+NavisecCli.exe\s+-np\s+ndu\s+-messner\s+-list -> SoftwarePackageLine
      
    SoftwarePackageLine      
      ^\s*Name of the software package:\s+${SoftwarePackage}
      ^\s*Revision of the software package:\s+${Revision}
      ^\s*Commit Required:\s+${CommitRequired}
      ^\s*Revert Possible:\s+${RevertPossible}
      ^\s*Active State:\s+${ActiveState}
      ^\s*Is installation completed:\s+${IsInstallationCompleted}
      ^\s*Is this System Software:\s+${IsSystemSoftware} -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Software-Packages worksheet

    :param workbook:
    :param content:
    """
    worksheet_name = 'Software-Packages'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = get_parser_header(NDU_TMPL)
    RowTuple = namedtuple('RowTuple', headers)

    build_header(worksheet, headers)

    cmd_ndu_out = run_parser_over(content, NDU_TMPL)
    cmd_ndu_out = check_empty_arrays(
        list(unique(cmd_ndu_out, key=itemgetter(0, 1))))
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, cmd_ndu_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            if cell.value != '-':
                cell.value = str.strip(col_value, '-')
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'SoftwarePackagesTable',
        'Software-Packages',
        final_col,
        final_row)
