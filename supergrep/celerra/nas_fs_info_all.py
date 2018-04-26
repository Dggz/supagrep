"""nas_fs_info_all Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output


NAS_FS_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required ID (.+)
    Value Name (.+)
    Value Acl (.+)
    Value InUse (.+)
    Value Type (.+)
    Value Worm (.+)
    Value Volume (.+)
    Value Pool (.+)
    Value MemberOf (.+)
    Value RWServers (.+)
    Value ROServers (.+)
    Value RWVdms (.+)
    Value ROVdms (.+)
    Value CheckptOf (.+)
    Value AutoExt (.+)
    Value Deduplication (.+)
    Value Used (.+)
    Value FullMark (.+)
    Value StorDevs (.+)
    Value Disks (.+)
    Value List DiskList (.+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> NASStart

    NASStart
      ^\s*Output from:\s+/nas/bin/nas_fs -info -all -o mpd -> DiskLines

    DiskLines
      ^\s*id\s*=\s*${ID} -> DiskLines
      ^\s*name\s*=\s*${Name}
      ^\s*acl\s*=\s*${Acl}
      ^\s*in_use\s*=\s*${InUse}
      ^\s*type\s*=\s*${Type}
      ^\s*worm\s*=\s*${Worm}
      ^\s*volume\s*=\s*${Volume}
      ^\s*pool\s*=\s*${Pool}
      ^\s*member_of\s*=\s*${MemberOf}
      ^\s*rw_servers\s*=\s*${RWServers}
      ^\s*ro_servers\s*=\s*${ROServers}
      ^\s*rw_vdms\s*=\s*${RWVdms}
      ^\s*ro_vdms\s*=\s*${ROVdms}
      ^\s*checkpt_of\s*=\s*${CheckptOf}
      ^\s*auto_ext\s*=\s*${AutoExt}
      ^\s*deduplication\s*=\s*${Deduplication}
      ^\s*used\s*=\s*${Used}
      ^\s*full(mark)\s*=\s*${FullMark}
      ^\s*stor_devs\s*=\s*${StorDevs}
      ^\s*disks\s*=\s*${Disks} -> DiskListLines
      ^\s*(\*+) -> Record Start
      
    DiskListLines
      ^\s*${DiskList} -> DiskListLines
      ^\s*$$ -> Record DiskLines
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process nas_fs_info_all worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('nas_fs_info_all')

    headers = get_parser_header(NAS_FS_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    nas_fs_out = run_parser_over(content, NAS_FS_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, nas_fs_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(column_format(col_n), row_n)]
            if isinstance(col_value, str):
                cell.value = str.strip(col_value)
            else:
                cell.alignment = Alignment(wrapText=True)
                cell.value = '\n'.join(col_value)
            style_value_cell(cell)
            if chr(col_n) not in ('B',):
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'nas_fs_info_allTable',
        'nas_fs_info_all',
        final_col,
        final_row)
