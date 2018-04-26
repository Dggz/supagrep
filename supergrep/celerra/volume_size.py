"""volume size Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output


NAS_VOLUME_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required ID (.+)
    Value Name (.+)
    Value Acl (.+)
    Value InUse (.+)
    Value Type (.+)
    Value Disks (.+)
    Value ClntVolume (.+)
    Value ClntFilesys (.+)
    Value TotalSizeMB (\S+)
    Value AvailableSizeMB (\S+)
    Value UsedSizeMB (\S+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> NASStart

    NASStart
      ^\s*Output from:\s+/nas/bin/nas_volume -info -size -all -> DiskLines

    DiskLines
      ^\s*id\s*=\s*${ID}
      ^\s*name\s*=\s*${Name}
      ^\s*acl\s*=\s*${Acl}
      ^\s*in_use\s*=\s*${InUse}
      ^\s*type\s*=\s*${Type}
      ^\s*disks\s*=\s*${Disks}
      ^\s*clnt_volume\s*=\s*${ClntVolume}
      ^\s*clnt_filesys\s*=\s*${ClntFilesys}
      ^\s*size\s*=\s+total\s+=\s+${TotalSizeMB}\s+avail\s+=\s+${AvailableSizeMB}\s+used\s+=\s+${UsedSizeMB}\s+ -> Record DiskLines
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process volume size worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('volume size')

    headers = get_parser_header(NAS_VOLUME_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    volume_size_out = run_parser_over(content, NAS_VOLUME_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, volume_size_out), 2):
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
        'VolumeSizeTable',
        'volume size',
        final_col,
        final_row)
