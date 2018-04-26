"""nas_replicate_info Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from openpyxl.styles import Alignment

from supergrep.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output


NAS_REPLICATE_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required ID (.+)
    Value Name (.+)
    Value SourceStatus (.+)
    Value NetworkStatus (.+)
    Value DestinationStatus (.+)
    Value LastSyncTime (.+)
    Value Type (.+)
    Value CelerraNetworkServer (.+)
    Value DartInterconnect (.+)
    Value PeerDartInterconnect (.+)
    Value ReplicationRole (.+)
    Value SourceFilesystem (.+)
    Value SourceDataMover (.+)
    Value SourceInterface (.+)
    Value DestinationFilesystem (.+)
    Value DestinationDataMover (.+)
    Value DestinationInterface (.+)
    Value MaxOutofSyncTimeMinutes (.+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> NASStart

    NASStart
      ^\s*Output from:\s+/nas/bin/nas_replicate -info -all -> DiskLines

    DiskLines
      ^\s*ID\s*=\s*${ID}
      ^\s*Name\s*=\s*${Name}
      ^\s*Source Status\s*=\s*${SourceStatus}
      ^\s*Network Status\s*=\s*${NetworkStatus}
      ^\s*Destination Status\s*=\s*${DestinationStatus}
      ^\s*Last Sync Time\s*=\s*${LastSyncTime}
      ^\s*Type\s*=\s*${Type}
      ^\s*Celerra Network Server\s*=\s*${CelerraNetworkServer}
      ^\s*Dart Interconnect\s*=\s*${DartInterconnect}
      ^\s*Peer Dart Interconnect\s*=\s*${PeerDartInterconnect}
      ^\s*Replication Role\s*=\s*${ReplicationRole}
      ^\s*Source Filesystem\s*=\s*${SourceFilesystem}
      ^\s*Source Data Mover\s*=\s*${SourceDataMover}
      ^\s*Source Interface\s*=\s*${SourceInterface}
      ^\s*Destination Filesystem\s*=\s*${DestinationFilesystem}
      ^\s*Destination Data Mover\s*=\s*${DestinationDataMover}
      ^\s*Destination Interface\s*=\s*${DestinationInterface}
      ^\s*Max Out of Sync Time \(minutes\)\s*=\s*${MaxOutofSyncTimeMinutes} -> Record DiskLines
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process nas_replicate_info worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('nas_replicate_info')

    headers = get_parser_header(NAS_REPLICATE_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    nas_replicate_out = run_parser_over(content, NAS_REPLICATE_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, nas_replicate_out), 2):
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
        'nas_replicate_infoTable',
        'nas_replicate_info',
        final_col,
        final_row)
