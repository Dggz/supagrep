"""NAS Summary Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from openpyxl.styles import Alignment

from ntsparser.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output

NAS_SUMMARY_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required SystemType (.+)
    Value Version (.+)
    Value NumberOfDataMovers (\d+)
    Value VDMs (\d+)
    Value List DataMovers (\S+)
    Value List DARTRelease (\S+)
    Value List Role (\S+)
    Value Filesystems (\S+|)
    Value UXFS (\S+)
    Value Checkpoint (\S+)
    Value Total (\S+)
    Value NumberOfDisksVols (.+)
    Value ReplicationMode (\S+)
    Value Replications (\S+)
    Value Copies (\S+|\s+)
    Value SessionTypes (\S+|\s+)
    Value Filesystem (\S+|\s+)
    Value VDM (\S+|\s+)
    Value iSCSINumber (\S+|\s+)
    Value States (\S+|\s+)
    Value OK (\S+|\s+)
    Value Inactive (\S+|\s+)
    Value FailedOver (\S+|\s+)
    Value SwitchedOver (\S+|\s+)
    Value Interconnects (\S+|\s+)
    Value RemoteCelerras (\S+|\s+)
    Value CIFS (\S+|\s+)
    Value iSCSI (\S+|\s+)
    Value NFS (\S+|\s+)
    Value Rep (\S+|\s+)
    Value SRDF (\S+|\s+)
    Value DHSM (\S+|\s+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$
      ^\s*Output from:\s+/nas/tools/nas_summary -> SystemLines

    SystemLines
      ^\s*System type:\s+${SystemType}
      ^\s*Version:\s+${Version} 
      ^\s*Number of data movers:\s+${NumberOfDataMovers}\s+VDMs:\s+${VDMs} -> DMLines
      
    DMLines
      ^\s*${DataMovers}:\s+(\S+)\s+${DARTRelease}\s+State:(\S+)\s+${Role}
      ^\s*(#+) -> FilesystemLines
      
    FilesystemLines
      ^\s*Filesystems:\s+${Filesystems}\s+UXFS:${UXFS}\s+Checkpoint:${Checkpoint}\s+Total:${Total}
      ^\s*Number of disk vols \(LUNs\):\s+${NumberOfDisksVols}
      ^\s*Replication mode:\s+${ReplicationMode}\s+#\s+Replications:\s+${Replications}\s+#\s+Copies:\s+${Copies}
      ^\s*Session types:${SessionTypes}\s+Filesystem:\s+${Filesystem}\s+VDM:\s+${VDM}\s+iSCSI:\s+${iSCSINumber}
      ^\s*States:${States}\s+OK:\s+${OK}\s+(\S+)\s+Inactive:\s+${Inactive}
      ^\s*Failed Over:\s+${FailedOver}\s+Switched Over:\s+${SwitchedOver}
      ^\s*# of Interconnects:\s+${Interconnects}\s+Remote Celerras:\s+${RemoteCelerras}
      ^\s*Features in use:\s+CIFS:\s+${CIFS}\s+iSCSI:\s+${iSCSI}\s+NFS:\s+${NFS}
      ^\s*Rep:\s+${Rep}\s+SRDF:\s+${SRDF}\s+DHSM:\s+${DHSM} -> Record Start
""")


def process(workbook: Any, content: str) -> None:
    """Process NAS Summary worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('NAS Summary')

    headers = get_parser_header(NAS_SUMMARY_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    nas_summary_out = run_parser_over(content, NAS_SUMMARY_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, nas_summary_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(column_format(col_n), row_n)]
            if isinstance(col_value, str):
                cell.value = str.strip(col_value)
            else:
                cell.alignment = Alignment(wrapText=True)
                cell.value = '\n'.join(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'NASSummaryTable',
        'NAS Summary',
        final_col,
        final_row)
