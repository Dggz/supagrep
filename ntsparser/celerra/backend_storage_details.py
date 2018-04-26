"""Backend Storage SP DETAILS Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from ntsparser.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output


BACKEND_DETAILS_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required SPIdentifier (.+)
    Value MicrocodeVersion (.+)
    Value SerialNum (.+)
    Value PromRev (.+)
    Value AgentRev (.+)
    Value PhysMemory (.+)
    Value NetworkName (.+)
    Value IPAddress (.+)
    Value SubnetMask (.+)
    Value GatewayAddress (.+)
    Value NumDiskVolumes (.+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine
      ^\s*Output from:\s+/nas/bin/nas_storage\s+-info\s+-all\s+-o\s+sync=no -> BackendDetailLines

    HostLine
      ^\s*${Hostname}\s*$$ -> NASStart

    NASStart
      ^\s*Output from:\s+/nas/bin/nas_storage\s+-info\s+-all\s+-o\s+sync=no -> BackendDetailLines

    BackendDetailLines
      ^\s*SP Identifier\s+=\s+${SPIdentifier}
      ^\s*microcode_version\s+=\s+${MicrocodeVersion}
      ^\s*serial_num\s+=\s+${SerialNum}
      ^\s*prom_rev\s+=\s+${PromRev}
      ^\s*agent_rev\s+=\s+${AgentRev}
      ^\s*phys_memory\s+=\s+${PhysMemory}
      ^\s*network_name\s+=\s+${NetworkName}
      ^\s*ip_address\s+=\s+${IPAddress}
      ^\s*subnet_mask\s+=\s+${SubnetMask}
      ^\s*gateway_address\s+=\s+${GatewayAddress}
      ^\s*num_disk_volumes\s+=\s+${NumDiskVolumes} -> Record BackendDetailLines
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Backend Storage SP DETAILS worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Backend Storage SP DETAILS')

    headers = get_parser_header(BACKEND_DETAILS_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    backend_details_out = run_parser_over(content, BACKEND_DETAILS_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, backend_details_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(column_format(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'BackendStorageSPDETAILSTable',
        'Backend Storage SP DETAILS',
        final_col,
        final_row)
