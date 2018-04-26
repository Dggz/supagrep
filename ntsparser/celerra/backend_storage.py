"""Backend Storage Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from ntsparser.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output


BACKEND_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required ID (.+)
    Value SerialNumber (.+)
    Value ArrayName (.+)
    Value Name (.+)
    Value Type (.+)
    Value Ident (.+)
    Value ModelType (.+)
    Value ModeNum (.+)
    Value NumDisks (.+)
    Value NumDevs (.+)
    Value NumPdevs (.+)
    Value NumStorageGrps (.+)
    Value NumRaidGrps (.+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> NASStart

    NASStart
      ^\s*Output from:\s+/nas/bin/nas_storage\s+-info\s+-all\s+-o\s+sync=no -> BackendLines

    BackendLines
      ^\s*id\s+=${ID}
      ^\s*serial_number\s+=\s+${SerialNumber}
      ^\s*arrayname\s+=${ArrayName}
      ^\s*name\s+=${Name}
      ^\s*type\s+=${Type}
      ^\s*ident\s+=\s+${Ident}
      ^\s*model_type\s+=${ModelType}
      ^\s*model_num\s+=${ModeNum}
      ^\s*num_disks\s+=${NumDisks}
      ^\s*num_devs\s+=${NumDevs}
      ^\s*num_pdevs\s+=${NumPdevs}
      ^\s*num_storage_grps\s+=${NumStorageGrps}
      ^\s*num_raid_grps\s+=${NumRaidGrps} -> Record
      ^\s*cache_page_size\s+=(.+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Backend Storage worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Backend Storage')

    headers = get_parser_header(BACKEND_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    backend_storage_out = run_parser_over(content, BACKEND_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, backend_storage_out), 2):
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
        'BackendStorageTable',
        'Backend Storage',
        final_col,
        final_row)
