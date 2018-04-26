"""Volumes (3Par) Sheet"""
import textwrap
from collections import namedtuple
from operator import itemgetter
from typing import Any

from cytoolz import concat, groupby

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output

SHOWVV_TMPL = textwrap.dedent("""\
    Value Filldown Name (\S+)
    Value Filldown Serial (\S+)
    Value VVName (\S+)
    Value Required ID (\S+)
    Value Domain (\S+)
    Value Prov (\S+)
    Value Type (\S+)
    Value CopyOf (\S+)
    Value BsId (\S+)
    Value Rd (\S+)
    Value DetailedState (\S+)
    Value AdmMB (\S+)
    Value SnpMB (\S+)
    Value UsrMB (\S+)
    Value VSizeMB (\S+)

    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+(\S+\s\S+|\S+)\s+${Serial}\s+
      ^\s*(\*+) -> HostStart

    HostStart
      ^\s*Id\s+Name\s+Domain\s+Prov\s+Type\s+CopyOf\s+BsId\s+Rd\s+-Detailed_State-\s+Adm\s+Snp\s+Usr\s+VSize -> HostLines

    HostLines
      ^\s*${ID}\s+${VVName}\s+${Domain}\s+${Prov}\s+${Type}\s+${CopyOf}\s+${BsId}\s+${Rd}\s+${DetailedState}\s+${AdmMB}\s+${SnpMB}\s+${UsrMB}\s+${VSizeMB} -> Record HostLines
      ^\s*(-+) -> Start
      ^\s*(\*+) -> Start
""")


SHOWVV_CPG_TMPL = textwrap.dedent("""\
    Value Filldown Name (\S+)
    Value Filldown Serial (\S+)
    Value VVName (\S+)
    Value Required ID (\S+)
    Value UsrCPG (\S+)
    Value SnpCPG (\S+)

    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+(\S+\s\S+|\S+)\s+${Serial}\s+
      ^\s*(\*+) -> HostStart

    HostStart
      ^\s*Id\s+Name\s+Domain\s+Prov\s+Type\s+UsrCPG\s+SnpCPG -> HostLines

    HostLines
      ^\s*${ID}\s+${VVName}\s+(\S+)\s+(\S+)\s+(\S+)\s+${UsrCPG}\s+${SnpCPG} -> Record HostLines
      ^\s*(-+) -> Start
      ^\s*(\*+) -> Start
""")


SHOWVLUN_TMPL = textwrap.dedent("""\
    Value Filldown Name (\S+)
    Value Filldown Serial (\S+)
    Value Required VVName (\S+)
    Value HostName (\S+)

    Start
      ^\s*ID\s+(-*)Name(-*)\s+(-*)Model(-*)\s+(-*)Serial(-*)\s+Nodes -> SysLine

    SysLine
      ^\s*(\S+)\s+${Name}\s+(\S+\s\S+|\S+)\s+${Serial}\s+
      ^\s*(\*+) -> VolumesStart

    VolumesStart
      ^\s*Domain\s+Lun\s+VVName\s+HostName\s+-Host_WWN/iSCSI_Name-\s+Port\s+Type -> VolumesLines

    VolumesLines
      ^\s*(\S+)\s+(\S+)\s+${VVName}\s+${HostName}\s+(\S+)\s+(\S+)\s+(\S+) -> Record VolumesLines
      ^\s*(-+) -> Start
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Volumes (3Par) worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Volumes')

    headers = list(concat([
        get_parser_header(SHOWVV_TMPL),
        get_parser_header(SHOWVV_CPG_TMPL)[4:],
        get_parser_header(SHOWVLUN_TMPL)[3:],
    ]))

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name
    build_header(worksheet, headers)

    show_vv_out = groupby(
        itemgetter(0, 1, 2), run_parser_over(content, SHOWVV_TMPL))
    show_vv_cpg_out = groupby(
        itemgetter(0, 1, 2), run_parser_over(content, SHOWVV_CPG_TMPL))
    showv_lun_out = groupby(
        itemgetter(0, 1, 2), run_parser_over(content, SHOWVLUN_TMPL))

    rows = []  # type: list
    for idfier in show_vv_out:
        lun_out = showv_lun_out[idfier][0][3:] if \
            idfier in showv_lun_out.keys() else ['']
        for idx, entry in enumerate(show_vv_out[idfier]):
            rows.append(
                entry + show_vv_cpg_out[idfier][idx][4:] + lun_out)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
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
        'VolumesTable',
        'Volumes',
        final_col,
        final_row)
