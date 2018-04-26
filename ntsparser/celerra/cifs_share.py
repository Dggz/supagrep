"""SMB, NFS, Multiprotocol Sheets"""
import textwrap
from collections import namedtuple
from operator import itemgetter
from typing import Any

from cytoolz import groupby
from openpyxl.styles import Alignment

from ntsparser.celerra.utils import classify_rows
from ntsparser.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from ntsparser.parsing import run_parser_over
from ntsparser.utils import sheet_process_output

SERVER_EXPORT_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required,Filldown Server (\S+)
    Value MountedPath (\S+)
    Value Required Type (\S+)
    Value Properties (.+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> ServerStart

    ServerStart
      ^\s*Output from:\s+/nas/bin/server_export\s+ -> ServerLine
      
    ServerLine
      ^\s*${Server}\s*:\s*$$ -> DataLines
      
    DataLines
      ^\s*${Type}\s+"${MountedPath}"\s+${Properties} -> Record DataLines
      ^\s*Output from:\s+/nas/bin/server_export\s+(\S+) -> ServerLine
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process SMB, NFS, Multiprotocol worksheets

    :param workbook:
    :param content:
    """
    sheets = ['SMB', 'NFS', 'Multiprotocol']

    row_tuples = dict()
    row_tuples['SMB'] = [
        'ArrayName', 'DataMover', 'ShareName', 'SharePath',
        'RootPath', 'Type', 'umask', 'maxusr', 'netbios', 'comment'
    ]

    row_tuples['NFS'] = [
        'Hostname', 'Server', 'MountedPath', 'FileSystem',
        'Type', 'rw', 'root', 'access'
    ]

    row_tuples['Multiprotocol'] = [
        'Hostname', 'Server',
        'MountedPath', 'Type'
    ]

    server_export_out = run_parser_over(content, SERVER_EXPORT_TMPL)
    server_export_grouped = groupby(itemgetter(3), server_export_out)
    share, export, multi = classify_rows(server_export_grouped)

    for sheet, data_list in zip(sheets, [share, export, multi]):
        worksheet = workbook.get_sheet_by_name(sheet)
        build_header(worksheet, row_tuples[sheet])
        RowTuple = namedtuple('RowTuple', row_tuples[sheet])

        final_col, final_row = 0, 0
        for row_n, row_tuple in enumerate(map(RowTuple._make, data_list), 2):
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
            '{}Table'.format(sheet),
            sheet,
            final_col,
            final_row)
