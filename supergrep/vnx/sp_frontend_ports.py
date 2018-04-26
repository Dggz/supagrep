"""SP Frontend Ports Sheet"""
import textwrap
from collections import namedtuple
from operator import itemgetter

from cytoolz.curried import (
    compose, concat, drop, first, join, juxt, map, second, unique)
from openpyxl import Workbook

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import run_parser_over, get_parser_header
from supergrep.utils import sheet_process_output
from supergrep.vnx.utils import check_empty_arrays

PORT_TMPL = textwrap.dedent("""\
    Value Filldown ArrayName (\S+)
    Value Required SPName (.+)
    Value Required SPPortID (.+)
    Value Required RegisteredInitiators (.+)
    Value Required LoggedInInitiators (.+)
    Value Required NotLoggedInInitiators (.+)
    
    Start
      ^\S+NavisecCli.exe\s+-np\s+arrayname -> Record ArrayNameLine

    ArrayNameLine
      ^\s*Array Name:\s+${ArrayName} -> PortStart
      
    PortStart
      ^\S+NavisecCli.exe\s+-np\s+port\s+-messner\s+-list\s+-all -> PortLine

    PortLine
      ^\s*SP Name:\s+${SPName}
      ^\s*SP Port ID:\s+${SPPortID}
      ^\s*Registered Initiators:\s+${RegisteredInitiators}
      ^\s*Logged-In Initiators:\s+${LoggedInInitiators}
      ^\s*Not Logged-In Initiators:\s+${NotLoggedInInitiators} -> Record
      ^\s*Information about each HBA -> Start
""")

SPPORTSPEED_TMPL = textwrap.dedent("""\
    Value Filldown ArrayName (\S+)
    Value Required StorageProcessor (.+)
    Value Required PortID (.+)
    Value Required SpeedValue (.+)
    Value Required ConnectionType (.+)
    Value Comments (.+)

    Start
      ^\S+NavisecCli.exe\s+-np\s+arrayname -> Record ArrayNameLine

    ArrayNameLine
      ^\s*Array Name:\s+${ArrayName} -> SPPortSpeedStart
      
    SPPortSpeedStart
      ^\S+NavisecCli.exe\s+-np\s+spportspeed\s+-get\s+-type -> Line

    Line
      ^\s*Storage Processor :\s+${StorageProcessor}
      ^\s*Port ID :\s+${PortID}
      ^\s*Speed Value :\s+${SpeedValue}
      ^\s*Connection Type:\s+${ConnectionType} -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Workbook, content: str) -> list:
    """Process SP-Frontend-Ports worksheet

    :param workbook:
    :param content:
    :return:
    """
    worksheet_name = 'SP-Frontend-Ports'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = list(concat([
        get_parser_header(PORT_TMPL),
        get_parser_header(SPPORTSPEED_TMPL)[3:],
    ]))
    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    cmd_port_out = run_parser_over(content, PORT_TMPL)
    cmd_spportspeed_out = run_parser_over(content, SPPORTSPEED_TMPL)

    common_columns = (0, 1, 2)
    common_columns_getter = itemgetter(*common_columns)
    cmd_merged_out = join(
        common_columns_getter, cmd_port_out,
        common_columns_getter, cmd_spportspeed_out)

    rows = map(compose(
        list,
        concat,
        juxt(
            first,
            compose(
                drop(3),
                second)))
    )(cmd_merged_out)
    rows = check_empty_arrays(list(unique(rows, key=common_columns_getter)))

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'SPFrontendPortsTable',
        'SP-Frontend-Ports',
        final_col,
        final_row)

    return rows
