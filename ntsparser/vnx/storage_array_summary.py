"""Storage Array Summary Sheet"""
import textwrap
from collections import namedtuple, defaultdict
from operator import itemgetter
from typing import Any

from cytoolz.curried import concat, unique, compose, groupby

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output
from ntsparser.vnx.utils import check_empty_arrays

ARRAY_NAME_TMPL = textwrap.dedent("""\
    Value Required ArrayName (\S+)

    Start
      ^\S+NavisecCli.exe\s+-np\s+arrayname -> Record ArrayLine

    ArrayLine
      ^\s*Array Name:\s+${ArrayName} -> Start
""")

GET_ARRAY_UID_TMPL = textwrap.dedent("""\
    Value Required Hostname (\S+)
    Value Required ArrayUID (\S+)

    Start
      ^\S+NavisecCli.exe\s+-np\s+getarrayuid -> Header
      
    Header
      ^.*Hostname -> HeaderLine
      
    HeaderLine
      ^.*---- -> ArrayLine

    ArrayLine
      ^\s*${Hostname}\s+${ArrayUID} -> Record
      ^.*---- -> Start
""")

GET_AGENT_TMPL = textwrap.dedent("""\
    Value AgentRev (.+)
    Value Name (.+)
    Value Desc (.+)
    Value Required Node (.+)
    Value PhysicalNode (.+)
    Value Signature (.+)
    Value PeerSignature (.+)
    Value Revision (.+)
    Value SCSIId (.+)
    Value Model (.+)
    Value ModelType (.+)
    Value PromRev (.+)
    Value SPMemory (.+)
    Value SerialNo (.+)
    Value SPIdentifier (.+)
    Value Cabinet (.+) 

    Start
      ^\S+NavisecCli.exe\s+-np\s+getagent -> Record ArrayLine

    ArrayLine
      ^\s*Agent Rev:\s+${AgentRev}
      ^\s*Name:\s+${Name}
      ^\s*Desc:\s+${Desc}
      ^\s*Node:\s+${Node}
      ^\s*Physical Node:\s+${PhysicalNode}
      ^\s*Signature:\s+${Signature}
      ^\s*Peer Signature:\s+${PeerSignature}
      ^\s*Revision:\s+${Revision}
      ^\s*SCSI Id:\s+${SCSIId}
      ^\s*Model:\s+${Model}
      ^\s*Model Type:\s+${ModelType}
      ^\s*Prom Rev:\s+${PromRev}
      ^\s*SP Memory:\s+${SPMemory}
      ^\s*Serial No:\s+${SerialNo}
      ^\s*SP Identifier:\s+${SPIdentifier}
      ^\s*Cabinet:\s+${Cabinet} -> Start
""")


def process(workbook: Any, content: str) -> tuple:
    """Process Storage-Array-Summary worksheet

    Also returns a list of array names used in other sheets

    :param workbook:
    :param content:
    :return:
    """
    worksheet = workbook.get_sheet_by_name('Storage-Array-Summary')

    headers = list(concat([
        get_parser_header(ARRAY_NAME_TMPL),
        get_parser_header(GET_ARRAY_UID_TMPL),
        get_parser_header(GET_AGENT_TMPL)
    ]))
    RowTuple = namedtuple('RowTuple', headers)   # pylint: disable=invalid-name

    build_header(worksheet, headers)

    cmd_arrayname_out = run_parser_over(content, ARRAY_NAME_TMPL)
    cmd_getarrayuid_out = run_parser_over(content, GET_ARRAY_UID_TMPL)
    cmd_getagent_out = run_parser_over(content, GET_AGENT_TMPL)

    # noinspection PyTypeChecker
    cmd_out = map(compose(
        list,
        concat),
        zip(
            cmd_arrayname_out,
            cmd_getarrayuid_out,
            cmd_getagent_out))

    array_names = defaultdict(str)    # type: defaultdict
    rows = check_empty_arrays(list(unique(cmd_out, key=itemgetter(0, 1))))
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        array_names[worksheet['Q{}'.format(row_n)].value] = \
            worksheet['A{}'.format(row_n)].value
        final_row = row_n

    sheet_process_output(
        worksheet,
        'StorageArraySummaryTable',
        'Storage-Array-Summary',
        final_col,
        final_row)

    array_models = groupby(itemgetter(12), rows)
    array_revisions = groupby(itemgetter(10), rows)
    return array_names, array_models, array_revisions
