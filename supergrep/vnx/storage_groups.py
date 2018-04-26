"""Storage-Groups sheet"""
# pylint: disable=no-name-in-module, redefined-builtin
# pylint: disable=anomalous-backslash-in-string, too-many-locals

import textwrap
from collections import namedtuple
from fnmatch import fnmatch
from operator import itemgetter
from typing import Any

from cytoolz.curried import (
    compose, concat, drop, first, groupby,
    join, juxt, last, map, second, valmap, unique)
from openpyxl.styles import Alignment  # pylint: disable=import-error

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.parsing import run_parser_over, get_parser_header
from supergrep.utils import sheet_process_output
from supergrep.vnx.utils import check_empty_arrays

STORAGEGROUP_TMPL = textwrap.dedent("""\
    Value Filldown ArrayName (\S+)
    Value StorageGroupName (.+)
    Value ServerName (.+)
    Value List HBAUIDandPorts (\S+\s+\S\S \S\s+\d+)
    Value List HLUALUPair (\d+\s+\d+)
    Value Shareable (\S+)
    Value StorageGroupUID (\S+)
    Value Comments (.+)

    Start
      ^\S+NavisecCli.exe\s+-np\s+arrayname -> ArrayNameLine

    ArrayNameLine
      ^\s*Array Name:\s+${ArrayName} -> StorageGroupStart

    StorageGroupStart
      ^\S+NavisecCli.exe\s+-np\s+storagegroup\s+-list -> StorageGroupLine

    StorageGroupLine
      ^\s*Storage Group Name:\s+${StorageGroupName}
      ^\s*Storage Group UID:\s+${StorageGroupUID}
      ^\s*HBA UID\s+SP Name\s+SPPort\s*$$ -> HBALines
      ^\s*HLU/ALU Pairs:\s* -> HLULine
      ^\s*Shareable:\s+${Shareable} -> Record StorageGroupLine
      ^\s*(\*+) -> Start
      
    HBALines
      ^\s*${HBAUIDandPorts}\s*$$
      ^\s*HLU/ALU Pairs:\s* -> HLULine
      ^\s*Shareable:\s+${Shareable} -> Record StorageGroupLine

    HLULine
      ^\s*${HLUALUPair}\s*$$
      ^\s*Shareable:\s+${Shareable} -> Record StorageGroupLine
""")

PORT_TMPL = textwrap.dedent("""\
    Value Filldown ArrayName (\S+)
    Value StorageGroupName (.+)
    Value ServerName (.+)
    
    Start
      ^\S+NavisecCli.exe\s+-np\s+arrayname -> ArrayNameLine

    ArrayNameLine
      ^\s*Array Name:\s+${ArrayName} -> PortStart 

    PortStart
      ^\S+NavisecCli.exe\s+-np\s+port\s+-messner\s+-list\s+-all -> PortLine

    PortLine
      ^\s*Information about each HBA:\s* -> HBALine
      ^\s*(\*+) -> Start
      
    HBALine
      ^Server Name:\s+${ServerName} 
      ^\s*Information about each port of this HBA:\s* -> HBAPortLine
      
    HBAPortLine
      ^\s*StorageGroup Name:\s+${StorageGroupName} -> Record PortLine
""")


def process(workbook: Any, content: str) -> Any:
    """Process Storage-Groups worksheet

    :param workbook:
    :param content:
    :return:
    """

    worksheet_name = 'Storage-Groups'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = list(concat([
        get_parser_header(PORT_TMPL),
        get_parser_header(STORAGEGROUP_TMPL)[3:],
    ]))
    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    cmd_storagegroup_out = run_parser_over(content, STORAGEGROUP_TMPL)
    cmd_port_out = run_parser_over(content, PORT_TMPL)

    common_columns = (0, 1)
    server_names_grouped = compose(
        valmap(
            compose(list, set, map(last))),
        groupby(
            itemgetter(*common_columns))
    )(cmd_port_out)

    cmd_port_relevant = map(
        juxt(
            compose(first, first),
            compose(second, first),
            second)
    )(server_names_grouped.items())

    common_columns_getter = itemgetter(*common_columns)
    cmd_merged_out = join(
        common_columns_getter, cmd_port_relevant,
        common_columns_getter, cmd_storagegroup_out)

    cmd_merged_out = sorted(cmd_merged_out)

    rows = list(map(
        compose(
            list,
            concat,
            juxt(
                first,
                compose(
                    drop(3),
                    second)))
    )(cmd_merged_out))

    portcmd = {(array, grp) for array, grp, *other in rows}
    strgp = {(array, grp) for array, grp, *other in cmd_storagegroup_out}
    no_server_groups = strgp - portcmd

    storage_list = list(filter(
        lambda storage_gr: any(
            fnmatch(str((storage_gr[0], storage_gr[1])), str(ctrlServer))
            for ctrlServer in no_server_groups),
        cmd_storagegroup_out))

    storage_list = check_empty_arrays(
        list(unique(storage_list + rows, key=itemgetter(0, 1))))

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, storage_list), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
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
        'StorageGroupsTable',
        'Storage-Groups',
        final_col,
        final_row)

    return [[lun_map[0], lun_map[1], lun_map[4]] for lun_map in storage_list]
