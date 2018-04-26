"""PStat sheet"""
import textwrap
from collections import namedtuple
from contextlib import suppress
from typing import Any

import xmltodict

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from ntsparser.isilon.utils import collected_alias_data
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output, search_tag_value

PSTAT_TMPL = textwrap.dedent("""\
    Value access (\S+)
    Value fsinfo (\S+)
    Value lookup (\S+)
    Value noop (\S+)
    Value read (\S+)
    Value readlink (\S+)
    Value rmdir (\S+)
    Value symlink (\S+)
    Value commit (\S+)
    Value getattr (\S+)
    Value mkdir (\S+)
    Value null (\S+)
    Value readdir (\S+)
    Value remove (\S+)
    Value setattr (\S+)
    Value write (\S+)
    Value create (\S+)
    Value link (\S+)
    Value mknod (\S+)
    Value pathconf (\S+)
    Value readdirplus (\S+)
    Value rename (\S+)
    Value statfs (\S+)
    Value Total (\S+)
    Value user (\S+)
    Value system (\S+)
    Value idle (\S+)
    Value In (.+)
    Value Out (.+)
    Value OneFSTotal (.+)
    Value InputMB (\S+)
    Value InputPkt (\S+)
    Value InputErrors (\S+)
    Value OutputMB (\S+)
    Value OutputPkt (\S+)
    Value OutputErrors (\S+)
    Value Disk (.+)
    Value DiskRead (.+)
    Value DiskWrite (.+)

    Start
      ^\s*_+NFS3 Operations Per Second_+ -> OpsLines
      
    OpsLines
      ^\s*access\s+${access}\s+commit\s+${commit}\s+create\s+${create}
      ^\s*fsinfo\s+${fsinfo}\s+getattr\s+${getattr}\s+link\s+${link}
      ^\s*lookup\s+${lookup}\s+mkdir\s+${mkdir}\s+mknod\s+${mknod}
      ^\s*noop\s+${noop}\s+null\s+${null}\s+pathconf\s+${pathconf}
      ^\s*read\s+${read}\s+readdir\s+${readdir}\s+readdirplus\s+${readdirplus}
      ^\s*readlink\s+${readlink}\s+remove\s+${remove}\s+rename\s+${rename}
      ^\s*rmdir\s+${rmdir}\s+setattr\s+${setattr}\s+statfs\s+${statfs}
      ^\s*symlink\s+${symlink}\s+write\s+${write}
      ^\s*Total\s+${Total} -> CPUOneFSLines
      
    CPUOneFSLines
      ^\s*user\s+${user}\s+In\s+${In}
      ^\s*system\s+${system}\s+Out\s+${Out}
      ^\s*idle\s+${idle}\s+Total\s+${OneFSTotal} -> NetworkDiskLines
      
    NetworkDiskLines
      ^\s*MB/s\s+${InputMB}\s+MB/s\s+${OutputMB}\s+Disk\s+${Disk}
      ^\s*Pkt/s\s+${InputPkt}\s+Pkt/s\s+${OutputPkt}\s+Read\s+${DiskRead}
      ^\s*Errors/s\s+${InputErrors}\s+Errors/s\s+${OutputErrors}\s+Write\s+${DiskWrite} -> Record Start
""")


def process(workbook: Any, contents: list) -> None:
    """Process PStat worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'PStat'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = ['Hostname'] + get_parser_header(PSTAT_TMPL)

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    rows = []  # type: list
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        stats = []  # type: list
        host = component_details['hostname']
        for entry in command_details:
            with suppress(TypeError):
                stats_content = collected_alias_data(
                    entry, 'cmd', 'isi statistics pstat')
                stats = run_parser_over(
                    stats_content, PSTAT_TMPL) if stats_content else stats
        rows += [[host] + row for row in stats]

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
        'PStatTable',
        'PStat',
        final_col,
        final_row)
