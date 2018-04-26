"""Storage Inventory sheet"""
import textwrap
from collections import namedtuple
from operator import itemgetter
from typing import Any

import xmltodict
from cytoolz.curried import concat
from cytoolz.functoolz import compose

from supergrep.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from supergrep.isilon.utils import collected_data
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output, search_tag_value

NODES_TMPL = textwrap.dedent("""\
    Value Nodes (\d+)

    Start
      ^\s*ID\s+Name\s+Nodes\s+Protection\s+Policy\s+Manual -> NodeLines

    NodeLines
      ^\s*(\S+)\s+(\S+)\s+${Nodes}\s+ -> Record
      ^\s*${Nodes}\s+$$ -> Record 
""")

DEDUPE_TMPL = textwrap.dedent("""\
    Value ClusterPhysicalSize (.+)
    Value ClusterUsedSize (.+)
    Value LogicalSizeDeduplicated (.+)
    Value LogicalSaving (.+)
    Value EstimatedSizeDeduplicated (.+)
    Value EstimatedPhysicalSaving (.+)

    Start
      ^\s*Cluster Physical Size:\s+${ClusterPhysicalSize}
      ^\s*Cluster Used Size:\s+${ClusterUsedSize}
      ^\s*Logical Size Deduplicated:\s+${LogicalSizeDeduplicated}
      ^\s*Logical Saving:\s+${LogicalSaving}
      ^\s*Estimated Size Deduplicated:\s+${EstimatedSizeDeduplicated}
      ^\s*Estimated Physical Saving:\s+${EstimatedPhysicalSaving} -> Record
""")


def process(workbook: Any, contents: list) -> None:
    """Process Storage Inventory worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'Storage Inventory'
    worksheet = workbook.get_sheet_by_name(worksheet_name)

    headers = list(concat([
        ['Hostname', 'Model', 'OS', 'Nodes'],
        get_parser_header(DEDUPE_TMPL)
    ]))
    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)

    rows = []
    for content in contents:
        doc = xmltodict.parse(content)
        component_details = search_tag_value(doc, 'component_details')
        command_details = search_tag_value(doc, 'command_details')

        dedupe, nodes = [], 0  # type: (list, int)
        for entry in command_details:
            nodes_content = collected_data(
                entry, 'cmd', 'isi storagepool nodepools list')
            nodes = max(map(compose(int, itemgetter(0)),
                            run_parser_over(
                                nodes_content,
                                NODES_TMPL))) if nodes_content else nodes

            dedupe_content = collected_data(entry, 'cmd', 'isi dedupe stats')
            dedupe = run_parser_over(
                dedupe_content, DEDUPE_TMPL) if dedupe_content else dedupe

        dedupe = dedupe if len(dedupe) > 1 else [['', '', '', '', '', '']]
        rows.append([
            component_details['hostname'],
            component_details['model'],
            component_details['os'], str(nodes), *dedupe[0]
        ])

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
        'StorageInventoryTable',
        'Storage Inventory',
        final_col,
        final_row)
