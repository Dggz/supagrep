"""SAN Hosts sheet"""
import textwrap
from collections import namedtuple
from typing import Any, Iterable

import xmltodict
from openpyxl.styles import Alignment

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell, column_format)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output, search_tag_value, \
    ordered_jsons, flatten_dict, multiple_join
from ntsparser.xiv.utils import luns_occurrences, expand_rows

SYSTEM_NAME_TMPL = textwrap.dedent("""\
    Value Filldown,Required SystemName (\w+)

    Start
      ^.*Starting command\s+/xiv/admin/xcli.py\s+--ignore-trace-dumper\s+-u\s+xiv_development\s+-p\s+cod_list\s+-f\s+all -> BackendLines

    BackendLines
      ^\s*system_name\s+ ${SystemName} -> Start
""")


def process(workbook: Any, contents: Iterable) -> None:
    """Process SAN hosts Status worksheet

    :param workbook:
    :param contents:
    """
    worksheet_name = 'SAN Hosts'
    worksheet = workbook.get_sheet_by_name(worksheet_name)
    headers = get_parser_header(SYSTEM_NAME_TMPL)

    headers += [
        'cluster_id', 'host_id', 'volume_id', 'map_id', 'creator', 'Hostname',
        'cluster', 'fc_ports', 'type', 'iscsi_chap_name', 'perf_class'
    ]

    RowTuple = namedtuple('RowTuple', headers)
    build_header(worksheet, headers)
    headers = [
        'cluster_id/@value', 'host_id/@value', 'volume_id/@value',
        'creator/@value', 'name/@value', 'cluster/@value',
        'fc_ports/@value', 'type/@value',
        'iscsi_chap_name/@value', 'perf_class/@value'
    ]

    hosts_rows, lun_rows = [], []  # type: list
    for sys_content, all_content, host_content in contents:
        system_name = run_parser_over(sys_content, SYSTEM_NAME_TMPL)[0]
        all_map_content = '\n'.join(all_content.split('\n')[1:])
        host_content = '\n'.join(host_content.split('\n')[1:])

        doc_map = xmltodict.parse(all_map_content)
        map_details = search_tag_value(doc_map, 'map')
        maps = luns_occurrences(map_details, headers)
        lun_rows = [system_name + row for row in maps]

        doc_host = xmltodict.parse(host_content)
        host_details = search_tag_value(doc_host, 'host')
        flat_data_host = [flatten_dict(data) for data in host_details]
        hosts = ordered_jsons(flat_data_host,
                              [headers[0], 'id/@value'] + headers[3:])

        hosts_rows += [system_name + host_row for host_row in hosts]

    no_cluster = filter(lambda x: x[1] == '-1', lun_rows)
    no_hosts = filter(lambda x: x[2] == '-1', lun_rows)
    clusters_hosts = filter(lambda x: x[2] != '-1' and x[1] != '-1', lun_rows)

    common_columns = (0, 2)
    rows_cluster = multiple_join(
        common_columns, [no_cluster, hosts_rows])

    common_columns = (0, 1)
    row_hosts = multiple_join(
        common_columns, [no_hosts, hosts_rows])

    common_columns = (0, 1, 2)
    row_all = multiple_join(
        common_columns, [clusters_hosts, hosts_rows])

    sub_rows = list(rows_cluster) + list(row_hosts) + list(row_all)
    rows = expand_rows(sub_rows, 3)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(column_format(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            if chr(col_n) != 'D':
                set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'SANHostsTable',
        'SANHosts',
        final_col,
        final_row)
