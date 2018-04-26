"""Storage Array Summary Sheet (XtremIO)"""
import textwrap
from collections import namedtuple, defaultdict
from operator import itemgetter
from typing import Any

from cytoolz.curried import concat, map, unique
from openpyxl.styles import Alignment

from ntsparser.formatting import (
    build_header, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output, multiple_join


SHOW_CLUSTERS_TMPL = textwrap.dedent("""\
    Value Required ClusterName (\S+)
    Value ConnState (\S+)
    Value NumOfVols (\d+)
    Value NumInternalVols (\d+)
    Value UDSSDSpace (\d+\.\d+T)
    Value LogicalSpaceInUse (\d+\.\d+T)
    Value UDSSDSpaceInUse (\d+\.\d+T)
    Value SizeAndCapacity (\S+)
    
    Start
      ^\s*Cluster-Name\s+Index State\s+Gates-Open\s+Conn-State\s+Num-of-Vols\s+Num-of-Internal-Volumes\s+Vol-Size\s+UD-SSD-Space\s+Logical-Space-In-Use\s+UD-SSD-Space-In-Use\s+Total-Writes\s+Total-Reads\s+Stop-Reason\s+Size-and-Capacity -> ClusterLine

    ClusterLine
      ^\s*${ClusterName}\s+(\d+)\s+(\S+)\s+(\S+)\s+${ConnState}\s+${NumOfVols}\s+${NumInternalVols}\s+(\S+)\s+${UDSSDSpace}\s+${LogicalSpaceInUse}\s+${UDSSDSpaceInUse}\s+(\S+)\s+(\S+)\s+(\S+)\s+${SizeAndCapacity}\s* -> Record
      ^\s*(\*+) -> Start
""")


SHOW_CLUSTERS_INFO_TMPL = textwrap.dedent("""\
    Value Required ClusterName (\S+)
    Value PSNT (\S+)

    Start
      ^\s*Cluster-Name\s+Index State\s+Conn-State Activation-Time\s+Start-Time\s+SW-Version\s+PSNT\s+Encryption-Mode\s+Encryption-Supported\s+Encryption-Mode-State\s+Compression-Mode\s+SSH-Firewall-Mode\s+OS-Upgrade-Ongoing\s+Cluster-Expansion-In-Progress\s+Upgrade-State -> ClusterInfoLine

    ClusterInfoLine
      ^\s*${ClusterName}\s+(\d+)\s+(\S+)\s+(\S+)\s+(\S+\s+\S+\s+\d+\s+\S+\s+\d+)\s+(\S+\s+\S+\s+\d+\s+\S+\s+\d+)\s+(\S+)\s+${PSNT}\s+(.+) -> Record
""")


SHOW_X_BRICKS_TMPL = textwrap.dedent("""\
    Value Required ClusterName (\S+)
    Value Required BrickName (\S+)

    Start
      ^\s*Brick-Name\s+Index\s+Cluster-Name\s+Index State -> BrickLine

    BrickLine
      ^\s*${BrickName}\s+(\d+)\s+${ClusterName} -> Record
""")


SHOW_STORAGE_INFO_TMPL = textwrap.dedent("""\
    Value Required ClusterName (\S+)
    Value StorageControllerName (\S+)
    Value SWVersion (\S+)
    Value HWModel (\S+)
    Value SerialNumber (\S+)

    Start
      ^\s*Storage-Controller-Name\s+Index\s+Mgr-Addr\s+Brick-Name\s+Index\s+Cluster-Name\s+Index\s+State\s+Conn-State\s+SW-Version\s+SW-Build\s+HW-Model\s+OS-Version\s+Serial-Number\s+Part-Number\s+Sym-Storage-Controller\s+SC-Start-Timestamp -> StorageLine

    StorageLine
      ^\s*${StorageControllerName}\s+(\d+)\s+(\S+)\s+(\S+)\s+(\d+)\s+${ClusterName}\s+(\d+)\s+(\S+)\s+(\S+)\s+${SWVersion}\s+(\d+)\s+${HWModel}\s+(\w+\s\w+\s\w+\s\S+)\s+${SerialNumber}\s+(.+) -> Record
""")


CLUSTERS_SAVINGS_TMPL = textwrap.dedent("""\
    Value Required ClusterName (\S+)
    Value DataReductionRatio (\S+)
    Value ThinProvisioningSavings (\S+)
    Value DedupRatio (\S+)
    Value CompressionFactor (\S+)

    Start
      ^\s*Cluster-Name\s+Index\s+Data-Reduction-Ratio\s+Thin-Provisioning-Savings\s+Dedup-Ratio\s+Compression-Factor  -> SavingsLine

    SavingsLine
      ^\s*${ClusterName}\s+(\S+)\s+${DataReductionRatio}\s+${ThinProvisioningSavings}\s+${DedupRatio}\s+${CompressionFactor} -> Record
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> list:
    """Process Storage-Array-Summary worksheet (XtremIO)

    :param workbook:
    :param content:
    :return:
    """
    worksheet = workbook.get_sheet_by_name('Storage-Array-Summary')

    headers = list(concat([
        get_parser_header(SHOW_CLUSTERS_TMPL),
        get_parser_header(SHOW_CLUSTERS_INFO_TMPL)[1:],
        get_parser_header(SHOW_X_BRICKS_TMPL)[1:],
        get_parser_header(SHOW_STORAGE_INFO_TMPL)[1:],
        get_parser_header(CLUSTERS_SAVINGS_TMPL)[1:]
    ]))
    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    show_clusters_out = run_parser_over(content, SHOW_CLUSTERS_TMPL)
    clusters_info_out = run_parser_over(content, SHOW_CLUSTERS_INFO_TMPL)
    show_bricks_out = run_parser_over(content, SHOW_X_BRICKS_TMPL)
    show_storage_out = run_parser_over(content, SHOW_STORAGE_INFO_TMPL)
    clusters_savings_out = run_parser_over(content, CLUSTERS_SAVINGS_TMPL)

    common_columns = (0,)
    # noinspection PyTypeChecker
    bricks = defaultdict(list)  # type: dict
    for cluster, brick in show_bricks_out:
        bricks[cluster].append(brick)

    show_bricks_out = [[key, bricks[key]] for key in bricks]

    storages = defaultdict(list)  # type: dict
    for cluster, *st_info in show_storage_out:
        storages[cluster].append(st_info)

    show_storage_out = []
    idx = 0
    for key in storages:
        show_storage_out.append([key])
        info = zip(*storages[key])
        for val in info:
            show_storage_out[idx].append(val)
        idx += 1

    rows = multiple_join(
        common_columns,
        [show_clusters_out,
         clusters_info_out,
         show_bricks_out,
         show_storage_out,
         clusters_savings_out])

    rows = unique(rows, key=itemgetter(0))
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            if isinstance(col_value, str):
                cell.value = str.strip(col_value)
                if chr(col_n) in ('E', 'F', 'G'):
                    cell.value += 'B'
            else:
                cell.alignment = Alignment(wrapText=True)
                cell.value = '\n'.join(col_value)
            set_cell_to_number(cell)
            style_value_cell(cell)
            final_col = col_n
        final_row = row_n

    clusters = [
        ("Cluster Names:", ", ".join([
            clusters[0] for clusters in show_clusters_out]), '', '')
    ]

    sheet_process_output(
        worksheet,
        'StorageArraySummaryTable',
        'Storage-Array-Summary',
        final_col,
        final_row)

    return clusters
