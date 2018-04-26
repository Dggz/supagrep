"""Isilon command entry point"""
import os

from openpyxl import load_workbook

from supergrep.isilon.count_logical import process as count_logical
from supergrep.isilon.count_modified import process as count_modified
from supergrep.isilon.drive_list import process as drive_list
from supergrep.isilon.file_system_protocol import \
    process as file_system_protocol
from supergrep.isilon.hw_status import process as hw_status
from supergrep.isilon.latency import process as latency
from supergrep.isilon.license_summary import process as license_summary
from supergrep.isilon.nfs_exports_list import process as nfs_exports_list
from supergrep.isilon.nfs_exports_zone import process as nfs_exports_zone
from supergrep.isilon.ops import process as ops
from supergrep.isilon.performance_dashboard import \
    process as performance_dashboard
from supergrep.isilon.protocol_stats import process as protocol_stats
from supergrep.isilon.pstat import process as pstat
from supergrep.isilon.quotas import process as quotas
from supergrep.isilon.smb_shares_list import process as smb_shares_list
from supergrep.isilon.smb_shares_zone import process as smb_shares_zone
from supergrep.isilon.storage_inventory import process as storage_inventory
from supergrep.isilon.storage_pool_summary import \
    process as storage_pool_summary
from supergrep.isilon.sync_policies import process as sync_policies
from supergrep.isilon.throughput import process as throughput
from supergrep.isilon.top_level_directories import \
    process as top_level_directories
from supergrep.isilon.zone_list import process as zone_list
from supergrep.parsing import raw_tar_content
from supergrep.isilon.utils import isilon_raw_content
from supergrep.utils import (
    get_bundle_dir, get_separated_content, relevant_content_file_join)

__all__ = ('main',)


def main(input_files: str, output_file: str) -> None:
    """Isilon Entry point

    :param input_files:
    :param output_file:
    """
    template_path = os.path.join(
        get_bundle_dir(), r'resources\isilon-template.xlsx')
    workbook = load_workbook(template_path)

    raw_content_patterns = (
        '*.xml.gz',
        '*fsa_export_module_directories.csv',
        '*fsa_export_module_filecountlogical.csv',
        '*perf_export_ext_net_bytes_by_direction.csv',
        '*fsa_export_module_filecountmodified.csv',
        '*perf_export_module_proto_latency.csv',
        '*perf_export_module_proto_op_rate.csv',
        '*.xml'
    )

    raw_content = list(isilon_raw_content(tuple(input_files)))

    perf_raw_content = []  # type: list
    for input_file in raw_content:
        perf_raw_content += raw_tar_content(
            (input_file[0], ), raw_content_patterns[1:])

    isilon_content = get_separated_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[7]))

    top_level_content = relevant_content_file_join(
        perf_raw_content, (raw_content_patterns[1], ))

    count_logical_content = relevant_content_file_join(
        perf_raw_content, (raw_content_patterns[2], ))

    throughput_content = relevant_content_file_join(
        perf_raw_content, (raw_content_patterns[3], ))

    count_modified_content = relevant_content_file_join(
        perf_raw_content, (raw_content_patterns[4], ))

    latency_content = relevant_content_file_join(
        perf_raw_content, (raw_content_patterns[5], ))

    ops_content = relevant_content_file_join(
        perf_raw_content, (raw_content_patterns[6], ))

    storage_inventory(workbook, isilon_content)
    storage_pool_summary(workbook, isilon_content)
    license_summary(workbook, isilon_content)
    nfs_rows = nfs_exports_list(workbook, isilon_content)
    smb_rows = smb_shares_list(workbook, isilon_content)
    quotas(workbook, isilon_content)
    file_system_protocol(workbook, nfs_rows, smb_rows)
    nfs_exports_zone(workbook, isilon_content)
    smb_shares_zone(workbook, isilon_content)
    zone_list(workbook, isilon_content)
    protocol_stats(workbook, isilon_content)
    pstat(workbook, isilon_content)
    drive_list(workbook, isilon_content)
    hw_status(workbook, isilon_content)
    sync_policies(workbook, isilon_content)
    perf_data = top_level_directories(workbook, top_level_content), \
        throughput(workbook, throughput_content), \
        ops(workbook, ops_content), \
        latency(workbook, latency_content), \
        count_logical(workbook, count_logical_content)

    performance_dashboard(workbook, perf_data)
    count_modified(workbook, count_modified_content)

    workbook.save(output_file)
