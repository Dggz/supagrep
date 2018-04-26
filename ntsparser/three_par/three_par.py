"""3Par command entry point"""
import os

from openpyxl import load_workbook

from ntsparser.parsing import raw_tar_content
from ntsparser.three_par.cage import process as cage
from ntsparser.three_par.cpg import process as cpg
from ntsparser.three_par.disks import process as disks
from ntsparser.three_par.hosts import process as hosts
from ntsparser.three_par.license_sheet import process as license_sheet
from ntsparser.three_par.nodes import process as nodes
from ntsparser.three_par.ports import process as ports
from ntsparser.three_par.storage_array_summary import process as storage_array_summary
from ntsparser.three_par.volumes import process as volumes
from ntsparser.utils import get_bundle_dir, get_relevant_content

__all__ = ('main',)


def main(input_files: str, output_file: str) -> None:
    """3Par Entry point

    :param input_files:
    :param output_file:
    """
    template_path = os.path.join(
        get_bundle_dir(), r'resources\3par-template.xlsx')
    workbook = load_workbook(template_path)

    raw_content_patterns = (
        '*showsys.out',
        '*showsys_-d.out',
        '*showversion.out',
        '*showpd.out',
        '*showcage.out',
        '*showport.out',
        '*showcpg.out',
        '*shownode_-d.out',
        '*showhost_-verbose.out',
        '*showvv.out',
        '*showvv_-cpgalloc.out',
        '*showvlun.out',
        '*controlencryption_status.out',
        '*showlicense.out'
    )

    raw_content = list(raw_tar_content(tuple(input_files), raw_content_patterns))

    storage_array_summary_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[1],
                      raw_content_patterns[2]), '*' * 20 + '\n')

    disks_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[3]), '*' * 20 + '\n')

    cage_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[4]), '*' * 20 + '\n')

    ports_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[5]), '*' * 20 + '\n')

    cpg_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[6]), '*' * 20 + '\n')

    nodes_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[7]), '*' * 20 + '\n')

    hosts_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[8]), '*' * 20 + '\n')

    volumes_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[9],
                      raw_content_patterns[10],
                      raw_content_patterns[11]), '*' * 20 + '\n')

    license_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[13],
                      raw_content_patterns[12]), '*' * 20 + '\n')

    storage_array_summary(workbook, storage_array_summary_content)
    disks(workbook, disks_content)
    cage(workbook, cage_content)
    ports(workbook, ports_content)
    cpg(workbook, cpg_content)
    nodes(workbook, nodes_content)
    hosts(workbook, hosts_content)
    volumes(workbook, volumes_content)
    license_sheet(workbook, license_content)

    workbook.save(output_file)
