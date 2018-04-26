"""VNX command entry point"""
import os

from openpyxl import load_workbook

from ntsparser.parsing import load_raw_content
from ntsparser.utils import get_bundle_dir, get_relevant_content
from ntsparser.vnx.disks import process as disks
from ntsparser.vnx.disks_pivot import process as disks_pivot
from ntsparser.vnx.initiator_type_pivot import process as initiator_type_pivot
from ntsparser.vnx.lun_storage_pivot import process as lun_storage_pivot
from ntsparser.vnx.luns import process as luns
from ntsparser.vnx.luns_pivot import process as luns_pivot
from ntsparser.vnx.mirror_view_a import process as mirror_view_a
from ntsparser.vnx.mirror_view_s import process as mirror_view_s
from ntsparser.vnx.raid_groups import process as raid_groups
from ntsparser.vnx.snap_clones import process as snap_clones
from ntsparser.vnx.snap_view import process as snap_view
from ntsparser.vnx.software_packages import process as software_packages
from ntsparser.vnx.sp_frontend_ports import process as sp_frontend_ports
from ntsparser.vnx.storage_array_pivot import process as storage_array_pivot
from ntsparser.vnx.storage_array_summary import process as storage_array_summary
from ntsparser.vnx.storage_groups import process as storage_groups

__all__ = ('main',)


def main(input_files: str, output_file: str) -> None:
    """VNX Entry point

    :param input_files:
    :param output_file:
    """
    template_path = os.path.join(
        get_bundle_dir(), r'resources\emc-vnx-template.xlsx')
    workbook = load_workbook(template_path)

    raw_content_patterns = (
        'SPA_cfg_info.txt',
        'SPB_cfg_info.txt',
        '*SPA_mirror_info.txt',
        '*SPB_mirror_info.txt',
        '*SPA_snap_info.txt',
        '*SPB_snap_info.txt',
        '*SPA_clone_info.txt',
        '*SPB_clone_info.txt'
    )

    raw_content = load_raw_content(tuple(input_files), raw_content_patterns)
    spa_spb_content = get_relevant_content(
        raw_content, (raw_content_patterns[0], raw_content_patterns[1]))

    mirror_view_content = get_relevant_content(
        raw_content, (raw_content_patterns[2], raw_content_patterns[3]))

    snap_view_content = get_relevant_content(
        raw_content, (raw_content_patterns[4], raw_content_patterns[5]))

    snap_clones_content = get_relevant_content(
        raw_content, (raw_content_patterns[6], raw_content_patterns[7]))

    array_names, array_models, array_revisions = storage_array_summary(
        workbook, spa_spb_content)
    sp_frontend_data = sp_frontend_ports(workbook, spa_spb_content)
    storage_array_pivot(
        workbook, (sp_frontend_data, array_models, array_revisions))
    initiator_type_pivot(workbook, sp_frontend_data)
    storage_groups_data = storage_groups(workbook, spa_spb_content)
    luns_data = luns(workbook, spa_spb_content, storage_groups_data)
    luns_pivot(workbook, luns_data)
    lun_storage_pivot(workbook, luns_data)
    software_packages(workbook, spa_spb_content)
    disks_data = disks(workbook, spa_spb_content)
    disks_pivot(workbook, disks_data)
    raid_groups(workbook, spa_spb_content)
    mirror_view_s(workbook, mirror_view_content, array_names)
    mirror_view_a(workbook, mirror_view_content, array_names)
    snap_view(workbook, snap_view_content, array_names)
    snap_clones(workbook, snap_clones_content, array_names)

    workbook.save(output_file)
