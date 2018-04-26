"""Celerra command entry point"""
import os

from openpyxl import load_workbook

from supergrep.celerra.backend_disk_info import process as backend_disk_info
from supergrep.celerra.backend_storage import process as backend_storage
from supergrep.celerra.backend_storage_details import process as backend_storage_details
from supergrep.celerra.cifs_share import process as cifs_share
from supergrep.celerra.disk_groups import process as disk_groups
from supergrep.celerra.fs_dedupe import process as fs_dedupe
from supergrep.celerra.nas_fs_info_all import process as nas_fs_info_all
from supergrep.celerra.nas_license import process as nas_license
from supergrep.celerra.nas_pool_info_all import process as nas_pool_info_all
from supergrep.celerra.nas_replicate_info import process as nas_replicate_info
from supergrep.celerra.nas_summary import process as nas_summary
from supergrep.celerra.physical_dm import process as physical_dm
from supergrep.celerra.pool_configuration import process as pool_configuration
from supergrep.celerra.server_df import process as server_df
from supergrep.celerra.system_details import process as system_details
from supergrep.celerra.virtual_dm import process as virtual_dm
from supergrep.celerra.volume_size import process as volume_size
from supergrep.parsing import load_raw_content
from supergrep.utils import get_bundle_dir, get_relevant_content

__all__ = ('main',)


def main(input_files: str, output_file: str) -> None:
    """Celerra Entry point

    :param input_files:
    :param output_file:
    """
    template_path = os.path.join(
        get_bundle_dir(), r'resources\celerra-template.xlsx')
    workbook = load_workbook(template_path)

    raw_content_patterns = (
        '*cmd_outputs/hostname',
        '*cmd_outputs/server_version',
        '*cmd_outputs/nas_summary',
        '*cmd_outputs/nas_license',
        '*cmd_outputs/nas_storage',
        '*cmd_outputs/nas_pool-info',
        '*cmd_outputs/*/navicli-getcrus',
        '*cmd_outputs/*/navicli-getagent',
        '*cmd_outputs/nas_server-list',
        '*cmd_outputs/server_export',
        '*cmd_outputs/server_df',
        '*cmd_outputs/fs_dedupe-list',
        '*cmd_outputs/nas_fs-info-all-o-mpd.txt',
        '*cmd_outputs/nas_replicate-info-all',
        '*cmd_outputs/nas_volume-info-size',
    )

    raw_content = load_raw_content(tuple(input_files), raw_content_patterns)

    system_details_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[1],
                      raw_content_patterns[6],
                      raw_content_patterns[7]), '*' * 20 + '\n')

    nas_summary_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[2]))

    nas_license_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[3]), '*' * 20 + '\n')

    pool_configuration_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[4]), '*' * 20 + '\n')

    nas_pool_info_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[5]), '*' * 20 + '\n')

    disk_groups_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[4]), '*' * 20 + '\n')

    backend_storage_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[4]), '*' * 20 + '\n')

    backend_disk_info_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[4]), '*' * 20 + '\n')

    backend_details_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[4]), '*' * 20 + '\n')

    physical_dm_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[8]), '*' * 20 + '\n')

    virtual_dm_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[8]), '*' * 20 + '\n')

    cifs_share_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[9]), '*' * 20 + '\n')

    serverd_df_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[10]), '*' * 20 + '\n')

    fs_dedupe_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[11]), '*' * 20 + '\n')

    nas_fs_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[12]), '*' * 20 + '\n')

    nas_replicate_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[13]), '*' * 20 + '\n')

    volume_size_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[14]), '*' * 20 + '\n')

    system_details(workbook, system_details_content)
    nas_summary(workbook, nas_summary_content)
    nas_license(workbook, nas_license_content)
    pool_configuration(workbook, pool_configuration_content)
    nas_pool_info_all(workbook, nas_pool_info_content)
    disk_groups(workbook, disk_groups_content)
    backend_storage(workbook, backend_storage_content)
    backend_disk_info(workbook, backend_disk_info_content)
    backend_storage_details(workbook, backend_details_content)
    physical_dm(workbook, physical_dm_content)
    virtual_dm(workbook, virtual_dm_content)
    cifs_share(workbook, cifs_share_content)
    server_df(workbook, serverd_df_content)
    fs_dedupe(workbook, fs_dedupe_content)
    nas_fs_info_all(workbook, nas_fs_content)
    nas_replicate_info(workbook, nas_replicate_content)
    volume_size(workbook, volume_size_content)

    workbook.save(output_file)
