"""XIV command entry point"""
import os

from openpyxl import load_workbook

from ntsparser.parsing import raw_tar_content
from ntsparser.utils import get_bundle_dir, get_relevant_content, \
    get_separated_content
from ntsparser.xiv.disk_drives import process as disk_drives
from ntsparser.xiv.pools import process as pools
from ntsparser.xiv.san_hosts import process as san_hosts
from ntsparser.xiv.storage_controllers import process as storage_controllers
from ntsparser.xiv.volumes import process as volumes

__all__ = ('main',)


def main(input_files: str, output_file: str) -> None:
    """XIV Entry point

    :param input_files:
    :param output_file:
    """

    template_path = os.path.join(
        get_bundle_dir(), r'resources\xiv-template.xlsx')
    workbook = load_workbook(template_path)

    raw_content_patterns = (
        '*cod_list_-f_all/cli.txt',
        '*version_get/cli.txt',
        '*conf_get_path=system/cli.txt',
        '*conf_get_path=misc.status/cli.txt',
        '*disk_list_-f_all/cli.txt',
        '*pool_list_-f_all/cli.txt',
        '*vol_list_-f_all_show_proxy=yes/cli.xml',
        '*all_mappings_list/cli.xml',
        '*host_list_-f_all/cli.xml'
    )

    raw_content = list(raw_tar_content(tuple(input_files), raw_content_patterns))
    storage_controllers_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[1],
                      raw_content_patterns[2]), '*' * 20 + '\n')

    disk_drivers_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[3],
                      raw_content_patterns[4]), '*' * 20 + '\n')

    pools_content = get_relevant_content(
        raw_content, (raw_content_patterns[0],
                      raw_content_patterns[5]), '*' * 20 + '\n')

    volumes_content_txt = get_separated_content(
        raw_content, (raw_content_patterns[0], ))

    volumes_content_xml = get_separated_content(
        raw_content, (raw_content_patterns[6], ))

    all_content_xml = get_separated_content(
        raw_content, (raw_content_patterns[7], ))

    hosts_mappings_content_xml = get_separated_content(
        raw_content, (raw_content_patterns[8], ))

    volumes_content = zip(volumes_content_txt, volumes_content_xml)
    hosts_content = zip(volumes_content_txt, all_content_xml,
                        hosts_mappings_content_xml)
    storage_controllers(workbook, storage_controllers_content)
    disk_drives(workbook, disk_drivers_content)
    pools(workbook, pools_content)
    volumes(workbook, volumes_content)
    san_hosts(workbook, hosts_content)

    workbook.save(output_file)
