"""EVA command entry point"""
import os

from openpyxl import load_workbook

from supergrep.eva.controller import process as controller
from supergrep.eva.disk_enclosure import process as disks_enclosure
from supergrep.eva.disk_group import process as disk_group
from supergrep.eva.disks import process as disks
from supergrep.eva.hosts import process as hosts
from supergrep.eva.storage_inventory import process as storage_inventory
from supergrep.eva.virtual_disks import process as virtual_disks
from supergrep.parsing import load_raw_content
from supergrep.utils import get_bundle_dir, get_separated_content

__all__ = ('main',)


def main(input_files: str, output_file: str) -> None:
    """EVA Entry point

    :param input_files:
    :param output_file:
    """
    template_path = os.path.join(
        get_bundle_dir(), r'resources\eva-template.xlsx')
    workbook = load_workbook(template_path)

    raw_content_patterns = (
        '*EVA_config.xml',
    )
    raw_content = load_raw_content(tuple(input_files), raw_content_patterns)

    eva_content = get_separated_content(
        raw_content, (raw_content_patterns[0], ))

    storage_inventory(workbook, eva_content)
    controller(workbook, eva_content)
    disk_group(workbook, eva_content)
    virtual_disks(workbook, eva_content)
    hosts(workbook, eva_content)
    disks(workbook, eva_content)
    disks_enclosure(workbook, eva_content)

    workbook.save(output_file)
