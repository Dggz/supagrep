"""Celerra utilities"""
import re
import shlex
from collections import defaultdict
from contextlib import suppress

from supergrep.utils import multi, method


def classify_rows(server_export_grouped: dict) -> tuple:
    """Determines if a row belongs to share, export and multiprotocol

    And stores them accordingly

    :param server_export_grouped:
    :return:
    """
    multiprotocol, share, export = [], [], []
    for export_row in server_export_grouped['export']:
        export_flag = False
        for share_row in server_export_grouped['share']:
            if (export_row[0], export_row[2].split('/')[-1]) \
                    == (share_row[0], share_row[2].split('/')[-1].strip('$')):
                multiprotocol.append(
                    process_properties([*share_row[:3], 'multiprotocol'])
                )
                export_flag = True
        if not export_flag:
            export.append(process_properties(export_row))

    # We need this second 'for' to be sure that we don't add any of the
    # 'share' rows that have already been added in multiprotocol
    for share_row in server_export_grouped['share']:
        share_flag = False
        for multi_row in multiprotocol:
            if (share_row[0], share_row[2].split('/')[-1].strip('$')) \
                    == (multi_row[0], multi_row[2].split('/')[-1].strip('$')):
                share_flag = True
        if not share_flag:
            share.append(process_properties(share_row))
    return share, export, multiprotocol


def share_split(filesys: str) -> list:
    """Splits '/' property into SharePath and RootPath for 'share'

    :param filesys:
    :return:
    """
    fs = re.split(r'(/)\s*', filesys)
    if len(fs) > 3:
        return [''.join(fs), ''.join(fs[2:-2])]
    else:
        return [''.join(fs), ''.join(fs[2:])]


def export_split(filesys: str) -> list:
    """Splits MountedPath column into MountedPath and Filesystem for 'export'

    :param filesys:
    :return:
    """
    fs = re.split(r'(/)\s*', filesys)
    return [''.join(fs), fs[-1]]


def get_properties(properties: list) -> dict:
    """Gets the properties for share or export rows

    :param properties:
    :return:
    """
    prop_dict = defaultdict(list)  # type: defaultdict
    for prop in properties:
        # Using suppress for property data which
        # can't be split by '='
        with suppress(ValueError):
            key, val = prop.split('=')
            prop_dict[key].append(val)
    return prop_dict


@multi
def process_properties(row: list) -> str:
    """Returns the row type such that the correct version of the method is called

    :param row:
    :return:
    """
    return row[3]


@method(process_properties)  # Multiprotocol
def _(row: list) -> list:
    """Strips unused info for Multiprotocol rows

    :param row:
    :return:
    """
    return [*row[:3], 'multiprotocol']


@method(process_properties, 'export')
def _(row: list) -> list:
    """Extracts rw, root and access properties for an export row

    :param row:
    :return:
    """
    properties = get_properties(shlex.split(row[4]))
    return [
        *row[:2], *export_split(row[2]), row[3], properties['rw'],
        properties['root'], properties['access']
    ]


@method(process_properties, 'share')
def _(row: list) -> list:
    """Extracts umask, maxusr, netbios and comment properties for a share row

    :param row:
    :return:
    """
    props = shlex.split(row[4])
    properties = get_properties(props)
    filesys = '' if '=' in props[0] else props[0]
    return [
        *row[:3], *share_split(filesys), row[3], properties['umask'],
        properties['maxusr'], properties['netbios'], properties['comment']
    ]
