"""Isilon utilities"""
import json
import re
from collections import defaultdict
from contextlib import suppress
from fnmatch import fnmatch
from typing import Any, Generator, Iterable

from cytoolz import first

from ntsparser.parsing import decode_bytes, raw_gz_content
from ntsparser.utils import column_sum, ordered_jsons, flatten_dict, percentile


def collected_data(content: dict, key: str, value: str) -> Any:
    """Returns target collected_data if we are at the desired key value

    :param content:
    :param key:
    :param value:
    :return:
    """
    with suppress(TypeError):
        if fnmatch(content[key], value):
            if isinstance(content['target'], list):
                return [data['collected_data'] for data in content['target']
                        if isinstance(data['collected_data'], str)]
            return content['target']['collected_data']
    return None


def collected_alias_data(content: dict, key: str, value: str) -> Any:
    """Returns target collected_data if we are at the desired key/@alias value

    :param content:
    :param key:
    :param value:
    :return:
    """
    with suppress(TypeError):
        if fnmatch(content[key]['@alias'], value):
            return content['target']['collected_data']
    return None


def quotas_json(content: list) -> Generator:
    """Returns the needed values from the Quotas jsons

    :param content:
    :return:
    """
    for row in content:
        yield ([
            row['path'], row['type'], row['appliesto'],
            str(row['thresholds']['hard']), str(row['thresholds']['soft']),
            str(row['thresholds']['advisory']), str(row['usage']['physical']),
            str(row['usage']['logical']), str(row['usage']['inodes'])
        ])


def squash(row: list, squash_start: int, squash_end: int) -> list:
    """Squashes csv rows resulted from bad formatting

    :param row:
    :param squash_start:
    :param squash_end:
    :return:
    """
    return \
        row[:squash_start] \
        + [''.join(row[squash_start:squash_end])] \
        + row[squash_end:]


def get_multiprotocol(entry: list) -> list:
    """Returns with Multiprotocol type the entry is both NFS and SMB

    :param entry:
    :return:
    """
    if len(entry) == 1:
        return entry[0]
    return entry[0][:-1] + ['Multiprotocol']


def convert_jsons(string_jsons: str) -> list:
    """Gets jsons from collected_data which are big strings

    :param string_jsons:
    :return:
    """
    if string_jsons.startswith('['):
        string_jsons = string_jsons[1:-1]
    converted = string_jsons.split('}, {')
    return \
        [converted[0] + '}'] \
        + ['{' + el + '}' for el in converted[1:-1]] \
        + ['{' + converted[-1]]


def process_nfs_zones(content: str, headers: list) -> Generator:
    """Processes jsons for the NFS Zones sheet

    :param content:
    :param headers:
    :return:
    """
    converted = convert_jsons(content)
    fixed = []  # type: list
    for row in converted:
        # json.loads() does not work on multiple lists of jsons
        # so this split is necessary
        if ']\n[' in row:
            to_fix = row.split(']\n[')
            fixed += [json.loads(fixed_json) for fixed_json in to_fix]
        else:
            fixed.append(json.loads(row))
    flattened = [flatten_dict(row) for row in fixed]
    return ordered_jsons(flattened, headers)


def process_smb_zones(content: str, headers: list) -> Generator:
    """Processes jsons for the SMB Zones sheet

    :param content:
    :param headers:
    :return:
    """
    rows = []  # type: list
    if ']\n[' in content:
        for row in content.split('\n'):
            rows += json.loads(row)
    else:
        rows = json.loads(content)
    flattened = [smb_flatten(row) for row in rows]
    return ordered_jsons(flattened, headers)


def process_drives(content: str, headers: list) -> tuple:
    """Processes json lines for the Drive List sheet

    :param content:
    :param headers:
    :return:
    """
    lines = content.split('\n')
    rows, bad_jsons = [], 0
    for line in lines:
        cluster, jsons = line.split()[0].strip(':'), ' '.join(line.split()[1:])
        try:
            for json_entry in json.loads(jsons):
                rows.append(
                    [cluster]
                    + first(ordered_jsons([flatten_dict(json_entry)], headers)))
        except json.JSONDecodeError:
            bad_jsons += 1
    return rows, bad_jsons


def smb_flatten(smb_dict: dict) -> dict:
    """Flattens the permissions list of dictionaries for SMB jsons

    :param smb_dict:
    :return:
    """
    permissions = defaultdict(list)  # type: dict
    for perm_entry in smb_dict['permissions']:
        for key, val in flatten_dict(perm_entry).items():
            permissions[key].append(val)
    smb_dict['permissions'] = permissions
    return flatten_dict(smb_dict)


def row_correction(row: list) -> list:
    """Corrects hw_status rows that were not/couldn't be split correctly

    :param row:
    :return:
    """
    if len(row) > 3:
        return row[:2] + [' '.join(row[2:])]
    if len(row) < 3:
        expanded = row[-1].split()
        return row[:-1] + [' '.join(expanded[:-1])] + [expanded[-1]]
    return row


def hw_split(data: str) -> list:
    """Splits data rows for hw_status sheet

    :param data:
    :return:
    """
    return [row_correction(re.split(r'(?::|=)', row))
            for row in data.split('\n')]


def perf_cluster_name(csv_file: str) -> str:
    """Gets cluster name for isilon performance files

    :param csv_file:
    :return:
    """
    return csv_file.split('\n')[0].split('/')[-1].split('_')[0]


def perf_dashboard(content: Any) -> list:
    """Performs computations for the Performance Dashboard sheet

    :param content:
    :return:
    """
    top_level_dirs, nr_dirs, nr_files, throughput_avg, ops_avg, \
        ops_percentile, latency_avg, total, under_128, under_128_percent, \
        over_128, over_128_percent = [0.0] * 12

    top_level_dirs = len(content[0])
    nr_dirs = column_sum(content[0], 2) / 1000000
    nr_files = column_sum(content[0], 3) / 1000000

    if content[1]:
        throughput_avg = column_sum(content[1], 2) \
                         / len(list(zip(*content[1]))[2]) / 1024 ** 2

    if content[2]:
        ops_avg = column_sum(content[2], 2) \
                  / len(list(zip(*content[2]))[2])

        ops_percentile = percentile(
            list(map(float, list(zip(*content[2]))[2])), 0.95)

    if content[3]:
        latency_avg = (column_sum(content[3], -3) + column_sum(content[3], -5))\
                      / len(content[3])

    if content[4]:
        total = column_sum(content[4], 2)

        under_128 = column_sum(content[4], 3) + column_sum(content[4], 4)
        over_128 = sum([
            column_sum(content[4], col) for col in range(5, len(content[4][0]))
        ])

        under_128_percent = under_128 / total * 100
        over_128_percent = over_128 / total * 100

    return [
        top_level_dirs, nr_dirs, nr_files, throughput_avg, ops_avg,
        ops_percentile, latency_avg, total, under_128, under_128_percent,
        over_128, over_128_percent
    ]


def isilon_raw_content(input_files: tuple) -> Iterable:
    """Gets files for Isilon parsing

    :param input_files:
    :return:
    """
    raw_content = list()  # type: list
    for input_file in input_files:
        if input_file.endswith('.xml'):
            with open(input_file, 'rb') as input_xml:
                raw_content.append(
                    (input_file, decode_bytes(input_xml.readlines())))
        else:
            raw_content += raw_gz_content((input_file, ))
    return raw_content
