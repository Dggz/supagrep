"""File parsing utilities"""
import gzip
import tarfile
from contextlib import suppress
from io import StringIO, BytesIO
from itertools import chain
from typing import List, Generator, Union, IO, Iterable
from zipfile import ZipFile

import textfsm

from ntsparser.utils import pattern_filter, tar_pattern_filter


def run_parser_over(content: str, template: str) -> list:     # pylint: disable=redefined-builtin
    """Run textfsm template over content

    :param content:
    :param template:
    :return:
    """
    fsm = textfsm.TextFSM(StringIO(template))
    result = fsm.ParseText(content)
    return result


def get_parser_header(template: str) -> List[str]:
    """Return list of header columns for parser

    :param template:
    :return:
    """
    fsm = textfsm.TextFSM(StringIO(template))
    return fsm.header


def get_file_content(zip_file: ZipFile, file_name: str) -> str:
    """Loads the file content of a file in a zip file

    :param zip_file:
    :param file_name:
    :return:
    """
    with suppress(UnicodeDecodeError):
        return zip_file.open(file_name).read().decode('utf-8')


def load_zip(zip_file_name: Union[IO[bytes], str]) -> Generator:
    """Loads files inside a nested .zip

    :param zip_file_name:
    :return:
    """
    with ZipFile(zip_file_name, 'r') as zip_file:
        for nested_file in zip_file.filelist:  # type: ignore
            if nested_file.filename.endswith('.zip'):
                nested_zip_name = BytesIO(
                    zip_file.open(nested_file).read())
                for x in load_zip(nested_zip_name):
                    yield x
            else:
                yield (nested_file.filename,
                       get_file_content(zip_file, nested_file.filename))


def load_zips(input_files: tuple) -> Generator:
    """Loads the file content inside the input .zips

    :param input_files:
    :return:
    """
    for first_zip_filename in input_files:
        for x in load_zip(first_zip_filename):
            yield x


def load_raw_content(input_files: tuple, patterns: tuple) -> Iterable:
    """Loads the data from all the needed input files

    :param input_files:
    :param patterns:
    :return:
    """
    return pattern_filter(load_zips(input_files), patterns)


def decode_bytes(nested_file_bytes: list) -> str:
    """Decodes the bytes of a file in utf-8, returns the file content as string

    :param nested_file_bytes:
    :return:
    """
    converted_to_str = ''
    for line in nested_file_bytes:
        converted_to_str += line.decode('utf-8')
    return converted_to_str


def get_files_from_tar(input_tar: str, patterns: tuple) -> list:
    """Gets the needed files from a .tbz2 archive

    :param input_tar:
    :param patterns:
    :return:
    """
    content_list = []
    with tarfile.open(input_tar) as tar_file:
        for nested_file in tar_pattern_filter(tar_file, patterns):
            content_list.append(
                (nested_file.name,
                 decode_bytes(tar_file.extractfile(nested_file).readlines())))
    return content_list


def raw_tar_content(input_files: tuple, patterns: tuple) -> Iterable:
    """Opens .zip or .tbz2 files for 3par

    The .zip case uses the filesystem

    :param input_files:
    :param patterns:
    :return:
    """
    raw_content = chain()  # type: chain
    for input_file in input_files:
        if input_file.endswith('.zip'):
            with ZipFile(input_file, 'r') as zf:
                for nested_file in zf.filelist:  # type: ignore
                    raw_content = chain(
                        raw_content, get_files_from_tar(zf.extract(
                            nested_file), patterns))
        else:
            with suppress(tarfile.ReadError):
                raw_content = chain(
                    raw_content, get_files_from_tar(input_file, patterns))
    return raw_content


def get_files_from_gz(input_gz: str) -> tuple:
    """Unpacks input .gz file

    :param input_gz:
    :return:
    """
    with gzip.open(input_gz) as gzip_file:
        return gzip_file.filename, decode_bytes(gzip_file.readlines())


def raw_gz_content(input_files: tuple) -> Iterable:
    """Opens .zip or .gz files for Isilon

    The .zip case uses the filesystem

    :param input_files:
    :return:
    """
    raw_content = list()  # type: list
    for input_file in input_files:
        if input_file.endswith('.zip'):
            with ZipFile(input_file, 'r') as zf:
                for nested_file in zf.filelist:  # type: ignore
                    raw_content.append(get_files_from_gz(
                        zf.extract(nested_file)))
        else:
            raw_content.append(get_files_from_gz(input_file))
    return raw_content
