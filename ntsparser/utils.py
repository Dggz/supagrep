"""Various utilities"""
import math
import os
import sys

from fnmatch import fnmatch
from logging import getLogger
from operator import itemgetter
from typing import Any, Callable, Iterable, List, Generator

from cytoolz.curried import (
    compose, concat, drop, first, join, juxt, map, second, groupby)
from openpyxl.styles import Alignment

from ntsparser.formatting import (
    add_worksheet_table, compute_column_dimensions, style_value_cell,
    set_cell_to_number)

logger = getLogger(__name__)


def get_bundle_dir() -> str:
    """Return path of project or bundle

    :return:
    """
    if getattr(sys, 'frozen', False):
        # we are running in a bundle
        bundle_dir = sys._MEIPASS
    else:
        # we are running in a normal Python environment
        dirname = os.path.dirname
        bundle_dir = dirname(dirname(os.path.abspath(__file__)))
    return bundle_dir


def pattern_filter(raw_content: Iterable, patterns: tuple) -> list:
    """Filters the content by the file name if it matches any of the patterns

    :param raw_content:
    :param patterns:
    :return:
    """
    content_dict = groupby(lambda x: x[0].split('/')[0], raw_content)
    contents = []
    for key in content_dict:
        for pattern in patterns:
            for file_content in content_dict[key]:
                if fnmatch(file_content[0], pattern):
                    contents.append(file_content)
    return contents


def tar_pattern_filter(tar_file: Iterable, patterns: tuple) -> list:
    """Filters the content by the file name if it matches any of the patterns

    :param tar_file:
    :param patterns:
    :return:
    """
    content_dict = groupby(lambda x: x.name.split('/')[0], tar_file)
    contents = []
    for key in content_dict:
        for pattern in patterns:
            for file_content in content_dict[key]:
                if fnmatch(file_content.name, pattern):
                    contents.append(file_content)
    return contents


def get_relevant_content(
        raw_content: Iterable,
        patterns: tuple,
        separator: str = '\n') -> str:
    """Gets the content of the files that match a pattern and concatenates it

    :param raw_content:
    :param patterns:
    :param separator:
    :return:
    """
    ct = pattern_filter(raw_content, patterns)
    return separator.join([content[1] for content in ct])


def get_separated_content(
        raw_content: Iterable,
        patterns: tuple) -> list:
    """Gets the content of the files that match a pattern

    :param raw_content:
    :param patterns:
    :return:
    """
    ct = pattern_filter(raw_content, patterns)
    return [content[1] for content in ct]


def relevant_content_file_join(
        raw_content: Iterable,
        patterns: tuple) -> List:
    """Joins the name of the file with the content

    If the the filename matches a pattern

    :param raw_content:
    :param patterns:
    :return:
    """
    ct = pattern_filter(raw_content, patterns)
    return ['\n'.join(content) for content in ct]


def sheet_process_output(
        worksheet: Any,
        table_name: str,
        sheet_name: str,
        final_col: int,
        final_row: int,
        start_col: int = ord('A'),
        start_row: int = 1) -> Any:
    """Verifies the result of each sheet's process function

    Saves the sheet if it has enough data
    Shows a message otherwise

    :param worksheet:
    :param table_name:
    :param sheet_name:
    :param final_col:
    :param final_row:
    :param start_col:
    :param start_row:
    """
    if final_col == 0 or final_row == 0:
        logger.error(
            '\nThere are no rows for the {} Table, {} sheet '
            '(source file or data not found).\n'.format(table_name, sheet_name))
        return

    add_worksheet_table(
        worksheet, table_name, final_col, final_row, start_col, start_row)
    compute_column_dimensions(worksheet)


def join_flatten(nr_common_cols: int, unflattened_list: list) -> Any:
    """Removes the common columns of the second list after a join operation on two lists

    Concatenates the two lists after removal

    :param nr_common_cols:
    :param unflattened_list:
    :return:
    """
    return map(
        compose(
            list,
            concat,
            juxt(
                first,
                compose(
                    drop(nr_common_cols),
                    second)))
    )(unflattened_list)


def multiple_join(common_columns: tuple, tables: list) -> Any:
    """Applies join on multiple lists

    Removes common columns by calling join_flatten

    :param common_columns:
    :param tables:
    :return:
    :raises ValueError:
    """
    if len(tables) < 2:
        raise ValueError
    rows = tables[0]
    for i in range(1, len(tables)):
        cmd_merged_out = join(
            itemgetter(*common_columns), rows,
            itemgetter(*common_columns), tables[i])
        rows = join_flatten(len(common_columns), cmd_merged_out)
    return rows


# noinspection TaskProblemsInspection
def multi(dispatch_fn: Callable) -> Callable:
    """Decorator that determines which version of a method should be called

    :param dispatch_fn:
    :return:
    """
    def _inner(*args, **kwargs):
        return _inner.__multi__.get(
            dispatch_fn(*args, **kwargs),
            _inner.__multi_default__
        )(*args, **kwargs)

    _inner.__multi__ = {}
    _inner.__multi_default__ = lambda *args, **kwargs: None  # Default default
    return _inner


# noinspection TaskProblemsInspection
def method(dispatch_fn: Callable, dispatch_key: Any = None) -> Callable:
    """Takes the output of multi decorator and calls the needed method version

    :param dispatch_fn:
    :param dispatch_key:
    :return:
    """
    def apply_decorator(fn: Callable) -> Callable:
        if dispatch_key is None:
            # Default case
            dispatch_fn.__multi_default__ = fn
        else:
            dispatch_fn.__multi__[dispatch_key] = fn
        return dispatch_fn

    return apply_decorator


def search_tag_value(raw_dict: dict, key_name: str) -> Any:
    """Recursively finds a key in a nested dictionary, returns its value

    :param raw_dict:
    :param key_name:
    :return:
    """
    if key_name in raw_dict.keys():
        return raw_dict[key_name]
    for key in raw_dict.keys():
        if isinstance(raw_dict[key], dict):
            found_tag = search_tag_value(raw_dict[key], key_name)
            if found_tag:
                return found_tag
    return None


def ordered_jsons(json_list: list, required_order: list) -> Generator:
    """For a list of jsons, takes the needed values in the specified order

    Sometimes there are unexpected dictionaries in the list coming from the
    XML so we ignore those

    :param json_list:
    :param required_order:
    :return:
    """
    for row in json_list:
        yield [str(row[col]) if col in row.keys() else ''
               for col in required_order]


def flatten_dict(unflattened: dict) -> Any:
    """Flattens a dictionary

    :param unflattened:
    :return:
    """
    flattened = dict()  # type: dict

    def build_dict(
            input_dict: dict,
            built_dict: dict,
            root_key: str = None) -> dict:
        """Builds a flattened dictionary

        Nested keys are joined with the root key

        :param input_dict:
        :param built_dict:
        :param root_key:
        :return:
        """
        for key, value in input_dict.items():
            if isinstance(value, dict):
                built_dict = build_dict(
                    value, built_dict, key)
            else:
                comp_key = '/'.join([root_key, key]) if root_key else key
                built_dict[comp_key] = value
        return built_dict
    return build_dict(unflattened, flattened)


def write_excel(
        rows: Iterable,
        worksheet: Any,
        RowTuple: Any,
        start_col: str) -> tuple:
    """Writes rows in excel from a specified start column

    :param rows:
    :param worksheet:
    :param RowTuple:
    :param start_col:
    :return:
    """
    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, rows), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord(start_col)):
            cell = worksheet['{}{}'.format(chr(col_n), row_n)]
            col_value = str(col_value) \
                if not isinstance(col_value, str) else col_value
            if isinstance(col_value, str):
                cell.value = str.strip(col_value)
            else:
                cell.alignment = Alignment(wrapText=True)
                cell.value = '\n'.join(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n
    return final_col, final_row


def percentile(values: list, percent: float) -> Any:
    """Find the percentile of a list of values.

    :param values:
    :param percent: float between 0.0 and 1.0
    :return:
    """
    sorted_values = sorted(values)
    if not sorted_values:
        return None
    percentile_value = (len(sorted_values) - 1) * percent
    floor = math.floor(percentile_value)
    ceiling = math.ceil(percentile_value)
    if floor == ceiling:
        return sorted_values[int(percentile_value)]
    d0 = sorted_values[int(floor)] * (ceiling - percentile_value)
    d1 = sorted_values[int(ceiling)] * (percentile_value - floor)
    return d0 + d1


def column_sum(rows: Iterable, column: int) -> float:
    """Computes the sum on a given column in each list of a list of lists

    :param rows:
    :param column:
    :return:
    """
    if rows:
        return sum(map(float, list(zip(*rows))[column]))
    return 0
