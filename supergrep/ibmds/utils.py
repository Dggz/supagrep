"""IBM DS utilities"""
from collections import Generator
from contextlib import suppress


def get_rows(rows: Generator, header: list) -> list:
    """Gets the rows from a multiple header csv

    :param rows:
    :param header:
    :return:
    """
    with suppress(StopIteration):
        row = next(rows)
        while row[:len(header)] != header:
            row = next(rows)

        content_rows = []
        row = next(rows)
        while row != [''] * len(row):
            content_rows.append(row[:len(header)])
            row = next(rows)
        return content_rows
