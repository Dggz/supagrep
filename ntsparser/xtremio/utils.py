"""XtremIO utilities"""
from copy import deepcopy


def compute_row(row: list) -> list:
    """Computes calculations for Performance Output row

    :param row:
    :return:
    """
    xls_row = deepcopy(row)
    xls_row[1:] = list(map(float, row[1:]))

    per_interval_read = str(xls_row[2] / (xls_row[1] + xls_row[2]) * 100)[:2] + r'%'
    per_interval_throughput = xls_row[1] + xls_row[2]
    per_interval_write_size = xls_row[1] * 1024 / xls_row[3]
    per_interval_read_size = xls_row[2] * 1024 / xls_row[4]
    per_interval_size = per_interval_write_size + per_interval_read_size

    return [
        *row, str(per_interval_read), str(per_interval_throughput),
        str(per_interval_write_size), str(per_interval_read_size),
        str(per_interval_size)
    ]


def store_summary(perf_storage: tuple, perf_row: list) -> tuple:
    """Stores data from Performance Output for Performance Summary calculations

    :param perf_storage:
    :param perf_row:
    :return:
    """
    for column_list, column_element in zip(
            perf_storage, perf_row[1:5] + perf_row[8:]):
        column_list.append(float(column_element.strip(r'%')))
    return perf_storage
