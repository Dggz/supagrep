"""XIV utilities"""

from supergrep.utils import search_tag_value, \
    ordered_jsons, flatten_dict


def luns_occurrences(cmd_out: list, headers: list) -> list:
    """Creates a list for map_details

    :param headers:
    :param cmd_out:
    :return:
    """
    lun_rows = []  # type: list
    for entry in cmd_out:
        flat_data_map = [flatten_dict(entry)]
        maps = ordered_jsons(flat_data_map, headers[:2])
        lun_details = search_tag_value(entry, 'lun')
        if isinstance(lun_details, list):
            flat_luns = [flatten_dict(data) for data in lun_details]
        elif isinstance(lun_details, dict):
            flat_luns = [flatten_dict(lun_details)]
        luns = [flat_lun[headers[2]] for flat_lun in flat_luns]
        lun_rows += [row + [luns] for row in maps]

    return lun_rows


def expand_rows(rows: list, position: int) -> list:
    """Expands a row by a sublist at position

    Used to avoid multiple values in one cell in excel sheet
    :param rows:
    :param position:
    :return:
    """
    return [
        dup_row[:position] + [req_row] + dup_row[position+1:]
        for dup_row in rows for req_row in dup_row[position]
    ]
