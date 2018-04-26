"""VNX utilities"""
from collections import defaultdict
from itertools import chain


def take_array_names(array_names: dict, cmd_out: list) -> list:
    """Takes the corresponding array name for a serial number

    If the serial number is found, the array name is set in cmd_out

    :param array_names:
    :param cmd_out:
    :return:
    """
    for entry in cmd_out:
        if entry[0] in array_names:
            entry[0] = array_names[entry[0]]
    return cmd_out


def array_occurrences(cmd_out: list) -> defaultdict:
    """Creates a dictionary that counts how many times each array appears

    :param cmd_out:
    :return:
    """
    array_frequency = defaultdict(int)  # type: defaultdict
    array_name = 0
    for entry in cmd_out:
        array_frequency[entry[array_name]] += 1
    return array_frequency


def check_empty_arrays(cmd_out: list) -> list:
    """Checks if the array did not have input, writes a comment in the last column

    Also removes extra blank rows which appear
    because of the Filldown attribute in the FSM

    :param cmd_out:
    :return:
    """
    array_counts = array_occurrences(cmd_out)
    new_cmd_out = []
    comments = -1
    for entry in cmd_out:
        flat_entry = chain(*entry)
        if ''.join(flat_entry) == entry[0]:
            if array_counts[entry[0]] == 1:
                entry[comments] = 'No data for this array'
                new_cmd_out.append(entry)
            else:
                array_counts[entry[0]] -= 1
        else:
            new_cmd_out.append(entry)
    return new_cmd_out


def capacity_conversion(
        capacity: str,
        conversion_factor: int = 1024 ** 2) -> str:
    """Converts capacity by a given factor, default is from MB to TB

    :param capacity:
    :param conversion_factor:
    :return:
    """
    if capacity.isdigit():
        return str(int(capacity) / conversion_factor)
    return str(0)


def get_luns(pairs: list) -> list:
    """Gets array luns and host luns from HLU/ALU pairs

    :param pairs:
    :return:
    """
    return [pair.split() for pair in pairs]
