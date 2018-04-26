"""Isilon utilities"""


def merge_dicts(dicts: list) -> dict:
    """Merges dictionaries that have the same keys

    :param dicts:
    :return:
    """
    new_dict = {}
    for k in dicts[0]:
        new_dict[k] = tuple(dic[k] for dic in dicts)
    return new_dict
