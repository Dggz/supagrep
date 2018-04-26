"""Various utilities"""

import itertools
from typing import Any, Callable


def epartial(func: Callable, *args: Any, **keywords: Any) -> Callable:
    """Enhanced partial

    A modified version of partial, allowing to omit some positional arguments
    between the leftmost and rightmost ones. The only drawback is that you can't
    pass anymore Ellipsis as argument.

    :param func:
    :param args:
    :param keywords:
    :return:
    """
    # pylint: disable=missing-docstring, missing-return-doc
    def newfunc(*fargs, **fkeywords):                       # type: ignore
        newkeywords = keywords.copy()
        newkeywords.update(fkeywords)
        return func(
            *(newfunc.leftmost_args + fargs + newfunc.rightmost_args),
            **newkeywords)
    # pylint: enable=missing-docstring, missing-return-doc
    newfunc.func = func                                     # type: ignore
    args = iter(args)                                       # type: ignore
    newfunc.leftmost_args = tuple(itertools.takewhile(      # type: ignore
        lambda v: v != Ellipsis, args))
    newfunc.rightmost_args = tuple(args)                    # type: ignore
    newfunc.keywords = keywords                             # type: ignore
    return newfunc
