# -*- coding: utf-8 -*-
import string


def column_letter(index):
    """In Sheets the columns are identified by letters, not integers.
    This function translates the column index into letter(s).

    Args:
        index (int): column index

    Returns:
        (str) the translated index as letter(s)
    """
    res = ""
    i = index
    while True:
        j = i % len(string.ascii_uppercase)
        res = string.ascii_uppercase[j] + res
        i = (i - j) // len(string.ascii_uppercase) - 1
        if i < 0:
            break
    return res


def column_index(column_str):
    """In Sheets the columns are identified by letters, not integers.
    This function translates the column string into an integer index.

    Args:
        index (str): column identifier

    Returns:
        (int) the translated index
    """
    res = string.ascii_uppercase.index(column_str[-1])
    mult = 1
    for l in column_str[-2::-1]:
        mult *= len(string.ascii_uppercase)
        res += (1 + string.ascii_uppercase.index(l)) * mult
    return res
