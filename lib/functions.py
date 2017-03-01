# coding: utf-8
import re


def xl_cell_to_row_col(cell_str):
    """
    Convert a cell reference in A1 notation to a zero indexed row and column.

    :param cell_str: A1 style string.
    :return: row, col: Zero indexed cell row and column indices.
    """
    if not cell_str:
        return 0, 0

    range_parts = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')
    match = range_parts.match(cell_str)
    col_str = match.group(2)
    row_str = match.group(4)

    # Convert base26 column string to number.
    expn = 0
    col = 0
    for char in reversed(col_str):
        col += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1

    # Convert 1-index to zero-index
    row = int(row_str) - 1
    col -= 1

    return row, col
