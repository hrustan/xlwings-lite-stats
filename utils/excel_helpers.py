"""
utils/excel_helpers.py — Shared utilities for reading from and writing to Excel sheets.

All scripts should use these helpers rather than accessing sheets directly,
keeping sheet I/O logic in one place.
"""
from collections import OrderedDict
from typing import Union

import pandas as pd


def get_range_as_df(sheet, range_address: str) -> pd.DataFrame:
    """Read an Excel range into a pandas DataFrame.

    Args:
        sheet: An xlwings Sheet object.
        range_address: The range address string, e.g. ``"A1:C10"``.

    Returns:
        A pandas DataFrame whose columns match the columns in the range.
        The first row of the range is used as column headers.
    """
    data = sheet[range_address].options(pd.DataFrame, header=True).value
    return data


def get_range_as_list(sheet, range_address: str) -> list:
    """Read an Excel range into a flat Python list of floats.

    Args:
        sheet: An xlwings Sheet object.
        range_address: The range address string, e.g. ``"A1:A20"``.

    Returns:
        A flat list of float values from the range, with ``None`` values dropped.
    """
    raw = sheet[range_address].value
    # Flatten nested lists (multi-row/column ranges) into a single list
    if isinstance(raw, list):
        flat = []
        for item in raw:
            if isinstance(item, list):
                flat.extend(item)
            else:
                flat.append(item)
        return [float(v) for v in flat if v is not None]
    if raw is None:
        return []
    return [float(raw)]


def write_df_to_sheet(sheet, df: pd.DataFrame, start_cell: str) -> None:
    """Write a pandas DataFrame back to a sheet starting at the given cell.

    Args:
        sheet: An xlwings Sheet object.
        df: The DataFrame to write.
        start_cell: The top-left cell address, e.g. ``"D2"``.
    """
    sheet[start_cell].options(pd.DataFrame, header=True, index=False).value = df


def write_results_block(
    sheet,
    start_cell: str,
    title: str,
    results_dict: Union[dict, OrderedDict],
) -> None:
    """Write a titled results block to the sheet.

    Writes the title in the first row, then key-value pairs (one per row)
    below it.

    Args:
        sheet: An xlwings Sheet object.
        start_cell: The top-left cell address, e.g. ``"D2"``.
        title: A string title for the results block.
        results_dict: An ordered or plain dict mapping label strings to values.
    """
    rng = sheet[start_cell]
    row = rng.row
    col = rng.column

    # Write title
    sheet.cells(row, col).value = title
    row += 1

    # Write key-value pairs
    for key, value in results_dict.items():
        sheet.cells(row, col).value = key
        sheet.cells(row, col + 1).value = value
        row += 1


def clear_range(sheet, range_address: str) -> None:
    """Clear (delete) the contents of a range on the sheet.

    Args:
        sheet: An xlwings Sheet object.
        range_address: The range address string, e.g. ``"D2:H50"``.
    """
    sheet[range_address].clear_contents()
