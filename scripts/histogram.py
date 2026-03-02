"""
scripts/histogram.py — Histogram analysis script for xlwings Lite.

Sheet setup:
  B2 : range address of the single-column data (e.g. "A2:A101")

Output written starting at D2:
  D2  : "Bin Edge"  header
  E2  : "Frequency" header
  D3+ : lower edges of each bin
  E3+ : count of values in each bin
"""
import numpy as np
import xlwings as xw

from utils.excel_helpers import get_range_as_list


@xw.script
def run_histogram(book: xw.Book) -> None:
    """Compute a histogram and write bin edges and frequencies to the sheet.

    Reads the data range address from cell B2 on the active sheet, computes
    histogram bins and frequencies using ``numpy.histogram``, then writes
    the results starting at cell D2.

    Args:
        book: The active xlwings Book object injected by xlwings Lite.
    """
    sheet = book.sheets.active

    # Read data range address from B2
    range_address = sheet["B2"].value
    if not range_address:
        sheet["D2"].value = "Error: Enter a data range address in B2."
        return

    data = get_range_as_list(sheet, range_address)
    if len(data) == 0:
        sheet["D2"].value = "Error: No numeric data found in the specified range."
        return

    # Compute histogram using numpy (auto bin count via Sturges' rule)
    counts, bin_edges = np.histogram(data)

    # Write headers
    sheet["D2"].value = "Bin Edge"
    sheet["E2"].value = "Frequency"

    # Write bin edges (left edge of each bin) and corresponding frequency
    for i, (edge, count) in enumerate(zip(bin_edges[:-1], counts)):
        sheet.cells(3 + i, 4).value = edge   # Column D
        sheet.cells(3 + i, 5).value = int(count)  # Column E

    # Write the final right edge with empty frequency cell for completeness
    sheet.cells(3 + len(counts), 4).value = bin_edges[-1]
