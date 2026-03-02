"""
scripts/scatterplot.py — Scatter plot script for xlwings Lite.

Sheet setup:
  B2 : X data range address (e.g. "A2:A50")
  B3 : Y data range address (e.g. "B2:B50")
  B4 : Chart title
  B5 : X-axis label
  B6 : Y-axis label

Output:
  A PNG scatter plot image inserted into the sheet.
"""
import io

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import xlwings as xw

from utils.excel_helpers import get_range_as_list


@xw.script
def run_scatterplot(book: xw.Book) -> None:
    """Create a scatter plot from two data ranges and insert it into the sheet.

    Reads X and Y range addresses from cells B2 and B3 on the active sheet.
    Optional chart title and axis labels are read from B4, B5, B6.
    The resulting PNG image is inserted into the sheet.

    Args:
        book: The active xlwings Book object injected by xlwings Lite.
    """
    sheet = book.sheets.active

    x_address = sheet["B2"].value
    y_address = sheet["B3"].value
    chart_title = sheet["B4"].value or "Scatter Plot"
    x_label = sheet["B5"].value or "X"
    y_label = sheet["B6"].value or "Y"

    if not x_address or not y_address:
        sheet["D2"].value = "Error: Enter X range in B2 and Y range in B3."
        return

    x_data = get_range_as_list(sheet, x_address)
    y_data = get_range_as_list(sheet, y_address)

    if len(x_data) == 0 or len(y_data) == 0:
        sheet["D2"].value = "Error: No numeric data found in the specified ranges."
        return

    if len(x_data) != len(y_data):
        sheet["D2"].value = "Error: X and Y ranges must have the same number of values."
        return

    # Create the scatter plot
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.scatter(x_data, y_data, alpha=0.7, edgecolors="steelblue", facecolors="lightblue")
    ax.set_title(chart_title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    ax.grid(True, linestyle="--", alpha=0.5)
    fig.tight_layout()

    # Save plot to bytes buffer and insert into sheet
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=100)
    buf.seek(0)
    plt.close(fig)

    sheet.pictures.add(buf, name="ScatterPlot", update=True, left=sheet["D2"].left, top=sheet["D2"].top)
