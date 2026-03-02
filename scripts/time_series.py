"""
scripts/time_series.py — Time series analysis script for xlwings Lite.

Sheet setup:
  B2 : Time series data range address (e.g. "A2:A100")
  B3 : Rolling window size (integer, default 3)

Output written starting at D2:
  D column : Original series values
  E column : Rolling mean
  F column : Rolling standard deviation
  G column : Linear trend line values
  A chart showing all four series is also inserted into the sheet.
"""
import io

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import xlwings as xw

from utils.excel_helpers import get_range_as_list


@xw.script
def run_time_series(book: xw.Book) -> None:
    """Compute rolling statistics and a linear trend for a time series.

    Reads the time series data range from cell B2 and the rolling window
    size from cell B3 (default 3) on the active sheet.  Computes rolling
    mean, rolling standard deviation, and a linear trend via
    ``numpy.polyfit``.  Writes the results to columns D–G and inserts a
    chart into the sheet.

    Args:
        book: The active xlwings Book object injected by xlwings Lite.
    """
    sheet = book.sheets.active

    range_address = sheet["B2"].value
    window_raw = sheet["B3"].value
    window = int(window_raw) if window_raw is not None else 3

    if not range_address:
        sheet["D2"].value = "Error: Enter time series range address in B2."
        return

    series = get_range_as_list(sheet, range_address)
    if len(series) == 0:
        sheet["D2"].value = "Error: No numeric data found in the specified range."
        return

    n = len(series)

    if window < 2:
        sheet["D2"].value = "Error: Rolling window (B3) must be at least 2."
        return
    if window > n:
        sheet["D2"].value = f"Error: Rolling window ({window}) exceeds series length ({n})."
        return

    arr = np.array(series, dtype=float)
    x = np.arange(n)

    # Rolling mean and std (manual computation for Pyodide compatibility)
    rolling_mean = [None] * n
    rolling_std = [None] * n
    for i in range(window - 1, n):
        window_data = arr[i - window + 1 : i + 1]
        rolling_mean[i] = float(np.mean(window_data))
        rolling_std[i] = float(np.std(window_data, ddof=1))

    # Linear trend
    coeffs = np.polyfit(x, arr, 1)
    trend = np.polyval(coeffs, x).tolist()

    # Write headers
    sheet["D2"].value = "Original"
    sheet["E2"].value = f"Rolling Mean ({window})"
    sheet["F2"].value = f"Rolling Std ({window})"
    sheet["G2"].value = "Trend"

    for i in range(n):
        row = 3 + i
        sheet.cells(row, 4).value = arr[i]           # D
        sheet.cells(row, 5).value = rolling_mean[i]  # E
        sheet.cells(row, 6).value = rolling_std[i]   # F
        sheet.cells(row, 7).value = trend[i]          # G

    # Chart
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.plot(x, arr, label="Original", alpha=0.7)
    rm_clean = [v if v is not None else float("nan") for v in rolling_mean]
    rs_clean = [v if v is not None else float("nan") for v in rolling_std]
    ax.plot(x, rm_clean, label=f"Rolling Mean ({window})", linewidth=2)
    ax.plot(x, rs_clean, label=f"Rolling Std ({window})", linewidth=1, linestyle="--")
    ax.plot(x, trend, label="Trend", linewidth=2, linestyle=":")
    ax.set_title("Time Series Analysis")
    ax.set_xlabel("Period")
    ax.set_ylabel("Value")
    ax.legend()
    ax.grid(True, linestyle="--", alpha=0.5)
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=100)
    buf.seek(0)
    plt.close(fig)

    sheet.pictures.add(buf, name="TimeSeriesPlot", update=True, left=sheet["I2"].left, top=sheet["I2"].top)
