"""
main_bundled.py — xlwings Lite add-in: all scripts and custom functions in one file.

HOW TO USE (for students / end users):
  1. Install the xlwings Lite add-in from Microsoft AppSource (free, one-time).
  2. Open Excel and open a workbook.
  3. Open the xlwings Lite task pane: Home → xlwings → Open Task Pane
  4. Click the Editor tab → Import → select THIS FILE (main_bundled.py).
  5. Click Restart in the task pane.
  6. Save the workbook. Done — all functions and scripts are now embedded.

Available task pane scripts (buttons):
  - Histogram        : reads range address from B2, writes bin edges & frequencies to D2+
  - Scatter Plot     : reads X from B2, Y from B3, title/labels from B4-B6; inserts chart
  - Regression       : reads Y from B2, X from B3, writes OLS results to D2+
  - Chi-Squared      : reads contingency table range from B2, writes results to D2+
  - Time Series      : reads range from B2, window from B3, writes to D2+ and inserts chart
  - Monte Carlo      : reads n_sims from B2, mean from B3, std from B4, periods from B5

Available Excel formulas (enter directly in a cell):
  =SLOPE_INTERCEPT(y_range, x_range)  → [[slope, intercept]]
  =R_SQUARED(y_range, x_range)        → R² scalar
  =DESCRIPTIVE_STATS(data_range)      → 9×2 stats table
  =Z_SCORE(data_range)                → same-length z-score array
"""

import io
from collections import OrderedDict
from typing import Union

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
import xlwings as xw
from scipy.stats import (
    chi2_contingency,
    kurtosis,
    linregress,
    skew,
    zscore as scipy_zscore,
)

# ── Shared helpers ─────────────────────────────────────────────────────────────


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

    Non-numeric and blank (``None``) values are silently skipped.

    Args:
        sheet: An xlwings Sheet object.
        range_address: The range address string, e.g. ``"A1:A20"``.

    Returns:
        A flat list of float values from the range, with non-numeric and
        ``None`` values dropped.
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
    elif raw is None:
        return []
    else:
        flat = [raw]

    result = []
    for v in flat:
        if v is None:
            continue
        try:
            result.append(float(v))
        except (TypeError, ValueError):
            continue
    return result


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


def _flatten(range_value) -> list:
    """Flatten an Excel range value (list-of-lists or flat list) to a list of floats."""
    if range_value is None:
        return []
    if not isinstance(range_value, list):
        return [float(range_value)]
    flat = []
    for item in range_value:
        if isinstance(item, list):
            flat.extend(item)
        else:
            flat.append(item)
    return [float(v) for v in flat if v is not None]


# ── Scripts (@xw.script → appear as task pane buttons) ────────────────────────


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

    # Compute histogram using numpy with default binning (10 bins)
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


@xw.script
def run_regression(book: xw.Book) -> None:
    """Fit an OLS regression model and write a formatted summary to the sheet.

    Reads Y and X range addresses from cells B2 and B3 on the active sheet.
    Supports both simple and multiple linear regression depending on the
    number of columns in the X range.  Results include coefficients,
    standard errors, t-statistics, p-values, R², adjusted R², and the
    F-statistic.

    Args:
        book: The active xlwings Book object injected by xlwings Lite.
    """
    sheet = book.sheets.active

    y_address = sheet["B2"].value
    x_address = sheet["B3"].value

    if not y_address or not x_address:
        sheet["D2"].value = "Error: Enter Y range in B2 and X range in B3."
        return

    # Read Y directly to preserve row alignment (do not use get_range_as_list
    # which silently drops blanks and could misalign rows relative to X).
    y_raw = sheet[y_address].value
    if y_raw is None:
        sheet["D2"].value = "Error: No data found in the Y range."
        return

    # Normalise Y to a 1-D list
    if not isinstance(y_raw, list):
        y_list = [y_raw]
    elif isinstance(y_raw[0], list):
        y_list = [row[0] for row in y_raw]
    else:
        y_list = y_raw

    # Read X — may be single or multi-column
    x_raw = sheet[x_address].value
    if x_raw is None:
        sheet["D2"].value = "Error: No data found in the X range."
        return

    # Normalise X to a 2-D list of rows
    if not isinstance(x_raw, list):
        x_rows = [[x_raw]]
    elif not isinstance(x_raw[0], list):
        x_rows = [[v] for v in x_raw]
    else:
        x_rows = x_raw

    if len(y_list) != len(x_rows):
        sheet["D2"].value = "Error: Y and X ranges must have the same number of rows."
        return

    # Convert to numeric arrays, validating no blanks or non-numeric cells
    if any(v is None for v in y_list):
        sheet["D2"].value = "Error: Y range must not contain blank cells."
        return
    try:
        y_arr = np.array([float(v) for v in y_list])
    except (TypeError, ValueError):
        sheet["D2"].value = "Error: Y range must contain only numeric values."
        return

    x_numeric_rows = []
    for row in x_rows:
        if any(v is None for v in row):
            sheet["D2"].value = "Error: X range must not contain blank cells."
            return
        try:
            x_numeric_rows.append([float(cell) for cell in row])
        except (TypeError, ValueError):
            sheet["D2"].value = "Error: X range must contain only numeric values."
            return

    x_data = np.array(x_numeric_rows)

    # Add constant for intercept
    x_with_const = sm.add_constant(x_data)

    model = sm.OLS(y_arr, x_with_const)
    results = model.fit()

    n_params = len(results.params)
    param_names = ["Intercept"] + [f"X{i}" for i in range(1, n_params)]

    # Build results block
    output = OrderedDict()
    output["R²"] = round(results.rsquared, 6)
    output["Adjusted R²"] = round(results.rsquared_adj, 6)
    output["F-statistic"] = round(results.fvalue, 6)
    output["F p-value"] = round(results.f_pvalue, 6)
    output["Observations"] = int(results.nobs)
    output[""] = ""  # spacer row

    for name, coef, se, tval, pval in zip(
        param_names,
        results.params,
        results.bse,
        results.tvalues,
        results.pvalues,
    ):
        output[f"{name} — Coef"] = round(coef, 6)
        output[f"{name} — Std Err"] = round(se, 6)
        output[f"{name} — t"] = round(tval, 6)
        output[f"{name} — p-value"] = round(pval, 6)

    write_results_block(sheet, "D2", "OLS Regression Results", output)


@xw.script
def run_chi_squared(book: xw.Book) -> None:
    """Run a chi-squared test of independence on a contingency table.

    Reads the contingency table range address from cell B2 on the active
    sheet, performs ``scipy.stats.chi2_contingency``, and writes the
    chi-squared statistic, p-value, degrees of freedom, and expected
    frequencies back to the sheet.

    Args:
        book: The active xlwings Book object injected by xlwings Lite.
    """
    sheet = book.sheets.active

    range_address = sheet["B2"].value
    if not range_address:
        sheet["D2"].value = "Error: Enter contingency table range address in B2."
        return

    raw = sheet[range_address].value
    if raw is None:
        sheet["D2"].value = "Error: No data found in the specified range."
        return

    # Ensure 2-D list
    if not isinstance(raw, list):
        raw = [[raw]]
    elif not isinstance(raw[0], list):
        raw = [[v] for v in raw]

    # Validate and convert cells to floats, rejecting blanks or non-numeric values
    numeric_rows = []
    for row in raw:
        numeric_row = []
        for cell in row:
            try:
                numeric_row.append(float(cell))
            except (TypeError, ValueError):
                sheet["D2"].value = (
                    "Error: Contingency table must contain only numeric values "
                    "and no blank cells."
                )
                return
        numeric_rows.append(numeric_row)

    observed = np.array(numeric_rows, dtype=float)

    chi2, p_value, dof, expected = chi2_contingency(observed)

    # Write summary statistics
    summary = OrderedDict()
    summary["Chi-Squared Statistic"] = round(chi2, 6)
    summary["p-value"] = round(p_value, 6)
    summary["Degrees of Freedom"] = int(dof)

    write_results_block(sheet, "D2", "Chi-Squared Test of Independence", summary)

    # Write expected frequencies table below the summary block
    # summary has 3 rows of data + 1 title row = offset of 5
    exp_start_row = sheet["D2"].row + 1 + len(summary) + 1  # +1 spacer
    exp_start_col = sheet["D2"].column

    sheet.cells(exp_start_row, exp_start_col).value = "Expected Frequencies"
    exp_start_row += 1

    for r_idx, row in enumerate(expected.tolist()):
        for c_idx, val in enumerate(row):
            sheet.cells(exp_start_row + r_idx, exp_start_col + c_idx).value = round(val, 4)


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


@xw.script
def run_monte_carlo(book: xw.Book) -> None:
    """Run a Monte Carlo simulation drawing from a normal distribution.

    Reads simulation parameters from cells B2–B5 on the active sheet and
    runs the specified number of simulations, each consisting of ``periods``
    draws from a normal distribution (the final value is the cumulative sum
    of all period returns).  Writes summary statistics to the sheet and
    inserts a histogram chart.

    Args:
        book: The active xlwings Book object injected by xlwings Lite.
    """
    sheet = book.sheets.active

    n_sims_raw = sheet["B2"].value
    mean_raw = sheet["B3"].value
    std_raw = sheet["B4"].value
    periods_raw = sheet["B5"].value

    n_sims = int(n_sims_raw) if n_sims_raw is not None else 10_000
    mean = float(mean_raw) if mean_raw is not None else 0.0
    std = float(std_raw) if std_raw is not None else 1.0
    periods = int(periods_raw) if periods_raw is not None else 1

    if std <= 0:
        sheet["D2"].value = "Error: Standard deviation (B4) must be greater than 0."
        return

    # Run simulation: shape = (n_sims, periods)
    draws = np.random.normal(loc=mean, scale=std, size=(n_sims, periods))
    # Final outcome = sum of all period values for each simulation
    outcomes = draws.sum(axis=1)

    mean_out = float(np.mean(outcomes))
    std_out = float(np.std(outcomes, ddof=1))
    pct5 = float(np.percentile(outcomes, 5))
    pct95 = float(np.percentile(outcomes, 95))
    min_out = float(np.min(outcomes))
    max_out = float(np.max(outcomes))

    results = OrderedDict()
    results["Simulations"] = n_sims
    results["Periods"] = periods
    results["Input Mean"] = mean
    results["Input Std Dev"] = std
    results[""] = ""  # spacer
    results["Mean Outcome"] = round(mean_out, 6)
    results["Std of Outcomes"] = round(std_out, 6)
    results["5th Percentile (VaR-style)"] = round(pct5, 6)
    results["95th Percentile"] = round(pct95, 6)
    results["Min Outcome"] = round(min_out, 6)
    results["Max Outcome"] = round(max_out, 6)

    write_results_block(sheet, "D2", "Monte Carlo Simulation Results", results)

    # Histogram chart
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.hist(outcomes, bins=50, edgecolor="white", color="steelblue", alpha=0.8)
    ax.axvline(pct5, color="red", linestyle="--", linewidth=1.5, label="5th pct")
    ax.axvline(pct95, color="green", linestyle="--", linewidth=1.5, label="95th pct")
    ax.axvline(mean_out, color="orange", linestyle="-", linewidth=2, label="Mean")
    ax.set_title("Monte Carlo Outcome Distribution")
    ax.set_xlabel("Outcome")
    ax.set_ylabel("Frequency")
    ax.legend()
    ax.grid(True, linestyle="--", alpha=0.4)
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=100)
    buf.seek(0)
    plt.close(fig)

    sheet.pictures.add(
        buf,
        name="MonteCarloHist",
        update=True,
        left=sheet["G2"].left,
        top=sheet["G2"].top,
    )


# ── Custom Excel functions (@xw.func → callable as Excel formulas) ─────────────


@xw.func
@xw.arg("y_range", doc="Dependent variable range (single column)")
@xw.arg("x_range", doc="Independent variable range (single column)")
def slope_intercept(y_range, x_range):
    """Return the slope and intercept of a simple linear regression.

    Uses ``scipy.stats.linregress`` to fit a line y = slope * x + intercept.

    Args:
        y_range: Excel range containing the dependent variable values.
        x_range: Excel range containing the independent variable values.

    Returns:
        A 1×2 array ``[[slope, intercept]]`` suitable for spilling into two
        adjacent cells in Excel, or ``[["Error", message]]`` if inputs are
        invalid.
    """
    y = _flatten(y_range)
    x = _flatten(x_range)
    if len(x) < 2 or len(y) < 2:
        return [["Error", "Need at least 2 data points"]]
    if len(x) != len(y):
        return [["Error", "X and Y must have the same length"]]
    slope, intercept, *_ = linregress(x, y)
    return [[round(slope, 8), round(intercept, 8)]]


@xw.func
@xw.arg("y_range", doc="Dependent variable range (single column)")
@xw.arg("x_range", doc="Independent variable range (single column)")
def r_squared(y_range, x_range):
    """Return the R² (coefficient of determination) of a simple linear regression.

    Args:
        y_range: Excel range containing the dependent variable values.
        x_range: Excel range containing the independent variable values.

    Returns:
        The R² value as a scalar float, or an error string if inputs are
        invalid.
    """
    y = _flatten(y_range)
    x = _flatten(x_range)
    if len(x) < 2 or len(y) < 2:
        return "Error: Need at least 2 data points"
    if len(x) != len(y):
        return "Error: X and Y must have the same length"
    _, _, r_value, *_ = linregress(x, y)
    return round(r_value ** 2, 8)


@xw.func
@xw.arg("data_range", doc="Data range to summarise")
def descriptive_stats(data_range):
    """Return a table of descriptive statistics for the supplied data range.

    Statistics returned (in order): count, mean, median, standard deviation,
    variance, minimum, maximum, skewness, and kurtosis (Fisher's definition).

    Args:
        data_range: Excel range containing the numeric data values.

    Returns:
        A 9×2 list of ``[statistic_name, value]`` rows suitable for spilling
        into a two-column block in Excel.
    """
    data = np.array(_flatten(data_range))
    if len(data) == 0:
        return [["Error", "No data"]]

    # std/var with ddof=1 require at least 2 observations
    if len(data) < 2:
        std_val = float("nan")
        var_val = float("nan")
    else:
        std_val = round(float(np.std(data, ddof=1)), 8)
        var_val = round(float(np.var(data, ddof=1)), 8)

    stats = [
        ["Count", int(len(data))],
        ["Mean", round(float(np.mean(data)), 8)],
        ["Median", round(float(np.median(data)), 8)],
        ["Std Dev", std_val],
        ["Variance", var_val],
        ["Min", round(float(np.min(data)), 8)],
        ["Max", round(float(np.max(data)), 8)],
        ["Skewness", round(float(skew(data)), 8)],
        ["Kurtosis", round(float(kurtosis(data)), 8)],
    ]
    return stats


@xw.func
@xw.arg("data_range", doc="Data range to compute z-scores for")
def z_score(data_range):
    """Return z-scores for each value in the supplied data range.

    Uses ``scipy.stats.zscore`` with ``ddof=1`` (sample standard deviation).

    Args:
        data_range: Excel range containing the numeric data values.

    Returns:
        A single-column list of z-score values in the same order as the input,
        suitable for spilling into a column in Excel.
    """
    raw = _flatten(data_range)
    if len(raw) == 0:
        return [["Error: No data"]]
    if len(raw) < 2:
        return [["Error: Need at least 2 data points for z-scores"]]

    data = np.array(raw)
    scores = scipy_zscore(data, ddof=1)
    return [[round(float(v), 8)] for v in scores]
