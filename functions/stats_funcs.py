"""
functions/stats_funcs.py — Custom Excel functions for descriptive statistics.

These UDFs are callable directly as Excel formulas once loaded via xlwings Lite.

Examples:
  =DESCRIPTIVE_STATS(A2:A100)   → returns a 9×2 array of stat name + value
  =Z_SCORE(A2:A100)             → returns a same-length array of z-scores
"""
import numpy as np
import xlwings as xw
from scipy.stats import skew, kurtosis, zscore as scipy_zscore


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

    stats = [
        ["Count", int(len(data))],
        ["Mean", round(float(np.mean(data)), 8)],
        ["Median", round(float(np.median(data)), 8)],
        ["Std Dev", round(float(np.std(data, ddof=1)), 8)],
        ["Variance", round(float(np.var(data, ddof=1)), 8)],
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

    data = np.array(raw)
    scores = scipy_zscore(data, ddof=1)
    return [[round(float(v), 8)] for v in scores]
