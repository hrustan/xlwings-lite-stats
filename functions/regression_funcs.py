"""
functions/regression_funcs.py — Custom Excel functions for regression analysis.

These UDFs are callable directly as Excel formulas once loaded via xlwings Lite.

Examples:
  =SLOPE_INTERCEPT(B2:B20, A2:A20)   → returns [[slope, intercept]]
  =R_SQUARED(B2:B20, A2:A20)         → returns the R² scalar
"""
import xlwings as xw
from scipy.stats import linregress


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
