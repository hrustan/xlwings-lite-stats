"""
scripts/regression.py — OLS regression analysis script for xlwings Lite.

Sheet setup:
  B2 : Y (dependent variable) range address (e.g. "B2:B51")
  B3 : X (independent variable(s)) range address (e.g. "A2:A51" or "A2:C51")

Output written starting at D2:
  A formatted results block with coefficients, standard errors, t-stats,
  p-values, R², adjusted R², and F-statistic.
"""
from collections import OrderedDict

import numpy as np
import statsmodels.api as sm
import xlwings as xw

from utils.excel_helpers import get_range_as_list, write_results_block


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

    y = get_range_as_list(sheet, y_address)

    # Read X — may be single or multi-column
    x_raw = sheet[x_address].value
    if x_raw is None:
        sheet["D2"].value = "Error: No data found in the X range."
        return

    # Normalise to a 2-D list of rows
    if not isinstance(x_raw, list):
        x_raw = [[x_raw]]
    elif not isinstance(x_raw[0], list):
        x_raw = [[v] for v in x_raw]

    x_data = np.array(
        [[float(cell) for cell in row] for row in x_raw if any(v is not None for v in row)]
    )
    y_arr = np.array(y)

    if len(y_arr) != len(x_data):
        sheet["D2"].value = "Error: Y and X ranges must have the same number of rows."
        return

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
