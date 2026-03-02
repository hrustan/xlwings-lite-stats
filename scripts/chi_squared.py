"""
scripts/chi_squared.py — Chi-squared test of independence script for xlwings Lite.

Sheet setup:
  B2 : Contingency table range address (e.g. "A2:C4")
       The cell B2 should contain the range address as a string.

Output written starting at D2:
  Chi-squared statistic, p-value, degrees of freedom, and expected
  frequencies table.
"""
from collections import OrderedDict

import numpy as np
from scipy.stats import chi2_contingency
import xlwings as xw

from utils.excel_helpers import write_results_block


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
