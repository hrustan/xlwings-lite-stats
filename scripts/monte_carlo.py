"""
scripts/monte_carlo.py — Monte Carlo simulation script for xlwings Lite.

Sheet setup:
  B2 : Number of simulations (int, default 10000)
  B3 : Mean of the normal distribution (float)
  B4 : Standard deviation of the normal distribution (float)
  B5 : Number of periods per simulation path (int, default 1)

Output written starting at D2:
  Summary statistics block (mean, std, 5th pct, 95th pct, min, max).
  A histogram chart of simulation outcomes is also inserted.
"""
import io
from collections import OrderedDict

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import xlwings as xw

from utils.excel_helpers import write_results_block


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
