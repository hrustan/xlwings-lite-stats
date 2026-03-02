# Agent Context — xlwings-lite-stats

## Purpose

This repository is a **Python-in-Excel statistical analysis add-in** built on [xlwings Lite](https://docs.xlwings.org/en/latest/xlwings_lite.html). It replaces a legacy VBA-based `.xlam` add-in that was distributed as part of the **Data Analysis and Decision Making (DADM)** course at UT. The original `.xlam` repo is here:

👉 **Original add-in:** <https://github.com/TexDS/DADM_UT>

Instead of VBA macros, this project runs Python via [Pyodide](https://pyodide.org/) (WebAssembly) directly inside Excel. Users only need the free xlwings Lite add-in from the Microsoft AppSource store — no local Python installation is required.

## What the Add-in Provides

Six **script-based tools** (triggered by task pane buttons) and four **custom Excel functions** (UDFs callable as formulas):

| Scripts | Custom Functions |
|---|---|
| Histogram | `=DESCRIPTIVE_STATS(data)` |
| Scatter Plot | `=Z_SCORE(data)` |
| Linear / Multi-variable Regression | `=SLOPE_INTERCEPT(y, x)` |
| Chi-Squared Test | `=R_SQUARED(y, x)` |
| Time Series (rolling mean & trend) | |
| Monte Carlo Simulation | |

## Architecture

- **`main.py`** — Entry point; registers all scripts and functions with xlwings Lite.
- **`scripts/`** — One file per analysis tool. Each exports a `@xw.script` function that reads config from well-known cells (e.g., `B2`, `B3`) and writes results back to the sheet.
- **`functions/`** — `@xw.func` UDFs that behave like built-in Excel formulas and return arrays.
- **`utils/excel_helpers.py`** — Shared helpers for bulk range reads/writes (`get_range_as_df`, `write_results_block`, etc.).

## Key Constraints

- **Pyodide only** — all packages must be available in the Pyodide environment. Confirmed available: `numpy`, `pandas`, `scipy`, `statsmodels`, `matplotlib`. See the full list at <https://pyodide.org/en/stable/usage/packages-in-pyodide.html>.
- **No local OS access** — `subprocess`, `multiprocessing`, `threading`, `os.system`, file I/O to disk, and sockets are all unavailable.
- **No stdout logging** — write all output back to the Excel sheet.
- **Range addresses are never hardcoded** — they are always read from well-known config cells.

For detailed coding patterns, style rules, and testing guidance, see [`.github/copilot-instructions.md`](.github/copilot-instructions.md).
