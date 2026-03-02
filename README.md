# xlwings-lite-stats

> **Statistical analysis add-in for Excel — powered by Python via WebAssembly. No local Python install required.**

---

## What Is This?

**xlwings-lite-stats** is an Excel add-in built on [xlwings Lite](https://docs.xlwings.org/en/latest/xlwings_lite.html) that replicates the functionality of a course-provided `.xlam` statistical analysis package. Python runs via [Pyodide](https://pyodide.org/) (WebAssembly) inside Excel — classmates only need to install the free xlwings Lite add-in from the Microsoft AppSource store and open the shared workbook.

---

## Prerequisites

1. Install the **xlwings Lite** add-in from the [Microsoft AppSource store](https://marketplace.microsoft.com/en-us/product/office/WA200008175) (free).
2. Open Excel (desktop or Excel for the web) and activate the add-in from the **Add-ins** ribbon.

No Python, pip, or conda installation is required on your machine.

---

## How to Load Into Excel

xlwings Lite has a built-in code editor accessible from its task pane. Follow these steps to load all modules:

1. Open the xlwings Lite task pane in Excel (**Home → xlwings → Open Task Pane**).
2. Click the **Editor** tab inside the task pane.
3. For each file listed below, create a new file in the editor using the exact filename shown, then paste the contents:

| File to create in editor | Source file in this repo |
|---|---|
| `main.py` | `main.py` |
| `requirements.txt` | `requirements.txt` |
| `utils/excel_helpers.py` | `utils/excel_helpers.py` |
| `scripts/histogram.py` | `scripts/histogram.py` |
| `scripts/scatterplot.py` | `scripts/scatterplot.py` |
| `scripts/regression.py` | `scripts/regression.py` |
| `scripts/chi_squared.py` | `scripts/chi_squared.py` |
| `scripts/time_series.py` | `scripts/time_series.py` |
| `scripts/monte_carlo.py` | `scripts/monte_carlo.py` |
| `functions/regression_funcs.py` | `functions/regression_funcs.py` |
| `functions/stats_funcs.py` | `functions/stats_funcs.py` |

4. Save each file and click **Restart** in the task pane to reload the Python environment.
5. The script buttons (Histogram, Scatter Plot, etc.) will appear in the task pane automatically.

---

## Available Tools

### Scripts (run via task pane buttons)

| Tool | Function | Sheet Setup | Output |
|---|---|---|---|
| **Histogram** | `run_histogram` | `B2`: data range address (e.g. `A2:A100`) | Bin edges & frequencies written to `D2` |
| **Scatter Plot** | `run_scatterplot` | `B2`: X range, `B3`: Y range, `B4`: title, `B5`: x-label, `B6`: y-label | Chart image inserted in sheet |
| **Regression** | `run_regression` | `B2`: Y range, `B3`: X range (single or multi-column) | Coefficients, SE, t-stats, p-values, R², adj-R², F-stat written to sheet |
| **Chi-Squared** | `run_chi_squared` | `B2`: contingency table range address | χ², p-value, df, expected frequencies written to sheet |
| **Time Series** | `run_time_series` | `B2`: series range address, `B3`: rolling window (default 3) | Rolling mean, std, trend values written to sheet; chart inserted |
| **Monte Carlo** | `run_monte_carlo` | `B2`: simulations (default 10000), `B3`: mean, `B4`: std dev, `B5`: periods (default 1) | Summary statistics written to sheet; histogram chart inserted |

### Custom Functions (callable as Excel formulas)

| Formula | Description | Example |
|---|---|---|
| `=SLOPE_INTERCEPT(y_range, x_range)` | Returns `[[slope, intercept]]` array | `=SLOPE_INTERCEPT(B2:B20, A2:A20)` |
| `=R_SQUARED(y_range, x_range)` | Returns R² of simple linear regression | `=R_SQUARED(B2:B20, A2:A20)` |
| `=DESCRIPTIVE_STATS(data_range)` | Returns 2-column array: stat name + value | `=DESCRIPTIVE_STATS(A2:A100)` |
| `=Z_SCORE(data_range)` | Returns same-shape array of z-scores | `=Z_SCORE(A2:A100)` |

---

## Project Structure

```
xlwings-lite-stats/
│
├── README.md               # This file
├── requirements.txt        # Pyodide-compatible packages
├── main.py                 # xlwings Lite entry point — registers all scripts & functions
│
├── scripts/                # @script functions (appear as task pane buttons)
│   ├── __init__.py
│   ├── histogram.py
│   ├── scatterplot.py
│   ├── regression.py
│   ├── chi_squared.py
│   ├── time_series.py
│   └── monte_carlo.py
│
├── functions/              # @func UDFs (callable as Excel formulas)
│   ├── __init__.py
│   ├── regression_funcs.py
│   └── stats_funcs.py
│
└── utils/                  # Shared helpers for Excel read/write
    ├── __init__.py
    └── excel_helpers.py
```

- **`scripts/`** — Each file contains one `@xw.script` function triggered by a task pane button. Scripts read parameters from well-known cells (e.g. `B2`, `B3`) and write results back to the sheet.
- **`functions/`** — Each file contains `@xw.func` UDFs that behave like Excel built-in functions and return arrays.
- **`utils/`** — Shared helpers (`get_range_as_df`, `write_results_block`, etc.) used across scripts to keep sheet I/O in one place.

---

## Contributing / Development

1. **Edit locally**: Clone this repo, edit the relevant `.py` file in your editor of choice.
2. **Commit**: Push your changes to GitHub as usual.
3. **Update Excel**: Open the xlwings Lite editor, navigate to the changed file, and paste the updated contents. Click **Restart** to reload.

```bash
git clone https://github.com/hrustan/xlwings-lite-stats.git
cd xlwings-lite-stats
# edit files…
git add .
git commit -m "Your change description"
git push
```

---

## Package Compatibility Note

All packages used in this project (`numpy`, `pandas`, `scipy`, `statsmodels`, `matplotlib`) are available in [Pyodide](https://pyodide.org/). For a full list of packages available in the Pyodide environment, see:

👉 https://pyodide.org/en/stable/usage/packages-in-pyodide.html

**Do NOT** add packages that are not in that list — they will not load in xlwings Lite.