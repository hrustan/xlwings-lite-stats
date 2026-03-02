# xlwings-lite-stats

> **Statistical analysis add-in for Excel — powered by Python via WebAssembly. No local Python install required.**

---

## What Is This?

**xlwings-lite-stats** is an Excel add-in built on [xlwings Lite](https://docs.xlwings.org/en/latest/xlwings_lite.html) that replicates the functionality of a course-provided `.xlam` statistical analysis package. Python runs via [Pyodide](https://pyodide.org/) (WebAssembly) inside Excel — classmates only need to install the free xlwings Lite add-in from the Microsoft AppSource store and open the shared workbook. All Python scripts are embedded directly in the workbook file, so they persist across closing, restarting, and shutting down Excel.

---

## Prerequisites

1. Install the **xlwings Lite** add-in from the [Microsoft AppSource store](https://marketplace.microsoft.com/en-us/product/office/WA200008175) (free).
2. Open Excel (desktop or Excel for the web) and activate the add-in from the **Add-ins** ribbon.

No Python, pip, or conda installation is required on your machine.

---

## How to Load Into Excel

xlwings Lite stores Python files as Custom XML parts **inside** the workbook.
This means the scripts travel with the `.xlsx` file and persist across closing,
restarting, and shutting down Excel — no "Mount local folder" step required and
no extra security pop-ups for students.

> **Why Import/Export instead of Mount Local Folder?**
> Mounting a local folder requires confirming security pop-ups, may need
> multiple restarts, and the mount can be lost when the Office cache is
> cleared. With Import/Export the scripts are embedded in the workbook itself,
> so they are always available as long as the xlwings Lite add-in is installed.

### Quick Start (Single-File Import)

For the fastest path to a working workbook, use the single bundled file:

1. Download **`main_bundled.py`** from the root of this repository.
2. Open Excel and open a workbook (new or existing).
3. Open the xlwings Lite task pane: **Home → xlwings → Open Task Pane**.
4. Click the **Editor** tab → **Import** → select `main_bundled.py`.
5. Click **Restart** in the task pane.
6. **Save the workbook** (`.xlsx`). All scripts and custom functions are now embedded.

That's it — no folder structure, no multiple imports.

> For those who prefer the full modular structure (separate files per script),
> see [Instructor / Initial Setup](#instructor--initial-setup-one-time) below.

### Instructor / Initial Setup (one-time)

The instructor prepares a workbook once and shares it with the class:

1. Clone or download this repository to your computer.
2. Open a new Excel workbook (or the workbook you want to distribute).
3. Open the xlwings Lite task pane (**Home → xlwings → Open Task Pane**).
4. Click the **Editor** tab inside the task pane.
5. Use the **Import** button to import all of the following files from the
   cloned repository folder, preserving the directory structure:

   | File to import | Description |
   |---|---|
   | `main.py` | Entry point — registers all scripts & functions |
   | `requirements.txt` | Pyodide-compatible package list |
   | `utils/__init__.py` | Marks `utils` as a package (empty file) |
   | `utils/excel_helpers.py` | Shared Excel I/O helpers |
   | `scripts/__init__.py` | Marks `scripts` as a package (empty file) |
   | `scripts/histogram.py` | Histogram analysis script |
   | `scripts/scatterplot.py` | Scatter plot script |
   | `scripts/regression.py` | OLS regression script |
   | `scripts/chi_squared.py` | Chi-squared test script |
   | `scripts/time_series.py` | Time series analysis script |
   | `scripts/monte_carlo.py` | Monte Carlo simulation script |
   | `functions/__init__.py` | Marks `functions` as a package (empty file) |
   | `functions/regression_funcs.py` | Regression UDFs |
   | `functions/stats_funcs.py` | Descriptive statistics UDFs |

6. Click **Restart** in the task pane to load the Python environment and
   install the required packages.
7. Verify the script buttons (Histogram, Scatter Plot, etc.) appear in the
   task pane.
8. **Save the workbook** (`.xlsx`). The Python files are now embedded inside it.
9. Distribute the saved workbook to students (e.g. via your LMS, email, or
   shared drive).

### Student Setup

Students only need two things:

1. Install the **xlwings Lite** add-in from the
   [Microsoft AppSource store](https://marketplace.microsoft.com/en-us/product/office/WA200008175) (free, one-time).
2. Open the workbook provided by the instructor.

The scripts and custom functions load automatically — no file copying, no
folder mounting, and no security pop-ups beyond the initial add-in install.

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
3. **Update the workbook**: Open the xlwings Lite editor, use **Import** to
   re-import the changed files (or paste updated contents manually), then
   click **Restart** to reload. Save the workbook to persist the changes.

```bash
git clone https://github.com/hrustan/xlwings-lite-stats.git
cd xlwings-lite-stats
# edit files…
git add .
git commit -m "Your change description"
git push
```

> **Tip:** Use the **Export** button in the xlwings Lite editor to export files
> back to your local repo before committing, so the repo always stays in sync
> with the workbook.

---

## Package Compatibility Note

All packages used in this project (`numpy`, `pandas`, `scipy`, `statsmodels`, `matplotlib`) are available in [Pyodide](https://pyodide.org/). For a full list of packages available in the Pyodide environment, see:

👉 https://pyodide.org/en/stable/usage/packages-in-pyodide.html

**Do NOT** add packages that are not in that list — they will not load in xlwings Lite.