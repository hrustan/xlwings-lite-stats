"""
main.py — xlwings Lite entry point.
Import all scripts and custom functions here so xlwings Lite can register them.
Scripts decorated with @xw.script appear as runnable buttons in the task pane.
Functions decorated with @xw.func are callable as Excel formulas.
"""
import xlwings as xw

# Scripts (appear as buttons in the task pane)
from scripts.histogram import run_histogram
from scripts.scatterplot import run_scatterplot
from scripts.regression import run_regression
from scripts.chi_squared import run_chi_squared
from scripts.time_series import run_time_series
from scripts.monte_carlo import run_monte_carlo

# Custom functions (callable as Excel formulas)
from functions.regression_funcs import slope_intercept, r_squared
from functions.stats_funcs import descriptive_stats, z_score
