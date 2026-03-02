# Copilot Instructions for xlwings-lite-stats

## Project Context
This is an xlwings Lite add-in for Excel. Python runs via Pyodide (WebAssembly) inside the xlwings Lite task pane — there is NO local Python interpreter. All code must be compatible with the Pyodide environment.

## Pyodide Constraints
- Only packages available in Pyodide can be used. See: https://pyodide.org/en/stable/usage/packages-in-pyodide.html
- Confirmed available: numpy, pandas, scipy, statsmodels, matplotlib
- Do NOT use: subprocess, multiprocessing, threading, os.system, file I/O (open/write to disk), socket, or any package requiring native C extensions not in Pyodide
- Do NOT use f-string debug prints or logging to stdout — write outputs back to the Excel sheet

## Code Patterns
- Scripts (run via buttons): decorated with `@xw.script` or imported in main.py; always accept `book: xw.Book` as first argument
- Custom functions (Excel formulas): decorated with `@xw.func`; use `@xw.arg` for type conversion
- Use `utils/excel_helpers.py` for all Excel read/write operations — do not inline sheet access in scripts
- All range addresses are read from well-known cells (e.g., `sheet["B2"].value`) — never hardcode data ranges

## Testing
- Test logic by mocking `xw.Book` or by running the pure Python logic (numpy/scipy/statsmodels calls) in isolation
- Each script's computation logic should be extractable into a pure function for unit testing

## Style
- All functions must have docstrings
- Use type hints where possible
- Keep each script file focused on one analysis type
