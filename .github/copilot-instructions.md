# Copilot Instructions

## Project Overview

This is a collection of Python-generated Excel spreadsheet calculators for personal finance. Each calculator lives in its own subdirectory and consists of a Python script that produces a standalone `.xlsx` file driven entirely by Excel formulas — no macros or VBA.

## Build Commands

```bash
# Install dependencies (uses uv)
uv pip install -r mortgage-amortization-calculator/requirements.txt

# Generate a calculator spreadsheet
uv run mortgage-amortization-calculator/generate_mortgage_calculator.py
```

There are no tests or linters configured.

## Architecture

Each calculator follows the same pattern:

1. **Generator script** (`generate_<name>.py`) — A single Python file using `openpyxl` that programmatically builds an entire `.xlsx` workbook: layout, styling, formulas, named ranges, data validations, conditional formatting, and charts.
2. **Template output** (`<name>.xlsx`) — The generated spreadsheet committed to the repo. Users open it in Excel and edit yellow-highlighted input cells; everything else is formula-driven.
3. **Local copy** (`<name>_local.xlsx`) — A gitignored copy for personal use, generated alongside the template.

The generator scripts do **not** perform calculations themselves. All math lives in Excel formulas written as strings (e.g., `'=ROUND(-PMT(B3/12,B4*12,B2),2)'`). The Python code is purely a workbook builder.

## Key Conventions

- **Input cells** are styled with a yellow fill (`INPUT_FILL`) so users know which cells to edit.
- **Layout is coordinate-based** — cell positions are defined via row/column constants at the top of each generator (e.g., `AMORT_COL = 8`, `RECURRING_DATA_START = 22`). When modifying layout, update these constants and all formulas that reference them.
- **Formulas reference cells by address** — since openpyxl writes formulas as strings, changing the layout means manually updating every formula string that references moved cells.
- **Named ranges** are used to make formulas more readable and are added via the `add_named_range` helper.
- **Two outputs per run** — `main()` saves both a tracked template and a `_local.xlsx` copy (gitignored via the `*_local.xlsx` pattern).
- **After any code change**, re-run the generator and open the `.xlsx` in Excel to verify formulas and layout are correct.
