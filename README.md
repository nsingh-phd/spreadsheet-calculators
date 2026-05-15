# Spreadsheet Calculators

A collection of Excel spreadsheet tools for personal finance and everyday calculations. Each calculator is a standalone `.xlsx` template driven entirely by formulas — no macros or external scripts needed after generation.

## Calculators

### [Mortgage Amortization Calculator](mortgage-amortization-calculator/)

A full-featured mortgage amortization schedule with support for extra payments.

**Features:**
- Standard amortization schedule (up to 30-year terms)
- Recurring extra payments (monthly or annual) with configurable start dates
- One-time lump-sum payments
- Summary dashboard: payoff date, interest saved, months saved, and more

**Quick start:**
1. Install dependencies: `uv pip install -r mortgage-amortization-calculator/requirements.txt`
2. Generate the spreadsheet: `uv run mortgage-amortization-calculator/generate_mortgage_calculator.py`
3. Open `mortgage-amortization-calculator/mortgage_amortization_calculator.xlsx` in Excel
4. Edit the yellow input cells and the schedule updates automatically

---

*More calculators coming soon.*
