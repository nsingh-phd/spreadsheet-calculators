#!/usr/bin/env python3
"""
Mortgage Amortization Calculator Generator

Generates a template Excel (.xlsx) spreadsheet with formula-driven
mortgage amortization schedule supporting:
- Recurring extra payments (monthly/annual) with configurable start dates
- One-time extra payments
- Summary comparison (with vs without extra payments)

Usage:
    python generate_mortgage_calculator.py

Output:
    output/mortgage_amortization_calculator.xlsx
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.workbook.defined_name import DefinedName
from copy import copy


def add_named_range(wb, name, attr_text):
    """Add a named range compatible with openpyxl 3.1.x."""
    dn = DefinedName(name, attr_text=attr_text)
    wb.defined_names.add(dn)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
MAX_MONTHS = 360  # 30-year max
NUM_RECURRING = 10
NUM_ONETIME = 15

# Layout rows (1-indexed)
INPUT_START_ROW = 2
RECURRING_HEADER_ROW = 9
RECURRING_DATA_START = 10
RECURRING_DATA_END = RECURRING_DATA_START + NUM_RECURRING - 1  # 19
ONETIME_HEADER_ROW = RECURRING_DATA_END + 2  # 21
ONETIME_DATA_START = ONETIME_HEADER_ROW + 1  # 22
ONETIME_DATA_END = ONETIME_DATA_START + NUM_ONETIME - 1  # 36
AMORT_HEADER_ROW = ONETIME_DATA_END + 2  # 38
AMORT_DATA_START = AMORT_HEADER_ROW + 1  # 39
AMORT_DATA_END = AMORT_DATA_START + MAX_MONTHS - 1  # 398

# Summary location (to the right of inputs)
SUMMARY_COL = 5  # Column E

# Styling
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
SECTION_FONT = Font(bold=True, size=12, color="2F5496")
INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
LIGHT_BLUE_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

# Currency and date formats
CURRENCY_FMT = '#,##0.00'
PERCENT_FMT = '0.000%'
DATE_FMT = 'MMM YYYY'
NUMBER_FMT = '#,##0'


def style_header_row(ws, row, cols, fill=None, font=None):
    """Apply header styling to a range of cells in a row."""
    fill = fill or HEADER_FILL
    font = font or HEADER_FONT
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER


def style_input_cell(cell, fmt=None):
    """Style a cell as user-editable input."""
    cell.fill = INPUT_FILL
    cell.border = THIN_BORDER
    cell.alignment = Alignment(horizontal="center")
    if fmt:
        cell.number_format = fmt


def create_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "Calculator"

    # -----------------------------------------------------------------------
    # A. INPUTS SECTION
    # -----------------------------------------------------------------------
    ws.merge_cells("A1:C1")
    ws.cell(row=1, column=1, value="MORTGAGE INPUTS").font = SECTION_FONT

    labels = [
        (2, "Loan Amount ($)"),
        (3, "Annual Interest Rate"),
        (4, "Loan Term (years)"),
        (5, "Start Date"),
        (6, "Monthly Payment"),
    ]
    for row_num, label in labels:
        ws.cell(row=row_num, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row_num, column=1).alignment = Alignment(horizontal="right")

    # Input cells (B2:B5) - user editable
    style_input_cell(ws.cell(row=2, column=2, value=300000), CURRENCY_FMT)   # Loan Amount
    style_input_cell(ws.cell(row=3, column=2, value=0.05875), PERCENT_FMT)   # Interest Rate
    style_input_cell(ws.cell(row=4, column=2, value=30), NUMBER_FMT)         # Term
    style_input_cell(ws.cell(row=5, column=2), DATE_FMT)                     # Start Date

    # Set sample start date
    from datetime import date
    ws.cell(row=5, column=2).value = date(2026, 1, 1)

    # Monthly payment formula: =ROUND(-PMT(B3/12, B4*12, B2), 2)
    ws.cell(row=6, column=2).value = '=ROUND(-PMT(B3/12,B4*12,B2),2)'
    ws.cell(row=6, column=2).number_format = CURRENCY_FMT
    ws.cell(row=6, column=2).font = Font(bold=True, color="006600")
    ws.cell(row=6, column=2).border = THIN_BORDER

    # Define named ranges
    add_named_range(wb, "LoanAmount", "Calculator!$B$2")
    add_named_range(wb, "AnnualRate", "Calculator!$B$3")
    add_named_range(wb, "TermYears", "Calculator!$B$4")
    add_named_range(wb, "StartDate", "Calculator!$B$5")
    add_named_range(wb, "MonthlyPmt", "Calculator!$B$6")

    # -----------------------------------------------------------------------
    # B. RECURRING EXTRA PAYMENTS TABLE
    # -----------------------------------------------------------------------
    ws.merge_cells(f"A8:D8")
    ws.cell(row=8, column=1, value="RECURRING EXTRA PAYMENTS").font = SECTION_FONT

    rec_headers = ["#", "Start Date", "Amount ($)", "Frequency"]
    for i, h in enumerate(rec_headers, 1):
        ws.cell(row=RECURRING_HEADER_ROW, column=i, value=h)
    style_header_row(ws, RECURRING_HEADER_ROW, 4)

    for idx in range(NUM_RECURRING):
        r = RECURRING_DATA_START + idx
        ws.cell(row=r, column=1, value=idx + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=1).border = THIN_BORDER
        style_input_cell(ws.cell(row=r, column=2), DATE_FMT)       # Start Date
        style_input_cell(ws.cell(row=r, column=3), CURRENCY_FMT)   # Amount
        style_input_cell(ws.cell(row=r, column=4))                  # Frequency

    # Add data validation for frequency column
    from openpyxl.worksheet.datavalidation import DataValidation
    freq_dv = DataValidation(
        type="list",
        formula1='"Monthly,Annual"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Invalid Frequency",
        error="Please select Monthly or Annual",
    )
    freq_dv.add(f"D{RECURRING_DATA_START}:D{RECURRING_DATA_END}")
    ws.add_data_validation(freq_dv)

    # Named range for recurring table
    add_named_range(wb, "RecurringTable",
                    f"Calculator!$B${RECURRING_DATA_START}:$D${RECURRING_DATA_END}")

    # -----------------------------------------------------------------------
    # C. ONE-TIME EXTRA PAYMENTS TABLE
    # -----------------------------------------------------------------------
    ws.merge_cells(f"A{ONETIME_HEADER_ROW - 1}:C{ONETIME_HEADER_ROW - 1}")
    ws.cell(row=ONETIME_HEADER_ROW - 1, column=1, value="ONE-TIME EXTRA PAYMENTS").font = SECTION_FONT

    ot_headers = ["#", "Date", "Amount ($)"]
    for i, h in enumerate(ot_headers, 1):
        ws.cell(row=ONETIME_HEADER_ROW, column=i, value=h)
    style_header_row(ws, ONETIME_HEADER_ROW, 3)

    for idx in range(NUM_ONETIME):
        r = ONETIME_DATA_START + idx
        ws.cell(row=r, column=1, value=idx + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=1).border = THIN_BORDER
        style_input_cell(ws.cell(row=r, column=2), DATE_FMT)       # Date
        style_input_cell(ws.cell(row=r, column=3), CURRENCY_FMT)   # Amount

    add_named_range(wb, "OneTimeTable",
                    f"Calculator!$B${ONETIME_DATA_START}:$C${ONETIME_DATA_END}")

    # -----------------------------------------------------------------------
    # D. AMORTIZATION SCHEDULE
    # -----------------------------------------------------------------------
    ws.merge_cells(f"A{AMORT_HEADER_ROW - 1}:F{AMORT_HEADER_ROW - 1}")
    ws.cell(row=AMORT_HEADER_ROW - 1, column=1, value="AMORTIZATION SCHEDULE").font = SECTION_FONT

    amort_headers = [
        "Payment #", "Date", "Payment", "Interest", "Principal", "Ending Balance",
    ]
    num_amort_cols = len(amort_headers)
    for i, h in enumerate(amort_headers, 1):
        ws.cell(row=AMORT_HEADER_ROW, column=i, value=h)
    style_header_row(ws, AMORT_HEADER_ROW, num_amort_cols)

    # --- Build formulas for each row ---
    # Layout: A=Payment#, B=Date, C=Payment, D=Interest, E=Principal, F=Ending Balance
    # Principal combines scheduled principal + extra payments (capped so balance >= 0)
    for idx in range(MAX_MONTHS):
        r = AMORT_DATA_START + idx
        n = idx + 1  # payment number
        prev_r = r - 1

        # Reference to beginning balance (previous row's ending balance, or LoanAmount)
        if idx == 0:
            bal_ref = "LoanAmount"
        else:
            bal_ref = f"F{prev_r}"

        # Col A: Payment # — only show if there's balance remaining
        if idx == 0:
            ws.cell(row=r, column=1, value=f'=IF(LoanAmount>0,{n},"")')
        else:
            ws.cell(row=r, column=1, value=f'=IF(AND(F{prev_r}>0.005,A{prev_r}<>""),{n},"")')

        # Col B: Date — StartDate + n-1 months using EDATE
        if idx == 0:
            ws.cell(row=r, column=2, value=f'=IF(A{r}<>"",StartDate,"")')
        else:
            ws.cell(row=r, column=2, value=f'=IF(A{r}<>"",EDATE(StartDate,{n - 1}),"")')
        ws.cell(row=r, column=2).number_format = DATE_FMT

        # Col D: Interest = Beginning Balance * Monthly Rate
        ws.cell(row=r, column=4, value=f'=IF(A{r}<>"",ROUND({bal_ref}*AnnualRate/12,2),"")')
        ws.cell(row=r, column=4).number_format = CURRENCY_FMT

        # Col C: Payment — min of monthly payment and (balance + interest)
        ws.cell(row=r, column=3, value=f'=IF(A{r}<>"",MIN(MonthlyPmt,{bal_ref}+D{r}),"")')
        ws.cell(row=r, column=3).number_format = CURRENCY_FMT

        # Col E: Principal = scheduled principal + extra payments, capped at balance
        # scheduled_principal = C{r} - D{r}
        # extra = SUMPRODUCT lookups (recurring monthly + recurring annual + one-time)
        # total = MIN(balance, scheduled_principal + extra)
        rec_start_range = f"$B${RECURRING_DATA_START}:$B${RECURRING_DATA_END}"
        rec_amount_range = f"$C${RECURRING_DATA_START}:$C${RECURRING_DATA_END}"
        rec_freq_range = f"$D${RECURRING_DATA_START}:$D${RECURRING_DATA_END}"
        ot_date_range = f"$B${ONETIME_DATA_START}:$B${ONETIME_DATA_END}"
        ot_amount_range = f"$C${ONETIME_DATA_START}:$C${ONETIME_DATA_END}"

        extra_sum = (
            # Recurring monthly
            f'SUMPRODUCT(({rec_start_range}<>"")*({rec_start_range}<=B{r})*({rec_freq_range}="Monthly")*{rec_amount_range})'
            f'+'
            # Recurring annual (month must match)
            f'SUMPRODUCT(({rec_start_range}<>"")*({rec_start_range}<=B{r})*({rec_freq_range}="Annual")*(MONTH({rec_start_range})=MONTH(B{r}))*{rec_amount_range})'
            f'+'
            # One-time (year and month match)
            f'SUMPRODUCT(({ot_date_range}<>"")*(YEAR({ot_date_range})=YEAR(B{r}))*(MONTH({ot_date_range})=MONTH(B{r}))*{ot_amount_range})'
        )

        principal_formula = (
            f'=IF(A{r}<>"",'
            f'MIN({bal_ref},(C{r}-D{r})+{extra_sum}),'
            f'"")'
        )
        ws.cell(row=r, column=5, value=principal_formula)
        ws.cell(row=r, column=5).number_format = CURRENCY_FMT

        # Col F: Ending Balance = Beginning Balance - Principal
        ws.cell(row=r, column=6, value=f'=IF(A{r}<>"",ROUND({bal_ref}-E{r},2),"")')
        ws.cell(row=r, column=6).number_format = CURRENCY_FMT

        # Light alternate row shading
        if idx % 2 == 1:
            for c in range(1, num_amort_cols + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = PatternFill(start_color="F2F7FB", end_color="F2F7FB", fill_type="solid")

        # Borders for all amortization cells
        for c in range(1, num_amort_cols + 1):
            ws.cell(row=r, column=c).border = THIN_BORDER
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")

    # -----------------------------------------------------------------------
    # E. SUMMARY SECTION (to the right of inputs, columns E-H)
    # -----------------------------------------------------------------------
    sc = SUMMARY_COL  # E

    ws.merge_cells(f"{get_column_letter(sc)}1:{get_column_letter(sc + 2)}1")
    ws.cell(row=1, column=sc, value="LOAN SUMMARY").font = SECTION_FONT

    summary_items = [
        (2, "Original Payoff Date"),
        (3, "Accelerated Payoff Date"),
        (4, "Accelerated Payoff Time"),
        (5, "Original Total Interest"),
        (6, "Accelerated Total Interest"),
        (7, "Interest Saved"),
        (8, "Months Saved"),
        (9, "Total of Payments (P+I)"),
        (10, "Total Extra Payments"),
    ]
    for row_num, label in summary_items:
        cell = ws.cell(row=row_num, column=sc, value=label)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="right")

    val_col = sc + 1  # F

    # Original payoff date: StartDate + (TermYears * 12 - 1) months
    ws.cell(row=2, column=val_col, value='=EDATE(StartDate,TermYears*12-1)')
    ws.cell(row=2, column=val_col).number_format = DATE_FMT

    # Accelerated payoff date: last non-blank date in amort table
    ws.cell(row=3, column=val_col,
            value=f'=LOOKUP(2,1/(B{AMORT_DATA_START}:B{AMORT_DATA_END}<>""),B{AMORT_DATA_START}:B{AMORT_DATA_END})')
    ws.cell(row=3, column=val_col).number_format = DATE_FMT

    # Accelerated payoff time in years and months
    # actual_months = count of payment rows with a number
    # years = INT(actual_months/12), remaining_months = MOD(actual_months,12)
    actual_months_formula = f'SUMPRODUCT(1*(ISNUMBER(A{AMORT_DATA_START}:A{AMORT_DATA_END})))'
    ws.cell(row=4, column=val_col,
            value=f'=INT({actual_months_formula}/12)&" years, "&MOD({actual_months_formula},12)&" months"')

    # Original total interest: PMT * term_months - loan_amount
    ws.cell(row=5, column=val_col,
            value='=ROUND(MonthlyPmt*TermYears*12-LoanAmount,2)')
    ws.cell(row=5, column=val_col).number_format = CURRENCY_FMT

    # Accelerated total interest: sum of interest column (now col D)
    ws.cell(row=6, column=val_col,
            value=f'=ROUND(SUM(D{AMORT_DATA_START}:D{AMORT_DATA_END}),2)')
    ws.cell(row=6, column=val_col).number_format = CURRENCY_FMT

    # Interest saved
    ws.cell(row=7, column=val_col, value=f'={get_column_letter(val_col)}5-{get_column_letter(val_col)}6')
    ws.cell(row=7, column=val_col).number_format = CURRENCY_FMT
    ws.cell(row=7, column=val_col).font = Font(bold=True, color="006600", size=12)

    # Months saved: original months - actual months (count only numeric payment #s)
    ws.cell(row=8, column=val_col,
            value=f'=TermYears*12-SUMPRODUCT(1*(ISNUMBER(A{AMORT_DATA_START}:A{AMORT_DATA_END})))')
    ws.cell(row=8, column=val_col).number_format = NUMBER_FMT
    ws.cell(row=8, column=val_col).font = Font(bold=True, color="006600", size=12)

    # Total of payments (principal + interest) = sum of payment column (now col C)
    ws.cell(row=9, column=val_col,
            value=f'=ROUND(SUM(C{AMORT_DATA_START}:C{AMORT_DATA_END}),2)')
    ws.cell(row=9, column=val_col).number_format = CURRENCY_FMT

    # Total extra payments = total principal paid - (total payments - total interest)
    # Extra = Principal (col E) - (Payment (col C) - Interest (col D)) = E_sum - C_sum + D_sum
    ws.cell(row=10, column=val_col,
            value=f'=ROUND(SUM(E{AMORT_DATA_START}:E{AMORT_DATA_END})-SUM(C{AMORT_DATA_START}:C{AMORT_DATA_END})+SUM(D{AMORT_DATA_START}:D{AMORT_DATA_END}),2)')
    ws.cell(row=10, column=val_col).number_format = CURRENCY_FMT

    # Style summary values
    for row_num in range(2, 11):
        cell = ws.cell(row=row_num, column=val_col)
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")

    # -----------------------------------------------------------------------
    # F. FORMATTING & FINAL TOUCHES
    # -----------------------------------------------------------------------

    # Column widths
    col_widths = {
        1: 22,  # A - labels / payment #
        2: 18,  # B - dates / amounts
        3: 18,  # C - payment / amounts
        4: 18,  # D - interest / frequency
        5: 26,  # E - principal / summary labels
        6: 22,  # F - ending balance / summary values
    }
    for col, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    # Freeze panes at amortization header
    ws.freeze_panes = f"A{AMORT_DATA_START}"

    # Conditional formatting: gray out rows where payment # is blank (post-payoff)
    gray_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    gray_font = Font(color="999999")

    return wb, ws


def main():
    wb, ws = create_workbook()

    # Save
    output_dir = os.path.join(os.path.dirname(__file__), "output")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "mortgage_amortization_calculator.xlsx")
    wb.save(output_path)
    print(f"✅ Mortgage calculator generated: {output_path}")
    print(f"   Open in Excel and edit the yellow input cells.")
    print(f"   - Loan inputs: B2–B5")
    print(f"   - Recurring extra payments: rows {RECURRING_DATA_START}–{RECURRING_DATA_END}")
    print(f"   - One-time extra payments: rows {ONETIME_DATA_START}–{ONETIME_DATA_END}")


if __name__ == "__main__":
    main()
