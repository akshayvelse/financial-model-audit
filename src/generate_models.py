"""Programmatic generator for flawed and audited financial valuation models.

Constructs two institutional Excel workbooks using openpyxl:
1. target_model_v1_flawed.xlsx: Inherited model containing 5 high-impact structural defects.
2. target_model_v2_audited.xlsx: Production-grade corrected model resolving all discrepancies.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"

COLOR_NAVY = "1B365D"
COLOR_HEADER_TEXT = "FFFFFF"
COLOR_INPUT = "0000FF"
COLOR_CALC = "000000"
COLOR_LINK = "006100"
COLOR_MUTED = "F2F5F9"
COLOR_ERROR_HIGHLIGHT = "FFF0F0"

FONT_NAME = "Calibri"
FONT_HEADER = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_HEADER_TEXT)
FONT_SECTION = Font(name=FONT_NAME, size=11, bold=True, color="000000")
FONT_INPUT = Font(name=FONT_NAME, size=10, bold=False, color=COLOR_INPUT)
FONT_CALC = Font(name=FONT_NAME, size=10, bold=False, color=COLOR_CALC)
FONT_LINK = Font(name=FONT_NAME, size=10, bold=False, color=COLOR_LINK)
FONT_TOTAL = Font(name=FONT_NAME, size=10, bold=True, color=COLOR_CALC)

FILL_NAVY = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
FILL_MUTED = PatternFill(start_color=COLOR_MUTED, end_color=COLOR_MUTED, fill_type="solid")
FILL_ERROR = PatternFill(start_color=COLOR_ERROR_HIGHLIGHT, end_color=COLOR_ERROR_HIGHLIGHT, fill_type="solid")

BORDER_TOTAL = Border(
    top=Side(style="thin", color="000000"),
    bottom=Side(style="double", color="000000"),
)

FORMAT_CURRENCY = "$#,##0"
FORMAT_PRICE = "$#,##0.00"
FORMAT_PERCENT = "0.0%"
FORMAT_PERCENT_TWO = "0.00%"


def format_grid(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    """Ensure standard grid lines and auto-fit column dimensions."""
    ws.views.sheetView[0].showGridLines = True
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            v = str(cell.value or "")
            if len(v) > max_len and not v.startswith("="):
                max_len = len(v)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)


def build_audit_model_workbook(is_flawed: bool = False) -> openpyxl.Workbook:
    """Construct complete financial valuation model with or without structural defects."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DCF Valuation"
    ws.views.sheetView[0].showGridLines = True

    model_title = "TARGET COMPANY VALUATION MODEL (AUDIT V1 - INHERITED)" if is_flawed else "TARGET COMPANY VALUATION MODEL (AUDIT V2 - AUDITED)"
    ws["B2"] = model_title
    ws["B2"].font = Font(name=FONT_NAME, size=13, bold=True, color=COLOR_NAVY)
    ws["B3"] = "Values in thousands of USD ($000s), unless specified otherwise"
    ws["B3"].font = Font(name=FONT_NAME, size=10, italic=True)

    # 1. Parameter Block
    ws["B5"] = "VALUATION PARAMETERS"
    ws["B5"].font = FONT_SECTION

    headers_param = ["Parameter", "Value", "Notes"]
    for c_idx, h in enumerate(headers_param, start=2):
        c = ws.cell(row=6, column=c_idx, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_NAVY

    params = [
        ("WACC (Discount Rate)", 0.0950, FORMAT_PERCENT_TWO, "Base cost of capital"),
        ("Terminal Perpetuity Growth Rate (g)", 0.0250, FORMAT_PERCENT_TWO, "Steady-state growth proxy"),
        ("Effective Corporate Tax Rate", 0.2100, FORMAT_PERCENT, "Federal + state blended statutory"),
        ("Cash and Marketable Securities", 45000, FORMAT_CURRENCY, "Balance sheet cash & liquid equivalents"),
        ("Total Debt Outstanding", 85000, FORMAT_CURRENCY, "Term loans + senior notes"),
        ("Diluted Shares Outstanding (000s)", 15000, "#,##0", "Treasury stock method diluted shares"),
    ]

    for idx, (p_name, val, fmt, note) in enumerate(params, start=7):
        c_lbl = ws.cell(row=idx, column=2, value=p_name)
        c_val = ws.cell(row=idx, column=3, value=val)
        c_not = ws.cell(row=idx, column=4, value=note)

        c_val.number_format = fmt
        c_val.font = FONT_INPUT
        c_val.fill = FILL_MUTED
        c_not.font = Font(name=FONT_NAME, size=9, color="555555")

    # 2. Operating Forecast Block
    ws["B14"] = "5-YEAR OPERATING FORECAST & FREE CASH FLOW TO FIRM"
    ws["B14"].font = FONT_SECTION

    headers_forecast = ["Financial Metric", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Terminal Year"]
    for c_idx, h in enumerate(headers_forecast, start=2):
        c = ws.cell(row=15, column=c_idx, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_NAVY
        c.alignment = Alignment(horizontal="left" if c_idx == 2 else "right")

    # Time indices row
    ws.cell(row=16, column=2, value="Period Index (t)").font = FONT_SECTION
    for idx, t_val in enumerate([1, 2, 3, 4, 5, "TV"], start=3):
        c = ws.cell(row=16, column=idx, value=t_val)
        c.font = FONT_CALC
        c.alignment = Alignment(horizontal="right")

    # Setup forecast data
    base_revenue = 250000  # $250M
    growth_rates = [0.080, 0.075, 0.070, 0.065, 0.060]
    cogs_pct = 0.580
    opex_pct = 0.220
    da_pct = 0.050
    capex_pct = 0.060
    nwc_pct = 0.015

    # Row definitions
    # Row 17: Total Revenue
    # Row 18: Cost of Goods Sold (COGS)
    # Row 19: Gross Profit
    # Row 20: Operating Expenses (SG&A + R&D)
    # Row 21: Operating Income (EBIT)
    # Row 22: Tax on EBIT
    # Row 23: Net Operating Profit After Tax (NOPAT)
    # Row 24: Depreciation & Amortization (D&A)
    # Row 25: Capital Expenditures (CapEx)
    # Row 26: Change in Non-Cash Working Capital (Delta NWC)
    # Row 27: Free Cash Flow to Firm (FCFF)
    # Row 28: Mid-Year Discount Factor
    # Row 29: Present Value of FCFF

    # Total Revenue (Row 17)
    ws.cell(row=17, column=2, value="Total Revenue").font = FONT_TOTAL
    curr_rev = base_revenue
    for idx in range(3, 8):
        col_let = get_column_letter(idx)
        prev_let = get_column_letter(idx - 1)
        curr_rev *= 1.0 + growth_rates[idx - 3]
        if idx == 3:
            cell = ws.cell(row=17, column=idx, value=curr_rev)
            cell.font = FONT_INPUT
        else:
            g_rate = growth_rates[idx - 3]
            cell = ws.cell(row=17, column=idx, value=f"={prev_let}17*(1+{g_rate})")
            cell.font = FONT_CALC
        cell.number_format = FORMAT_CURRENCY
        cell.alignment = Alignment(horizontal="right")

    # Terminal Revenue
    ws.cell(row=17, column=8, value="=G17*(1+$C$8)").font = FONT_CALC
    ws.cell(row=17, column=8).number_format = FORMAT_CURRENCY
    ws.cell(row=17, column=8).alignment = Alignment(horizontal="right")

    # COGS (Row 18) - ERROR 3: Hardcoded override in Year 4 if is_flawed
    ws.cell(row=18, column=2, value="Cost of Goods Sold (COGS)")
    for idx in range(3, 9):
        col_let = get_column_letter(idx)
        cell = ws.cell(row=18, column=idx)
        if is_flawed and idx == 6:  # Year 4 (Column F)
            # Hardcoded override error: static 185000 instead of =F17*0.58
            cell.value = 185000
            cell.font = FONT_INPUT
            cell.fill = FILL_ERROR
        else:
            cell.value = f"={col_let}17*{cogs_pct}"
            cell.font = FONT_CALC
        cell.number_format = FORMAT_CURRENCY
        cell.alignment = Alignment(horizontal="right")

    # Gross Profit (Row 19)
    ws.cell(row=19, column=2, value="Gross Profit").font = FONT_SECTION
    for idx in range(3, 9):
        col_let = get_column_letter(idx)
        cell = ws.cell(row=19, column=idx, value=f"={col_let}17-{col_let}18")
        cell.font = FONT_CALC
        cell.number_format = FORMAT_CURRENCY
        cell.border = Border(top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"))
        cell.alignment = Alignment(horizontal="right")

    # Operating Expenses (Row 20)
    ws.cell(row=20, column=2, value="Operating Expenses (SG&A)")
    for idx in range(3, 9):
        col_let = get_column_letter(idx)
        cell = ws.cell(row=20, column=idx, value=f"={col_let}17*{opex_pct}")
        cell.font = FONT_CALC
        cell.number_format = FORMAT_CURRENCY
        cell.alignment = Alignment(horizontal="right")

    # EBIT (Row 21)
    ws.cell(row=21, column=2, value="Operating Income (EBIT)").font = FONT_SECTION
    for idx in range(3, 9):
        col_let = get_column_letter(idx)
        cell = ws.cell(row=21, column=idx, value=f"={col_let}19-{col_let}20")
        cell.font = FONT_CALC
        cell.number_format = FORMAT_CURRENCY
        cell.alignment = Alignment(horizontal="right")

    # Tax on EBIT (Row 22)
    ws.cell(row=22, column=2, value="Effective Tax on EBIT")
    for idx in range(3, 9):
        col_let = get_column_letter(idx)
        cell = ws.cell(row=22, column=idx, value=f"={col_let}21*$C$9")
        cell.font = FONT_CALC
        cell.number_format = FORMAT_CURRENCY
        cell.alignment = Alignment(horizontal="right")

    # NOPAT (Row 23)
    ws.cell(row=23, column=2, value="Net Operating Profit After Tax (NOPAT)").font = FONT_SECTION
    for idx in range(3, 9):
        col_let = get_column_letter(idx)
        cell = ws.cell(row=23, column=idx, value=f"={col_let}21-{col_let}22")
        cell.font = FONT_TOTAL
        cell.number_format = FORMAT_CURRENCY
        cell.border = Border(top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"))
        cell.alignment = Alignment(horizontal="right")

    # D&A (Row 24)
    ws.cell(row=24, column=2, value="(+) Depreciation and Amortization")
    for idx in range(3, 9):
        col_let = get_column_letter(idx)
        cell = ws.cell(row=24, column=idx, value=f"={col_let}17*{da_pct}")
        cell.font = FONT_CALC
        cell.number_format = FORMAT_CURRENCY
        cell.alignment = Alignment(horizontal="right")

    # CapEx (Row 25) - ERROR 2: Terminal Year CapEx / D&A Imbalance if is_flawed
    ws.cell(row=25, column=2, value="(-) Capital Expenditures (CapEx)")
    for idx in range(3, 8):  # Years 1-5
        col_let = get_column_letter(idx)
        cell = ws.cell(row=25, column=idx, value=f"={col_let}17*{capex_pct}")
        cell.font = FONT_CALC
        cell.number_format = FORMAT_CURRENCY
        cell.alignment = Alignment(horizontal="right")

    # Terminal CapEx (Column H, Row 25)
    cell_term_capex = ws.cell(row=25, column=8)
    if is_flawed:
        # Error 2: CapEx projected at 2.5x D&A in perpetuity (violates steady-state reinvestment math)
        cell_term_capex.value = "=H24*2.5"
        cell_term_capex.fill = FILL_ERROR
    else:
        # Audited: CapEx aligned with D&A + steady-state real growth (1.08x D&A)
        cell_term_capex.value = "=H24*1.08"
    cell_term_capex.font = FONT_CALC
    cell_term_capex.number_format = FORMAT_CURRENCY
    cell_term_capex.alignment = Alignment(horizontal="right")

    # Delta NWC (Row 26)
    ws.cell(row=26, column=2, value="(-) Change in Non-Cash Working Capital")
    for idx in range(3, 9):
        col_let = get_column_letter(idx)
        cell = ws.cell(row=26, column=idx, value=f"={col_let}17*{nwc_pct}")
        cell.font = FONT_CALC
        cell.number_format = FORMAT_CURRENCY
        cell.alignment = Alignment(horizontal="right")

    # FCFF (Row 27) - ERROR 1: Working Capital Sign Flip if is_flawed
    ws.cell(row=27, column=2, value="Free Cash Flow to Firm (FCFF)").font = FONT_TOTAL
    for idx in range(3, 9):
        col_let = get_column_letter(idx)
        cell = ws.cell(row=27, column=idx)
        if is_flawed and idx in [4, 5]:  # Year 2 and Year 3 sign flip error
            # Error 1: Adding instead of subtracting Delta NWC (+ Delta NWC instead of - Delta NWC)
            cell.value = f"={col_let}23+{col_let}24-{col_let}25+{col_let}26"
            cell.fill = FILL_ERROR
        else:
            cell.value = f"={col_let}23+{col_let}24-{col_let}25-{col_let}26"
        cell.font = FONT_TOTAL
        cell.number_format = FORMAT_CURRENCY
        cell.border = BORDER_TOTAL
        cell.alignment = Alignment(horizontal="right")

    # Mid-Year Discount Factor (Row 28)
    ws.cell(row=28, column=2, value="Mid-Year Discount Factor")
    for idx in range(3, 8):  # Years 1-5
        col_let = get_column_letter(idx)
        cell = ws.cell(row=28, column=idx, value=f"=1/(1+$C$7)^({col_let}16-0.5)")
        cell.font = FONT_CALC
        cell.number_format = "0.0000"
        cell.alignment = Alignment(horizontal="right")

    # Present Value of FCFF (Row 29)
    ws.cell(row=29, column=2, value="Present Value of Explicit FCFF").font = FONT_SECTION
    for idx in range(3, 8):
        col_let = get_column_letter(idx)
        cell = ws.cell(row=29, column=idx, value=f"={col_let}27*{col_let}28")
        cell.font = FONT_TOTAL
        cell.number_format = FORMAT_CURRENCY
        cell.border = Border(top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"))
        cell.alignment = Alignment(horizontal="right")

    # 3. Valuation Bridge & Terminal Value Summary
    ws["B32"] = "VALUATION BRIDGE & EQUITY VALUE DERIVATION"
    ws["B32"].font = FONT_SECTION

    # Cumulative PV
    ws.cell(row=33, column=2, value="Cumulative PV of Explicit FCFF (Years 1-5)").font = FONT_SECTION
    ws.cell(row=33, column=3, value="=SUM(C29:G29)").font = FONT_TOTAL
    ws.cell(row=33, column=3).number_format = FORMAT_CURRENCY

    # Terminal Value (Gordon Growth)
    ws.cell(row=34, column=2, value="Terminal Value (Perpetuity Growth Method)")
    ws.cell(row=34, column=3, value="=H27/($C$7-$C$8)").font = FONT_CALC
    ws.cell(row=34, column=3).number_format = FORMAT_CURRENCY

    # PV of Terminal Value - ERROR 4: Discount Factor Off-by-One if is_flawed
    ws.cell(row=35, column=2, value="Present Value of Terminal Value")
    cell_pv_tv = ws.cell(row=35, column=3)
    if is_flawed:
        # Error 4: Discounting by (1 + WACC)^6 instead of matching Year 5 exponent (1 + WACC)^5
        cell_pv_tv.value = "=C34/(1+$C$7)^6"
        cell_pv_tv.fill = FILL_ERROR
    else:
        cell_pv_tv.value = "=C34/(1+$C$7)^5"
    cell_pv_tv.font = FONT_CALC
    cell_pv_tv.number_format = FORMAT_CURRENCY

    # Enterprise Value
    ws.cell(row=36, column=2, value="Enterprise Value (EV)").font = FONT_SECTION
    ws.cell(row=36, column=3, value="=C33+C35").font = FONT_TOTAL
    ws.cell(row=36, column=3).number_format = FORMAT_CURRENCY
    ws.cell(row=36, column=3).border = Border(top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"))

    # ERROR 5: Net Debt Bridge Omission if is_flawed
    if is_flawed:
        # Error 5: Subtracting Total Debt directly from EV, completely omitting Cash & Liquid Equivalents
        ws.cell(row=37, column=2, value="(-) Total Debt (Net Debt Cash Omitted)")
        cell_debt = ws.cell(row=37, column=3, value="=$C$11")
        cell_debt.font = FONT_INPUT
        cell_debt.fill = FILL_ERROR
        cell_debt.number_format = FORMAT_CURRENCY

        ws.cell(row=38, column=2, value="Implied Equity Value").font = FONT_TOTAL
        cell_eq = ws.cell(row=38, column=3, value="=C36-C37")
        cell_eq.font = FONT_TOTAL
        cell_eq.number_format = FORMAT_CURRENCY
        cell_eq.border = BORDER_TOTAL

        ws.cell(row=39, column=2, value="Diluted Shares Outstanding (000s)")
        ws.cell(row=39, column=3, value="=$C$12").number_format = "#,##0"

        ws.cell(row=40, column=2, value="Implied Target Share Price").font = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_NAVY)
        cell_px = ws.cell(row=40, column=3, value="=C38/C39")
        cell_px.font = Font(name=FONT_NAME, size=11, bold=True, color="000000")
        cell_px.number_format = FORMAT_PRICE
        cell_px.border = BORDER_TOTAL
    else:
        ws.cell(row=37, column=2, value="(+) Cash and Marketable Securities")
        ws.cell(row=37, column=3, value="=$C$10").number_format = FORMAT_CURRENCY
        ws.cell(row=37, column=3).font = FONT_CALC

        ws.cell(row=38, column=2, value="(-) Total Debt Outstanding")
        ws.cell(row=38, column=3, value="=$C$11").number_format = FORMAT_CURRENCY
        ws.cell(row=38, column=3).font = FONT_CALC

        ws.cell(row=39, column=2, value="Net Debt").font = FONT_SECTION
        ws.cell(row=39, column=3, value="=C38-C37").number_format = FORMAT_CURRENCY
        ws.cell(row=39, column=3).font = FONT_TOTAL

        ws.cell(row=40, column=2, value="Implied Equity Value").font = FONT_TOTAL
        cell_eq = ws.cell(row=40, column=3, value="=C36-C39")
        cell_eq.font = FONT_TOTAL
        cell_eq.number_format = FORMAT_CURRENCY
        cell_eq.border = BORDER_TOTAL

        ws.cell(row=41, column=2, value="Diluted Shares Outstanding (000s)")
        ws.cell(row=41, column=3, value="=$C$12").number_format = "#,##0"

        ws.cell(row=42, column=2, value="Implied Target Share Price").font = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_NAVY)
        cell_px = ws.cell(row=42, column=3, value="=C40/C41")
        cell_px.font = Font(name=FONT_NAME, size=11, bold=True, color="000000")
        cell_px.number_format = FORMAT_PRICE
        cell_px.border = BORDER_TOTAL

    format_grid(ws)
    return wb


def generate_audit_models() -> tuple[Path, Path]:
    """Generate both flawed v1 and audited v2 workbooks on disk."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    path_flawed = OUTPUT_DIR / "target_model_v1_flawed.xlsx"
    path_audited = OUTPUT_DIR / "target_model_v2_audited.xlsx"

    wb_flawed = build_audit_model_workbook(is_flawed=True)
    wb_flawed.save(path_flawed)

    wb_audited = build_audit_model_workbook(is_flawed=False)
    wb_audited.save(path_audited)

    return path_flawed, path_audited


if __name__ == "__main__":
    flawed_p, audited_p = generate_audit_models()
    print(f"Flawed model written to: {flawed_p}")
    print(f"Audited model written to: {audited_p}")
