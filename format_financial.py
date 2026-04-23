"""
Rebuild PnL and Cashflow sheets with professional financial report styling.
"""
import sys, io, glob, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment, numbers
)
from openpyxl.utils import get_column_letter

# ── Find the file ──────────────────────────────────────────────────────────────
BASE = r"C:\Users\Administrator\Desktop\WorkflowGenie-AI-For-Office-Automation\WorkFlow Genie\backend\data\excel_files"
matches = sorted(glob.glob(os.path.join(BASE, "e03233e6*FinancialTest.xlsx")))
if not matches:
    print("File not found"); sys.exit(1)
FILEPATH = matches[0]
print("Formatting:", os.path.basename(FILEPATH))

wb = openpyxl.load_workbook(FILEPATH)

# ── Colour palette ─────────────────────────────────────────────────────────────
C_NAVY      = "1F3864"   # deep navy  – main title bg
C_BLUE      = "2E75B6"   # mid blue   – section headers
C_LBLUE     = "D6E4F0"   # light blue – subtotal rows
C_GREEN     = "1E6B3C"   # dark green – net income / net cash
C_LGREENB   = "E2EFDA"   # pale green – net income bg
C_ORANGE    = "F4B942"   # amber      – warning / negative accent
C_GREY      = "F2F2F2"   # light grey – alternate rows
C_WHITE     = "FFFFFF"
C_BLACK     = "000000"
C_DARK      = "1A1A2E"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, size=11, color=C_BLACK, italic=False):
    return Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)

def border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def thick_border():
    t = Side(style="medium")
    n = Side(style=None)
    return Border(bottom=t)

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

CURRENCY_FMT = '#,##0.00_);[Red](#,##0.00)'
PCT_FMT      = '0.0"%"'

def style_cell(ws, row, col, value=None,
               bg=None, fg=C_BLACK, bold=False, size=11,
               italic=False, number_fmt=None,
               h_align="left", border_style="thin", indent=0):
    cell = ws.cell(row=row, column=col)
    if value is not None:
        cell.value = value
    if bg:
        cell.fill = fill(bg)
    cell.font = Font(name="Calibri", bold=bold, size=size,
                     color=fg, italic=italic)
    if number_fmt:
        cell.number_format = number_fmt
    cell.alignment = Alignment(horizontal=h_align, vertical="center",
                                indent=indent if h_align=="left" else 0)
    if border_style:
        s = Side(style=border_style)
        cell.border = Border(left=s, right=s, top=s, bottom=s)
    return cell

def merge_style(ws, row, c1, c2, value,
                bg=C_NAVY, fg=C_WHITE, bold=True, size=13,
                h_align="center", border_style="thin"):
    ws.merge_cells(start_row=row, start_column=c1,
                   end_row=row, end_column=c2)
    cell = ws.cell(row=row, column=c1)
    cell.value = value
    cell.fill = fill(bg)
    cell.font = Font(name="Calibri", bold=bold, size=size, color=fg)
    cell.alignment = Alignment(horizontal=h_align, vertical="center")
    if border_style:
        s = Side(style=border_style)
        cell.border = Border(left=s, right=s, top=s, bottom=s)
    # fill merged cells too
    for c in range(c1+1, c2+1):
        ws.cell(row=row, column=c).fill = fill(bg)

def section_header(ws, row, label):
    merge_style(ws, row, 1, 3, label,
                bg=C_BLUE, fg=C_WHITE, bold=True, size=10,
                h_align="left")

def data_row(ws, row, label, value, alt=False, indent=1):
    bg = C_GREY if alt else C_WHITE
    style_cell(ws, row, 1, label, bg=bg, bold=False, size=10,
               h_align="left", indent=indent)
    style_cell(ws, row, 2, value, bg=bg, bold=False, size=10,
               number_fmt=CURRENCY_FMT, h_align="right")
    style_cell(ws, row, 3, None, bg=bg, border_style="thin")

def subtotal_row(ws, row, label, value, strong=False):
    bg  = C_LGREENB if strong else C_LBLUE
    fg  = C_GREEN   if strong else C_DARK
    sz  = 11        if strong else 10
    style_cell(ws, row, 1, label, bg=bg, fg=fg, bold=True, size=sz, h_align="left")
    style_cell(ws, row, 2, value, bg=bg, fg=fg, bold=True, size=sz,
               number_fmt=CURRENCY_FMT, h_align="right")
    style_cell(ws, row, 3, None,  bg=bg, border_style="thin")
    # thick bottom border
    for c in range(1, 4):
        cell = ws.cell(row=row, column=c)
        existing = cell.border
        cell.border = Border(
            left=existing.left, right=existing.right,
            top=existing.top, bottom=Side(style="medium")
        )

def blank_row(ws, row):
    for c in range(1, 4):
        ws.cell(row=row, column=c).fill = fill(C_WHITE)
        ws.cell(row=row, column=c).value = None

def set_col_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

def set_row_height(ws, row, h):
    ws.row_dimensions[row].height = h


# ═══════════════════════════════════════════════════════════════════════════════
# P&L SHEET
# ═══════════════════════════════════════════════════════════════════════════════
if "PnL" in wb.sheetnames:
    del wb["PnL"]
ws_pnl = wb.create_sheet("PnL")

revenue        = 623_000
cogs           = 249_200
gross_profit   = revenue - cogs          # 373_800
op_ex          = 368_000
ebitda         = gross_profit - op_ex    # 5_800
depreciation   = 15_000
ebit           = ebitda - depreciation   # -9_200
interest       = 8_000
ebt            = ebit - interest         # -17_200
tax            = round(ebt * 0.30, 2)    # -5_160
net_income     = ebt - tax               # -12_040

r = 1
# Title block
merge_style(ws_pnl, r, 1, 3, "PROFIT & LOSS STATEMENT",
            bg=C_NAVY, fg=C_WHITE, bold=True, size=16); set_row_height(ws_pnl, r, 30); r += 1
merge_style(ws_pnl, r, 1, 3, "Financial Year 2024",
            bg=C_NAVY, fg="BDD7EE", bold=False, size=11, h_align="center"); set_row_height(ws_pnl, r, 18); r += 1
blank_row(ws_pnl, r); set_row_height(ws_pnl, r, 6); r += 1

# REVENUE
section_header(ws_pnl, r, "  REVENUE"); set_row_height(ws_pnl, r, 20); r += 1
data_row(ws_pnl, r, "  Total Revenue", revenue, alt=False); r += 1
subtotal_row(ws_pnl, r, "GROSS REVENUE", revenue); r += 1
blank_row(ws_pnl, r); set_row_height(ws_pnl, r, 6); r += 1

# COGS
section_header(ws_pnl, r, "  COST OF GOODS SOLD"); set_row_height(ws_pnl, r, 20); r += 1
data_row(ws_pnl, r, "  COGS (40% of Revenue)", cogs, alt=True); r += 1
blank_row(ws_pnl, r); set_row_height(ws_pnl, r, 6); r += 1

# GROSS PROFIT
subtotal_row(ws_pnl, r, "GROSS PROFIT", gross_profit); set_row_height(ws_pnl, r, 22); r += 1
data_row(ws_pnl, r, "  Gross Margin", f"{gross_profit/revenue*100:.1f}%", alt=False); r += 1
blank_row(ws_pnl, r); set_row_height(ws_pnl, r, 6); r += 1

# OPERATING EXPENSES
section_header(ws_pnl, r, "  OPERATING EXPENSES"); set_row_height(ws_pnl, r, 20); r += 1
data_row(ws_pnl, r, "  Total Operating Expenses", op_ex, alt=True); r += 1
blank_row(ws_pnl, r); set_row_height(ws_pnl, r, 6); r += 1

# EBITDA
subtotal_row(ws_pnl, r, "EBITDA", ebitda); set_row_height(ws_pnl, r, 22); r += 1
blank_row(ws_pnl, r); set_row_height(ws_pnl, r, 6); r += 1

# BELOW LINE
section_header(ws_pnl, r, "  BELOW-LINE ITEMS"); set_row_height(ws_pnl, r, 20); r += 1
data_row(ws_pnl, r, "  Depreciation", depreciation, alt=False); r += 1
data_row(ws_pnl, r, "  Interest Expense", interest, alt=True); r += 1
blank_row(ws_pnl, r); set_row_height(ws_pnl, r, 6); r += 1

# EBIT / EBT
subtotal_row(ws_pnl, r, "EBIT", ebit); set_row_height(ws_pnl, r, 20); r += 1
data_row(ws_pnl, r, "  Tax (30%)", tax, alt=False); r += 1
blank_row(ws_pnl, r); set_row_height(ws_pnl, r, 6); r += 1

# NET INCOME
subtotal_row(ws_pnl, r, "NET INCOME", net_income, strong=True); set_row_height(ws_pnl, r, 26); r += 1

set_col_widths(ws_pnl, [38, 18, 4])
ws_pnl.sheet_view.showGridLines = False


# ═══════════════════════════════════════════════════════════════════════════════
# CASHFLOW SHEET
# ═══════════════════════════════════════════════════════════════════════════════
if "Cashflow" in wb.sheetnames:
    del wb["Cashflow"]
ws_cf = wb.create_sheet("Cashflow")

op_income    = net_income    # from P&L
dep_addback  = depreciation
total_ops    = op_income + dep_addback
capex        = -25_000
total_invest = capex
financing    = 0
net_cash     = total_ops + total_invest + financing

r = 1
merge_style(ws_cf, r, 1, 3, "CASH FLOW STATEMENT",
            bg=C_NAVY, fg=C_WHITE, bold=True, size=16); set_row_height(ws_cf, r, 30); r += 1
merge_style(ws_cf, r, 1, 3, "Financial Year 2024",
            bg=C_NAVY, fg="BDD7EE", bold=False, size=11, h_align="center"); set_row_height(ws_cf, r, 18); r += 1
blank_row(ws_cf, r); set_row_height(ws_cf, r, 6); r += 1

# OPERATING
section_header(ws_cf, r, "  OPERATING ACTIVITIES"); set_row_height(ws_cf, r, 20); r += 1
data_row(ws_cf, r, "  Net Income",             op_income,   alt=False); r += 1
data_row(ws_cf, r, "  Add: Depreciation",      dep_addback, alt=True);  r += 1
subtotal_row(ws_cf, r, "Total Operating Cash Flow", total_ops); set_row_height(ws_cf, r, 22); r += 1
blank_row(ws_cf, r); set_row_height(ws_cf, r, 6); r += 1

# INVESTING
section_header(ws_cf, r, "  INVESTING ACTIVITIES"); set_row_height(ws_cf, r, 20); r += 1
data_row(ws_cf, r, "  Capital Expenditures", capex, alt=False); r += 1
subtotal_row(ws_cf, r, "Total Investing Cash Flow", total_invest); set_row_height(ws_cf, r, 22); r += 1
blank_row(ws_cf, r); set_row_height(ws_cf, r, 6); r += 1

# FINANCING
section_header(ws_cf, r, "  FINANCING ACTIVITIES"); set_row_height(ws_cf, r, 20); r += 1
data_row(ws_cf, r, "  Financing Activities", financing, alt=False); r += 1
subtotal_row(ws_cf, r, "Total Financing Cash Flow", financing); set_row_height(ws_cf, r, 22); r += 1
blank_row(ws_cf, r); set_row_height(ws_cf, r, 6); r += 1

# NET CASH
subtotal_row(ws_cf, r, "NET CASH FLOW", net_cash, strong=True); set_row_height(ws_cf, r, 26); r += 1

set_col_widths(ws_cf, [38, 18, 4])
ws_cf.sheet_view.showGridLines = False


# ── Reorder sheets: Sheet1 first, then PnL, then Cashflow ─────────────────────
sheet_order = ["Sheet1", "PnL", "Cashflow"]
for i, name in enumerate(sheet_order):
    if name in wb.sheetnames:
        wb.move_sheet(name, offset=wb.sheetnames.index(name) - i)

wb.save(FILEPATH)
print("Done! File saved:", os.path.basename(FILEPATH))
