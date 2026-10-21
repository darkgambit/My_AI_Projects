# -*- coding: utf-8 -*-
"""
Builds the J&Z Trucking Command Center workbook (Excel / Google Sheets compatible).
Run:  python3 build_workbook.py
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

# ---------------------------------------------------------------- palette
NAVY   = "0B2545"
NAVY2  = "13315C"
TEAL   = "0E7C7B"
TEAL_L = "E3F2F1"
AMBER  = "B7791F"
AMBER_L= "FDF6E3"
GREEN  = "1E6F42"
GREEN_L= "E6F4EA"
RED    = "B42318"
RED_L  = "FDECEA"
GREY_B = "F4F6F8"
GREY_L = "EEF1F4"
GREY_M = "D5DBE1"
INK    = "1F2937"
MUTED  = "667085"
WHITE  = "FFFFFF"

PESO   = '"\u20b1"#,##0.00;[Red]-"\u20b1"#,##0.00;"\u2013"'
PESO0  = '"\u20b1"#,##0;[Red]-"\u20b1"#,##0;"\u2013"'
PCT    = '0.0%'
PCT2   = '0.00%'
DATE   = 'dd mmm yyyy'
TONS   = '#,##0.00'

def fill(c):        return PatternFill("solid", fgColor=c)
def fnt(sz=10, bold=False, color=INK, italic=False, name="Calibri", underline=None):
    return Font(name=name, size=sz, bold=bold, color=color, italic=italic, underline=underline)
def al(h="left", v="center", wrap=False, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)
def side(c=GREY_M, s="thin"):  return Side(style=s, color=c)
def box(c=GREY_M):             return Border(left=side(c), right=side(c), top=side(c), bottom=side(c))
def bottom(c=GREY_M):          return Border(bottom=side(c))

def put(ws, ref, value=None, fmt=None, font=None, bg=None, align=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    if fmt:    c.number_format = fmt
    if font:   c.font = font
    if bg:     c.fill = fill(bg)
    if align:  c.alignment = align
    if border: c.border = border
    return c

def band(ws, r1, c1, r2, c2, color, border=None):
    """fill a rectangular block"""
    for r in range(r1, r2 + 1):
        for cidx in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=cidx)
            cell.fill = fill(color)
            if border: cell.border = border

def outline(ws, r1, c1, r2, c2, color=GREY_M):
    for r in range(r1, r2 + 1):
        for cidx in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=cidx)
            left   = side(color) if cidx == c1 else cell.border.left
            right  = side(color) if cidx == c2 else cell.border.right
            top    = side(color) if r == r1 else cell.border.top
            bot    = side(color) if r == r2 else cell.border.bottom
            cell.border = Border(left=left, right=right, top=top, bottom=bot)

def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w

def title_bar(ws, last_col_letter, text, sub=None, bg=NAVY, logo="\U0001F69A"):
    ws.merge_cells(f"A1:{last_col_letter}1")
    put(ws, "A1", f"{logo}  {text}", font=fnt(18, True, WHITE), bg=bg, align=al("left", "center", indent=1))
    ws.row_dimensions[1].height = 40
    if sub is not None:
        ws.merge_cells(f"A2:{last_col_letter}2")
        put(ws, "A2", sub, font=fnt(9.5, False, "D6E4F0", italic=True), bg=NAVY2, align=al("left", "center", indent=1))
        ws.row_dimensions[2].height = 20

def section(ws, row, c1, c2, text, bg=TEAL, fg=WHITE, size=10.5):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    cell = ws.cell(row=row, column=c1, value=text)
    cell.font = fnt(size, True, fg)
    cell.fill = fill(bg)
    cell.alignment = al("left", "center", indent=1)
    ws.row_dimensions[row].height = 22

def table_head(ws, row, c1, labels, bg=GREY_L):
    for i, lab in enumerate(labels):
        cell = ws.cell(row=row, column=c1 + i, value=lab)
        cell.font = fnt(9, True, NAVY)
        cell.fill = fill(bg)
        cell.alignment = al("center", "center", wrap=True)
        cell.border = Border(bottom=side(GREY_M), top=side(GREY_M))
    ws.row_dimensions[row].height = 28

def note(ws, ref, text, size=8.5, color=MUTED, italic=True, span=1):
    if span > 1:
        col = ws[ref].column
        ws.merge_cells(start_row=ws[ref].row, start_column=col, end_row=ws[ref].row, end_column=col + span - 1)
    put(ws, ref, text, font=fnt(size, False, color, italic=italic), align=al("left", "center", wrap=False))

TODAY = dt.date(2026, 10, 21)
import os
TEST_MODE = os.environ.get("JZ_TEST") == "1"     # small build used by the formula test harness
DATA_R1 = 5
DATA_R2 = 60 if TEST_MODE else 1004              # data rows on Trips / Expenses

wb = Workbook()

# ======================================================================
#  SHEET ORDER
# ======================================================================
ws_dash = wb.active;  ws_dash.title = "Dashboard"
ws_trip = wb.create_sheet("Trips")
ws_exp  = wb.create_sheet("Expenses")
ws_pay  = wb.create_sheet("Payroll")
ws_set  = wb.create_sheet("Setup")
ws_read = wb.create_sheet("Read Me")
ws_calc = wb.create_sheet("Calc")
ws_calc.sheet_state = "hidden"

for ws, colour in [(ws_dash, NAVY), (ws_trip, TEAL), (ws_exp, AMBER),
                   (ws_pay, GREEN), (ws_set, NAVY2), (ws_read, MUTED), (ws_calc, GREY_M)]:
    ws.sheet_properties.tabColor = colour
    ws.sheet_view.showGridLines = False

# ======================================================================
#  SETUP  (address map used by every other sheet)
# ======================================================================
S = ws_set
widths(S, {"A": 34, "B": 26, "C": 22, "D": 14, "E": 16, "F": 14, "G": 13,
           "H": 15, "I": 15, "J": 17, "K": 14, "L": 20, "M": 12})
title_bar(S, "M", "SETUP  \u2014  one-time configuration",
          "Fill this once. Everything else (rates, crew, loans, expense types) is driven from here.")

put(S, "G4", "Yellow cells = you type here   \u00b7   grey cells = computed   \u00b7   add rows at the bottom of any table",
    font=fnt(8.5, False, AMBER, italic=True), align=al("left", "center"))

# --- 1. business profile
section(S, 4, 1, 12, "1 \u00b7 BUSINESS PROFILE")
profile = [("Company Name", "J&Z Trucking Services", None),
           ("Owner / Proprietor", "", None),
           ("Contact Number", "", None),
           ("Business Address", "", None),
           ("TIN (for invoicing)", "", None),
           ("Currency", "\u20b1 PHP \u2013 Philippine Peso", None)]
r = 5
for label, val, fmt in profile:
    put(S, f"A{r}", label, font=fnt(10, False, INK), align=al("left", indent=1))
    put(S, f"B{r}", val, font=fnt(10, True, NAVY), bg=AMBER_L, border=box(), align=al("left", indent=1))
    S.merge_cells(f"B{r}:D{r}")
    r += 1

# --- 2. pay rules
section(S, 12, 1, 12, "2 \u00b7 MONEY RULES  (edit the yellow cells \u2014 every trip recalculates)")
rules = [
    (13, "Withholding tax on gross fee", 0.02, PCT, "Deducted from the hauling fee before crew shares."),
    (14, "Driver share of NET fee",      0.10, PCT, "Net fee = gross fee less withholding tax."),
    (15, "Helper share of NET fee",      0.05, PCT, "Same base as the driver share."),
    (16, "Default fuel surcharge",       0.18, PCT, "Added on top of the rate table on every trip (editable per trip)."),
    (17, "Default allowance per trip",   500,  PESO0, "Cash given to the driver per trip (editable per trip)."),
]
for row, label, val, fmt, help_txt in rules:
    put(S, f"A{row}", label, font=fnt(10, False, INK), align=al("left", indent=1))
    put(S, f"B{row}", val, fmt=fmt, font=fnt(11, True, NAVY), bg=AMBER_L, border=box(), align=al("center", indent=0))
    S.merge_cells(f"C{row}:F{row}")
    put(S, f"C{row}", help_txt, font=fnt(8.5, False, MUTED, italic=True), align=al("left", indent=1))

# --- 3. rate table
section(S, 19, 1, 12, "3 \u00b7 RATE TABLE  \u2014  \u20b1 per TON (from J&Z Bulk Rate sheet, before fuel surcharge)")
table_head(S, 20, 1, ["SOURCE (origin)", "DESTINATION", "RATE \u20b1 / TON",
                      "SURCHARGE", "EFFECTIVE \u20b1/TON"], bg=GREY_L)
RATES = [
    ("BATAAN", "TARLAC", 820), ("BATAAN", "BMPP", 950), ("BATAAN", "BULACAN", 870),
    ("BATAAN", "BMEG BATAAN", 104), ("BATAAN", "BATAAN EXTERNAL WHSE", 104),
    ("BATAAN", "ISABELA", 2100), ("BATAAN", "ARMOR", 900),
    ("ORION", "BATAAN", 350), ("ORION", "TARLAC", 730), ("ORION", "BMPP", 900),
    ("ORION", "ROSALES", 790), ("ORION", "BULACAN", 710), ("ORION", "ISABELA", 1000),
    ("ORION", "CAPAS", 550),
    ("MABINI, BATANGAS", "IBAAN EXT.", 320), ("MABINI, BATANGAS", "MASAYA EXT.", 320),
    ("MABINI, BATANGAS", "ARMOR", 420), ("MABINI, BATANGAS", "TARLAC", 1300),
    ("MABINI, BATANGAS", "BMPP", 1450), ("MABINI, BATANGAS", "BATAAN", 1460),
    ("MABINI, BATANGAS", "BULACAN", 920), ("MABINI, BATANGAS", "ISABELA", 1000),
    ("MABINI, BATANGAS", "MONCADA", 975),
]
RATE_R1 = 21
RATE_SLOTS = 40
RATE_R2 = RATE_R1 + RATE_SLOTS - 1
for i in range(RATE_SLOTS):
    row = RATE_R1 + i
    bg = WHITE if i % 2 == 0 else GREY_B
    src, dst, amt = RATES[i] if i < len(RATES) else ("", "", None)
    put(S, f"A{row}", src, font=fnt(10), bg=bg, border=side(GREY_M), align=al("left", indent=1))
    put(S, f"B{row}", dst, font=fnt(10), bg=bg, border=side(GREY_M), align=al("left", indent=1))
    put(S, f"C{row}", amt, fmt=PESO, font=fnt(10), bg=AMBER_L if amt is not None else bg,
        border=side(GREY_M), align=al("right", indent=1))
    put(S, f"D{row}", f"=IF($C{row}=\"\",\"\",Setup!$B$16)", fmt=PCT, font=fnt(9, False, MUTED),
        bg=bg, border=side(GREY_M), align=al("center"))
    put(S, f"E{row}", f"=IF($C{row}=\"\",\"\",ROUND($C{row}*(1+$D{row}),2))", fmt=PESO0,
        font=fnt(9, False, MUTED), bg=bg, border=side(GREY_M), align=al("right", indent=1))

# --- 4. trucks
TRUCK_R1, TRUCK_SLOTS = 63, 10
TRUCK_R2 = TRUCK_R1 + TRUCK_SLOTS - 1
section(S, TRUCK_R1 - 2, 1, 12, "4 \u00b7 TRUCKS / UNITS  (add a row when you buy your next truck)")
table_head(S, TRUCK_R1 - 1, 1, ["UNIT / TRUCK NAME", "PLATE NO.", "CAPACITY (tons)",
                                "STATUS", "ASSIGNED DRIVER", "LENDER"], bg=GREY_L)
TRUCKS = [("TRUCK 1", "NCR 1234", 20, "Active", "R. Dela Cruz", "BDO Leasing")]
for i in range(TRUCK_SLOTS):
    row = TRUCK_R1 + i
    bg = WHITE if i % 2 == 0 else GREY_B
    vals = TRUCKS[i] if i < len(TRUCKS) else ("", "", None, "", "", "")
    for j, v in enumerate(vals):
        cell = S.cell(row=row, column=1 + j, value=v)
        cell.font = fnt(10); cell.fill = fill(bg); cell.border = side(GREY_M)
        cell.alignment = al("left", indent=1)
        if j == 2: cell.number_format = TONS; cell.alignment = al("center")

# --- 5. crew
CREW_R1, CREW_SLOTS = 76, 12
CREW_R2 = CREW_R1 + CREW_SLOTS - 1
section(S, CREW_R1 - 2, 1, 12, "5 \u00b7 CREW  (drivers & helpers)  \u2014  share % can differ per person")
table_head(S, CREW_R1 - 1, 1, ["NAME", "ROLE", "SHARE % (blank = default)", "PHONE"], bg=GREY_L)
CREW = [("R. Dela Cruz", "Driver", 0.10, "0917 000 0001"),
        ("M. Reyes", "Driver", 0.10, "0917 000 0002"),
        ("J. Santos", "Helper", 0.05, "0917 000 0003"),
        ("A. Bautista", "Helper", 0.05, "0917 000 0004")]
for i in range(CREW_SLOTS):
    row = CREW_R1 + i
    bg = WHITE if i % 2 == 0 else GREY_B
    vals = CREW[i] if i < len(CREW) else ("", "", None, "")
    put(S, f"A{row}", vals[0], font=fnt(10), bg=bg, border=side(GREY_M), align=al("left", indent=1))
    put(S, f"B{row}", vals[1], font=fnt(10), bg=bg, border=side(GREY_M), align=al("center"))
    put(S, f"C{row}", vals[2], fmt=PCT, font=fnt(10), bg=bg, border=side(GREY_M), align=al("center"))
    put(S, f"D{row}", vals[3], font=fnt(9), bg=bg, border=side(GREY_M), align=al("left", indent=1))

# --- 6. expense types -> cost group
EXP_R1, EXP_SLOTS = 99, 30
EXP_R2 = EXP_R1 + EXP_SLOTS - 1
section(S, EXP_R1 - 2, 1, 12,
        "6 \u00b7 EXPENSE TYPES  \u2192  COST GROUP  (used by the Dashboard P&L \u2014 keep the group names as they are)")
table_head(S, EXP_R1 - 1, 1, ["EXPENSE TYPE  (dropdown in Expenses)", "COST GROUP",
                              "WHAT BELONGS HERE"], bg=GREY_L)
EXP_TYPES = [
    ("Fuel \u2013 Trip",            "Vehicle",   "Fuel bought per trip (or log fuel in the Trips sheet instead \u2014 not both)"),
    ("Fuel \u2013 Bulk Purchase",   "Vehicle",   "Liters bought in bulk / delivered to your tank"),
    ("Toll & Parking",             "Vehicle",   "RFID load, expressway tolls, parking"),
    ("Repairs & Maintenance",      "Vehicle",   "PMS, mechanic, labour, engine repair"),
    ("Tires & Batteries",          "Vehicle",   "New tires, retread, battery"),
    ("Spare Parts & Lubricants",   "Vehicle",   "Filters, oil, belts, bulbs"),
    ("Insurance",                  "Vehicle",   "Comprehensive / CTPL premium"),
    ("LTO Registration",           "Vehicle",   "Registration, emission test, MVIC"),
    ("Truck Washing",              "Vehicle",   "Carwash, greasing"),
    ("Truck Accessories",          "Vehicle",   "Tarpaulin, mats, tools, weighing fees"),
    ("Loan Payment",               "Financing", "Monthly amortisation to any lender"),
    ("Cash Advance",               "Transfer",  "Cash given to crew \u2014 settled in Payroll (NOT an extra cost)"),
    ("Payroll Payout",             "Transfer",  "Cash released to crew \u2014 settled in Payroll (NOT an extra cost)"),
    ("Reimbursement Payout",       "Transfer",  "Paying back the driver's receipts already logged in Trips"),
    ("Office / Staff Salary",      "Overhead",  "Dispatcher, bookkeeper, office staff"),
    ("Rent",                       "Overhead",  "Yard / office / parking rent"),
    ("Utilities",                  "Overhead",  "Electricity, water, internet, phone"),
    ("Office Supplies",            "Overhead",  "Printer, forms, bond paper"),
    ("Professional Fees",          "Overhead",  "Accountant, lawyer, broker's fee"),
    ("Bank Charges",               "Overhead",  "Transfer fees, interest on overdraft"),
    ("Permits & Regulatory",       "Overhead",  "Mayor's permit, BIR, LTFRB, insurance broker"),
    ("Miscellaneous",              "Overhead",  "Small costs that fit nowhere else"),
    ("Owner's Draw",               "Other",     "Owner's personal withdrawal \u2014 excluded from P&L"),
    ("Equipment / Asset Purchase", "Other",     "Capital outlay \u2014 excluded from P&L"),
    ("", "", ""), ("", "", ""), ("", "", ""), ("", "", ""), ("", "", ""), ("", "", ""),
]
for i in range(EXP_SLOTS):
    row = EXP_R1 + i
    bg = WHITE if i % 2 == 0 else GREY_B
    vals = EXP_TYPES[i] if i < len(EXP_TYPES) else ("", "", "")
    put(S, f"A{row}", vals[0], font=fnt(10), bg=bg, border=side(GREY_M), align=al("left", indent=1))
    put(S, f"B{row}", vals[1], font=fnt(10, True, TEAL), bg=bg, border=side(GREY_M), align=al("center"))
    S.merge_cells(f"C{row}:F{row}")
    put(S, f"C{row}", vals[2], font=fnt(8.5, False, MUTED, italic=True), bg=bg, border=side(GREY_M),
        align=al("left", indent=1))

# --- 7. loans
LOAN_R1, LOAN_SLOTS = 133, 6
LOAN_R2 = LOAN_R1 + LOAN_SLOTS - 1
section(S, LOAN_R1 - 2, 1, 12, "7 \u00b7 LOANS / FINANCING  (every lender for every truck \u2014 self-updating balance)")
table_head(S, LOAN_R1 - 1, 1, ["LENDER", "FOR TRUCK", "PRINCIPAL BORROWED", "ANNUAL RATE",
                               "MONTHLY AMORT.", "TERM (months)", "START DATE"], bg=GREY_L)
# computed columns to the right of the input table
for j, lab in enumerate(["PAYMENTS MADE", "TOTAL PAID", "MONTHS LEFT", "REMAINING BALANCE",
                         "REMAINING OBLIGATION", "PROGRESS"]):
    cell = S.cell(row=LOAN_R1 - 1, column=8 + j, value=lab)
    cell.font = fnt(9, True, NAVY); cell.fill = fill(GREY_L)
    cell.alignment = al("center", "center", wrap=True)
    cell.border = Border(bottom=side(GREY_M), top=side(GREY_M))
S.row_dimensions[LOAN_R1 - 1].height = 28
LOANS = [("BDO Leasing", "TRUCK 1", 1200000, 0.12, 32000, 48, dt.date(2025, 6, 1)),
         ("Security Bank", "TRUCK 1", 600000, 0.14, 18500, 36, dt.date(2025, 9, 1)),
         ("Private Lender", "TRUCK 1", 300000, 0.00, 12500, 24, dt.date(2026, 1, 1))]
for i in range(LOAN_SLOTS):
    row = LOAN_R1 + i
    bg = WHITE if i % 2 == 0 else GREY_B
    vals = LOANS[i] if i < len(LOANS) else ("", "", None, None, None, None, None)
    put(S, f"A{row}", vals[0], font=fnt(10, True, NAVY), bg=bg, border=side(GREY_M), align=al("left", indent=1))
    put(S, f"B{row}", vals[1], font=fnt(10), bg=bg, border=side(GREY_M), align=al("center"))
    put(S, f"C{row}", vals[2], fmt=PESO0, font=fnt(10), bg=bg, border=side(GREY_M), align=al("right", indent=1))
    put(S, f"D{row}", vals[3], fmt=PCT, font=fnt(10), bg=bg, border=side(GREY_M), align=al("center"))
    put(S, f"E{row}", vals[4], fmt=PESO0, font=fnt(10), bg=bg, border=side(GREY_M), align=al("right", indent=1))
    put(S, f"F{row}", vals[5], fmt='0', font=fnt(10), bg=bg, border=side(GREY_M), align=al("center"))
    put(S, f"G{row}", vals[6], fmt=DATE, font=fnt(10), bg=bg, border=side(GREY_M), align=al("center"))
    put(S, f"H{row}", f'=IF($A{row}="","",COUNTIFS(Expenses!$B$5:$B$1004,"Loan Payment",Expenses!$E$5:$E$1004,$A{row}))',
        fmt='0', font=fnt(10, False, MUTED), bg=GREY_L, border=side(GREY_M), align=al("center"))
    put(S, f"I{row}", f'=IF($A{row}="","",SUMIFS(Expenses!$F$5:$F$1004,Expenses!$B$5:$B$1004,"Loan Payment",Expenses!$E$5:$E$1004,$A{row}))',
        fmt=PESO0, font=fnt(10, False, MUTED), bg=GREY_L, border=side(GREY_M), align=al("right", indent=1))
    put(S, f"J{row}", f'=IF($A{row}="","",MAX(0,$F{row}-$H{row}))', fmt='0',
        font=fnt(10, False, MUTED), bg=GREY_L, border=side(GREY_M), align=al("center"))
    put(S, f"K{row}", f'=IF($A{row}="","",MAX(0,$C{row}-$I{row}))', fmt=PESO0,
        font=fnt(10, True, NAVY), bg=GREY_L, border=side(GREY_M), align=al("right", indent=1))
    put(S, f"L{row}", f'=IF($A{row}="","",MAX(0,$J{row}*$E{row}))', fmt=PESO0,
        font=fnt(10, True, RED), bg=GREY_L, border=side(GREY_M), align=al("right", indent=1))
    put(S, f"M{row}", f'=IF($A{row}="","",IF($C{row}=0,0,MIN(1,$I{row}/$C{row})))', fmt=PCT,
        font=fnt(10, False, MUTED), bg=GREY_L, border=side(GREY_M), align=al("center"))

note(S, f"A{LOAN_R2+1}",
     "Remaining Balance = principal less everything already paid (cash view). Remaining Obligation = months left \u00d7 amortisation "
     "= what you still have to produce in cash.", size=8.5, span=12)

# --- 8. destination list (dashboard ranking) + payment methods
DEST_R1, DEST_SLOTS = 145, 20
DEST_R2 = DEST_R1 + DEST_SLOTS - 1
section(S, DEST_R1 - 2, 1, 12, "8 \u00b7 DESTINATION LIST  (feeds the dashboard ranking \u2014 add a name when you open a new lane)")
table_head(S, DEST_R1 - 1, 1, ["DESTINATION (as typed in Trips)", "SOURCE (as typed in Trips)"], bg=GREY_L)
DESTS = sorted({d for _, d, _ in RATES})
for i in range(DEST_SLOTS):
    row = DEST_R1 + i
    bg = WHITE if i % 2 == 0 else GREY_B
    put(S, f"A{row}", DESTS[i] if i < len(DESTS) else "", font=fnt(9.5), bg=bg,
        border=side(GREY_M), align=al("left", indent=1))

section(S, DEST_R2 + 2, 1, 12, "9 \u00b7 OTHER LISTS")
table_head(S, DEST_R2 + 3, 1, ["PAYMENT METHOD", "STATUS"], bg=GREY_L)
METHODS = ["Cash", "GCash", "Bank Transfer", "Check"]
for i, m in enumerate(METHODS):
    row = DEST_R2 + 4 + i
    put(S, f"A{row}", m, font=fnt(10), border=side(GREY_M), align=al("left", indent=1))
for i in range(4):
    row = DEST_R2 + 4 + i
    put(S, f"B{row}", ["Active", "Active", "Active", "Active"][i], font=fnt(9, False, MUTED),
        border=side(GREY_M), align=al("center"))
METHOD_R2 = DEST_R2 + 7

# ======================================================================
#  workbook-level helpers / constants
# ======================================================================
M_R1, M_R2 = 11, 46                    # Calc month table rows (36 months)
MONTHS = [dt.date(2026, 1, 1)]
for i in range(1, 36):
    y, m = MONTHS[0].year + (i // 12), (i % 12) + 1
    MONTHS.append(dt.date(y, m, 1))

def q(sheet):        # sheet name for formulas
    return f"'{sheet}'" if " " in sheet else sheet

T_DATE, T_TRUCK, T_DRV, T_HLP = "$A", "$C", "$D", "$E"
T_SRC, T_DST, T_TON, T_RATE = "$F", "$G", "$H", "$I"
T_GROSS, T_TAX, T_NET = "$K", "$L", "$M"
T_DPAY, T_HPAY, T_ALLOW, T_REIM = "$N", "$O", "$P", "$Q"
T_FUEL, T_TOLL, T_OTHER = "$S", "$T", "$U"
T_COST, T_PROFIT, T_MONTH = "$V", "$W", "$Z"

def rng(sheet, col, r1=DATA_R1, r2=DATA_R2):
    return f"{q(sheet)}!{col}{r1}:{col}{r2}"

def truck_branch(sheet, col, sum_col, month_col="$Z", truck_col="$C", extra=""):
    """SUMIFS that respects Dashboard month + truck filters (blank truck rows included on 'All')."""
    base = rng(sheet, sum_col)
    crit = f"{rng(sheet, month_col)},Calc!$B$3"
    if extra:
        crit += "," + extra
    return (f'=IF(Calc!$B$4="All",SUMIFS({base},{crit}),'
            f'SUMIFS({base},{crit},{rng(sheet, truck_col)},Calc!$B$4))')

def exp_branch(sum_col="$F", extra=""):
    return truck_branch("Expenses", None, sum_col, month_col="$J", truck_col="$D", extra=extra)

def exp_group(group, sum_col="$F"):
    return (f'=IF(Calc!$B$4="All",SUMIFS({rng("Expenses", sum_col)},{rng("Expenses","$J")},Calc!$B$3,'
            f'{rng("Expenses","$I")},"{group}"),'
            f'SUMIFS({rng("Expenses", sum_col)},{rng("Expenses","$J")},Calc!$B$3,'
            f'{rng("Expenses","$I")},"{group}",{rng("Expenses","$D")},Calc!$B$4))')

# ======================================================================
#  CALC  (hidden engine: filters, monthly P&L, destination / crew / group roll-ups)
# ======================================================================
C_ = ws_calc
TR_Z, TR_C, TR_D, TR_E = rng("Trips", "$Z"), rng("Trips", "$C"), rng("Trips", "$D"), rng("Trips", "$E")
EX_J, EX_I, EX_D, EX_B, EX_F = rng("Expenses", "$J"), rng("Expenses", "$I"), rng("Expenses", "$D"), rng("Expenses", "$B"), rng("Expenses", "$F")

M_HDR = 21                                  # section title row; row M_HDR+1 = column headers
M_R1, M_R2 = M_HDR + 2, M_HDR + 25          # 24 monthly rows: Jan 2026 - Dec 2027
D_HDR = M_R2 + 3
D_R1, D_R2 = D_HDR + 2, D_HDR + 21
W_HDR = D_R2 + 3
W_R1, W_R2 = W_HDR + 2, W_HDR + 13          # crew count must match Setup crew slots
G_HDR = W_R2 + 3
G_R1, G_R2 = G_HDR + 2, G_HDR + 6
L_TOT = G_R2 + 2

put(C_, "A1", "CALC \u2014 engine sheet (do not edit, do not delete)", font=fnt(12, True, RED))
put(C_, "A2", "Every dashboard number is computed from here \u2014 the two filters are typed on the Dashboard.",
    font=fnt(9, False, MUTED, italic=True))

put(C_, "A3", "Selected month", font=fnt(9, False, MUTED))
put(C_, "B3", "=Dashboard!$C$4", fmt=DATE, font=fnt(9, True, NAVY))
put(C_, "A4", "Selected truck", font=fnt(9, False, MUTED))
put(C_, "B4", "=Dashboard!$F$4", font=fnt(9, True, NAVY))

# ---------------- key figures: B7..B16
KPI_ROWS = [
    (7,  "Revenue \u2014 selected month",           None,                                           PESO),
    (8,  "Net profit \u2014 selected month",        None,                                           PESO),
    (9,  "Gross margin \u2014 selected month",      '=IF($B$7=0,"",$B$8/$B$7)',                     PCT),
    (10, "Trips \u2014 selected month",             None,                                           '0'),
    (11, "Tons hauled \u2014 selected month",       None,                                           TONS),
    (12, "Revenue \u2014 all months",               None,                                           PESO),
    (13, "Net profit \u2014 all months",            None,                                           PESO),
    (14, "Direct cost \u2014 selected month",       None,                                           PESO),
    (15, "Crew cost (share + allowance) \u2014 selected month", None,                                PESO),
    (16, "Expenses booked \u2014 selected month",   None,                                           PESO),
]
put(C_, "A6", "KEY FIGURES (read by the dashboard)", font=fnt(10, True, WHITE), bg=NAVY)
put(C_, "B6", "", bg=NAVY)
put(C_, "C6", "KPI", font=fnt(9, True, WHITE), bg=NAVY)
for col in "DEF":
    put(C_, f"{col}6", "", bg=NAVY)

C_["B7"]  = (f'=IF($B$4="All",SUMIFS({rng("Trips","$K")},{TR_Z},$B$3),'
             f'SUMIFS({rng("Trips","$K")},{TR_Z},$B$3,{TR_C},$B$4))')
C_["B8"]  = (f'=IF($B$4="All",SUMIFS({rng("Trips","$V")},{TR_Z},$B$3),'
             f'SUMIFS({rng("Trips","$V")},{TR_Z},$B$3,{TR_C},$B$4))')
C_["B10"] = (f'=IF($B$4="All",COUNTIFS({TR_Z},$B$3),COUNTIFS({TR_Z},$B$3,{TR_C},$B$4))')
C_["B11"] = (f'=IF($B$4="All",SUMIFS({rng("Trips","$H")},{TR_Z},$B$3),'
             f'SUMIFS({rng("Trips","$H")},{TR_Z},$B$3,{TR_C},$B$4))')
C_["B12"] = f'=IF($B$4="All",SUM({rng("Trips","$K")}),SUMIFS({rng("Trips","$K")},{TR_C},$B$4))'
C_["B13"] = f'=IF($B$4="All",SUM({rng("Trips","$V")}),SUMIFS({rng("Trips","$V")},{TR_C},$B$4))'
C_["B14"] = (f'=IF($B$4="All",SUMIFS({rng("Trips","$U")},{TR_Z},$B$3),'
             f'SUMIFS({rng("Trips","$U")},{TR_Z},$B$3,{TR_C},$B$4))')
C_["B15"] = (f'=IF($B$4="All",SUMIFS({rng("Trips","$N")},{TR_Z},$B$3)+SUMIFS({rng("Trips","$O")},{TR_Z},$B$3)'
             f'+SUMIFS({rng("Trips","$P")},{TR_Z},$B$3),'
             f'SUMIFS({rng("Trips","$N")},{TR_Z},$B$3,{TR_C},$B$4)+SUMIFS({rng("Trips","$O")},{TR_Z},$B$3,{TR_C},$B$4)'
             f'+SUMIFS({rng("Trips","$P")},{TR_Z},$B$3,{TR_C},$B$4))')
C_["B16"] = (f'=IF($B$4="All",SUMIFS({EX_F},{EX_J},$B$3),'
             f'SUMIFS({EX_F},{EX_J},$B$3,{EX_D},$B$4))')
for rr, lab, formula, fmt in KPI_ROWS:
    if formula:
        C_[f"B{rr}"] = formula
    put(C_, f"B{rr}", C_[f"B{rr}"].value, fmt=fmt, font=fnt(10, True, NAVY), align=al("right"))
    put(C_, f"C{rr}", lab, font=fnt(8.5, False, MUTED))

# ---------------- monthly table (all trucks) rows M_HDR..M_R2
put(C_, f"A{M_HDR}", "MONTHLY LEDGER (all trucks) \u2192 feeds the trend charts", font=fnt(10, True, WHITE), bg=NAVY)
heads = ["MONTH KEY", "LABEL", "TRIPS", "TONS", "GROSS FEE", "TAX", "NET FEE",
         "CREW PAY", "ALLOWANCE", "REIMBURSED", "DIRECT COST", "TOTAL COST",
         "NET PROFIT", "MARGIN", "AVG PROFIT / TRIP"]
for j, h in enumerate(heads, 1):
    put(C_, f"{get_column_letter(j)}{M_HDR+1}", h, font=fnt(8, True, NAVY), bg=GREY_L, align=al("center", wrap=True))
for i in range(M_R2 - M_R1 + 1):
    r = M_R1 + i
    mon = MONTHS[i]
    lab = mon.strftime("%b %Y")
    put(C_, f"A{r}", mon, fmt="yyyy-mm-dd", font=fnt(9))
    put(C_, f"B{r}", lab, font=fnt(9))
    put(C_, f"C{r}", f'=COUNTIFS({TR_Z},$A{r})', fmt='0', font=fnt(9))
    put(C_, f"D{r}", f'=SUMIFS({rng("Trips","$H")},{TR_Z},$A{r})', fmt=TONS, font=fnt(9))
    put(C_, f"E{r}", f'=SUMIFS({rng("Trips","$K")},{TR_Z},$A{r})', fmt=PESO, font=fnt(9))
    put(C_, f"F{r}", f'=SUMIFS({rng("Trips","$L")},{TR_Z},$A{r})', fmt=PESO, font=fnt(9))
    put(C_, f"G{r}", f'=SUMIFS({rng("Trips","$M")},{TR_Z},$A{r})', fmt=PESO, font=fnt(9))
    put(C_, f"H{r}", f'=SUMIFS({rng("Trips","$N")},{TR_Z},$A{r})+SUMIFS({rng("Trips","$O")},{TR_Z},$A{r})', fmt=PESO, font=fnt(9))
    put(C_, f"I{r}", f'=SUMIFS({rng("Trips","$P")},{TR_Z},$A{r})', fmt=PESO, font=fnt(9))
    put(C_, f"J{r}", f'=SUMIFS({rng("Trips","$Q")},{TR_Z},$A{r})', fmt=PESO, font=fnt(9))
    put(C_, f"K{r}", f'=SUMIFS({rng("Trips","$U")},{TR_Z},$A{r})', fmt=PESO, font=fnt(9))
    put(C_, f"L{r}", f'=H{r}+I{r}+J{r}+K{r}', fmt=PESO, font=fnt(9))
    put(C_, f"M{r}", f'=E{r}-L{r}', fmt=PESO, font=fnt(9, True))
    put(C_, f"N{r}", f'=IF(E{r}=0,"",M{r}/E{r})', fmt=PCT, font=fnt(9))
    put(C_, f"O{r}", f'=IF(C{r}=0,"",M{r}/C{r})', fmt=PESO0, font=fnt(9))

# ---------------- destination roll-up
put(C_, f"A{D_HDR}", "RANKING BY DESTINATION", font=fnt(10, True, WHITE), bg=NAVY)
for j, h in enumerate(["DESTINATION", "TRIPS", "TONS", "GROSS FEE", "TOTAL COST", "NET PROFIT", "MARGIN"], 1):
    put(C_, f"{get_column_letter(j)}{D_HDR+1}", h, font=fnt(8, True, NAVY), bg=GREY_L, align=al("center", wrap=True))
for i in range(D_R2 - D_R1 + 1):
    r = D_R1 + i
    srow = DEST_R1 + i
    put(C_, f"A{r}", f'=IF(Setup!$A{srow}="","",Setup!$A{srow})', font=fnt(9))
    put(C_, f"B{r}", f'=IF($A{r}="","",IF($B$4="All",COUNTIFS({rng("Trips","$G")},$A{r},{TR_Z},$B$3),'
                      f'COUNTIFS({rng("Trips","$G")},$A{r},{TR_Z},$B$3,{TR_C},$B$4)))', fmt='0', font=fnt(9))
    for cl, col in [(3, "$H"), (4, "$K"), (5, "$U"), (6, "$V")]:
        L = get_column_letter(cl)
        put(C_, f"{L}{r}", f'=IF($A{r}="","",IF($B$4="All",SUMIFS({rng("Trips",col)},{rng("Trips","$G")},$A{r},{TR_Z},$B$3),'
                           f'SUMIFS({rng("Trips",col)},{rng("Trips","$G")},$A{r},{TR_Z},$B$3,{TR_C},$B$4)))',
            fmt=TONS if cl == 3 else PESO, font=fnt(9))
    put(C_, f"G{r}", f'=IF($A{r}="","",IF($D{r}=0,"",$F{r}/$D{r}))', fmt=PCT, font=fnt(9))
    put(C_, f"H{r}", f'=IF($A{r}="","",COUNTIF($D${D_R1}:$D${D_R2},">"&$D{r})+COUNTIF($D${D_R1}:$D{r},$D{r}))',
        fmt='0', font=fnt(8, False, MUTED))

# sorted (ranked) destination view for the dashboard
put(C_, f"J{D_HDR}", "TOP DESTINATIONS (sorted by gross fee)", font=fnt(8, True, WHITE), bg=NAVY)
for j, h in enumerate(["DESTINATION", "TRIPS", "TONS", "GROSS FEE", "TRIP PROFIT", "MARGIN"], 10):
    put(C_, f"{get_column_letter(j)}{D_HDR+1}", h, font=fnt(8, True, NAVY), bg=GREY_L, align=al("center"))
for i in range(D_R2 - D_R1 + 1):
    r = D_R1 + i
    k = i + 1
    idx = f'MATCH({k},$H${D_R1}:$H${D_R2},0)'
    put(C_, f"J{r}", f'=IFERROR(INDEX($A${D_R1}:$A${D_R2},{idx}),"")', font=fnt(9))
    put(C_, f"K{r}", f'=IFERROR(INDEX($B${D_R1}:$B${D_R2},{idx}),"")', fmt='0', font=fnt(9))
    put(C_, f"L{r}", f'=IFERROR(INDEX($C${D_R1}:$C${D_R2},{idx}),"")', fmt=TONS, font=fnt(9))
    put(C_, f"M{r}", f'=IFERROR(INDEX($D${D_R1}:$D${D_R2},{idx}),"")', fmt=PESO, font=fnt(9))
    put(C_, f"N{r}", f'=IFERROR(INDEX($F${D_R1}:$F${D_R2},{idx}),"")', fmt=PESO, font=fnt(9))
    put(C_, f"O{r}", f'=IF($J{r}="","",IF($M{r}=0,"",$N{r}/$M{r}))', fmt=PCT, font=fnt(9))

# ---------------- crew roll-up (by name, selected month)
put(C_, f"A{W_HDR}", "CREW ROLL-UP (selected month)", font=fnt(10, True, WHITE), bg=NAVY)
for j, h in enumerate(["NAME", "ROLE", "TRIPS", "GROSS FEE HANDLED", "SHARE EARNED",
                       "ALLOWANCE", "REIMBURSED", "ADVANCE", "RELEASED", "BALANCE"], 1):
    put(C_, f"{get_column_letter(j)}{W_HDR+1}", h, font=fnt(8, True, NAVY), bg=GREY_L, align=al("center", wrap=True))
for i in range(W_R2 - W_R1 + 1):
    r = W_R1 + i
    srow = CREW_R1 + i
    put(C_, f"A{r}", f'=IF(Setup!$A{srow}="","",Setup!$A{srow})', font=fnt(9))
    put(C_, f"B{r}", f'=IF($A{r}="","",Setup!$B{srow})', font=fnt(9))
    put(C_, f"C{r}", f'=IF($A{r}="","",COUNTIFS({TR_Z},$B$3,{TR_D},$A{r})+COUNTIFS({TR_Z},$B$3,{TR_E},$A{r}))', fmt='0', font=fnt(9))
    put(C_, f"D{r}", f'=IF($A{r}="","",SUMIFS({rng("Trips","$K")},{TR_Z},$B$3,{TR_D},$A{r})+SUMIFS({rng("Trips","$K")},{TR_Z},$B$3,{TR_E},$A{r}))', fmt=PESO, font=fnt(9))
    put(C_, f"E{r}", f'=IF($A{r}="","",SUMIFS({rng("Trips","$N")},{TR_Z},$B$3,{TR_D},$A{r})+SUMIFS({rng("Trips","$O")},{TR_Z},$B$3,{TR_E},$A{r}))', fmt=PESO, font=fnt(9))
    put(C_, f"F{r}", f'=IF($A{r}="","",SUMIFS({rng("Trips","$P")},{TR_Z},$B$3,{TR_D},$A{r}))', fmt=PESO, font=fnt(9))
    put(C_, f"G{r}", f'=IF($A{r}="","",SUMIFS({rng("Trips","$Q")},{TR_Z},$B$3,{TR_D},$A{r}))', fmt=PESO, font=fnt(9))
    put(C_, f"H{r}", f'=IF($A{r}="","",SUMIFS({EX_F},{EX_B},"Cash Advance",{EX_J},$B$3,{rng("Expenses","$E")},$A{r}))', fmt=PESO, font=fnt(9))
    put(C_, f"I{r}", f'=IF($A{r}="","",SUMIFS({EX_F},{EX_B},"Payroll Payout",{EX_J},$B$3,{rng("Expenses","$E")},$A{r}))', fmt=PESO, font=fnt(9))
    put(C_, f"J{r}", f'=IF($A{r}="","",$E{r}+$F{r}+$G{r}-$H{r}-$I{r})', fmt=PESO, font=fnt(9, True))

# ---------------- expense groups
put(C_, f"A{G_HDR}", "EXPENSES BY COST GROUP (selected month)", font=fnt(10, True, WHITE), bg=NAVY)
for j, h in enumerate(["GROUP", "AMOUNT"], 1):
    put(C_, f"{get_column_letter(j)}{G_HDR+1}", h, font=fnt(8, True, NAVY), bg=GREY_L, align=al("center"))
GROUPS = ["Vehicle", "Financing", "Transfer", "Overhead", "Other"]
for i, g in enumerate(GROUPS):
    r = G_R1 + i
    put(C_, f"A{r}", g, font=fnt(9))
    put(C_, f"B{r}", f'=IF($B$4="All",SUMIFS({EX_F},{EX_J},$B$3,{EX_I},$A{r}),'
                     f'SUMIFS({EX_F},{EX_J},$B$3,{EX_I},$A{r},{EX_D},$B$4))', fmt=PESO, font=fnt(9))
put(C_, f"A{L_TOT}", "TOTAL BOOKED EXPENSES", font=fnt(9, True))
put(C_, f"B{L_TOT}", f"=SUM(B{G_R1}:B{G_R2})", fmt=PESO, font=fnt(9, True))
put(C_, f"A{L_TOT+1}", "Loan totals", font=fnt(9, True, MUTED))
put(C_, f"B{L_TOT+1}", f"=SUM(Setup!$I${LOAN_R1}:$I${LOAN_R2})", fmt=PESO, font=fnt(9, True))
put(C_, f"C{L_TOT+1}", "monthly amortisation", font=fnt(8, False, MUTED))
put(C_, f"A{L_TOT+2}", "Loan balance", font=fnt(9, True, MUTED))
put(C_, f"B{L_TOT+2}", f"=SUM(Setup!$K${LOAN_R1}:$K${LOAN_R2})", fmt=PESO, font=fnt(9, True))
put(C_, f"C{L_TOT+2}", "remaining principal", font=fnt(8, False, MUTED))
put(C_, f"A{L_TOT+3}", "Loan obligation", font=fnt(9, True, MUTED))
put(C_, f"B{L_TOT+3}", f"=SUM(Setup!$L${LOAN_R1}:$L${LOAN_R2})", fmt=PESO, font=fnt(9, True))
put(C_, f"C{L_TOT+3}", "cash still to be produced", font=fnt(8, False, MUTED))

# ======================================================================
#  setup extras: hidden rate key + source list used by dropdowns
# ======================================================================
for i in range(RATE_SLOTS):
    r = RATE_R1 + i
    put(S, f"P{r}", f'=IF($C{r}="","",UPPER(TRIM($A{r}))&"|"&UPPER(TRIM($B{r})))',
        font=fnt(8, False, MUTED))
for i in range(RATE_SLOTS):
    r = RATE_R1 + i
    put(S, f"Q{r}", f'=IF($C{r}="","","|"&UPPER(TRIM($B{r})))', font=fnt(8, False, MUTED))
S.column_dimensions["P"].hidden = True
S.column_dimensions["Q"].hidden = True
put(S, f"P{RATE_R1-1}", "KEY (hidden)", font=fnt(8, False, MUTED))

put(S, f"C{DEST_R2+3}", "SOURCE (dropdown on Trips)", font=fnt(9, True, NAVY), bg=GREY_L,
    align=al("center", wrap=True), border=Border(bottom=side(GREY_M), top=side(GREY_M)))
SOURCES = ["BATAAN", "ORION", "MABINI, BATANGAS", "OTHER"]
for i, src in enumerate(SOURCES):
    put(S, f"C{DEST_R2+4+i}", src, font=fnt(10), border=side(GREY_M), align=al("left", indent=1))
SRC_R2 = DEST_R2 + 3 + len(SOURCES)

LST_TRUCK = f"=Setup!$A${TRUCK_R1}:$A${TRUCK_R2}"
LST_CREW  = f"=Setup!$A${CREW_R1}:$A${CREW_R2}"
LST_SRC   = f"=Setup!$C${DEST_R2+4}:$C${SRC_R2}"
LST_DEST  = f"=Setup!$A${DEST_R1}:$A${DEST_R2}"
LST_TYPE  = f"=Setup!$A${EXP_R1}:$A${EXP_R2}"
LST_METH  = f"=Setup!$A${DEST_R2+4}:$A${DEST_R2+7}"
LST_LOAN  = f"=Setup!$A${LOAN_R1}:$A${LOAN_R2}"

SHARE_DRV = "Setup!$B$14"
SHARE_HLP = "Setup!$B$15"
SURCH     = "Setup!$B$16"
ALLOW     = "Setup!$B$17"

# ======================================================================
#  TRIPS  (input #1)
# ======================================================================
T = ws_trip
TRIP_COLS = [
    # letter, header, width, kind(input/calc/hidden), fmt, note
    ("A", "DATE",               11, "in",   DATE,    "date of trip"),
    ("B", "INVOICE / DR NO.",   16, "in",   None,    "your invoice or delivery receipt #"),
    ("C", "TRUCK",              12, "in",   None,    "dropdown"),
    ("D", "DRIVER",             16, "in",   None,    "dropdown from Setup crew"),
    ("E", "HELPER",             16, "in",   None,    "dropdown  (\u2013 if none)"),
    ("F", "SOURCE",             18, "in",   None,    "origin, dropdown"),
    ("G", "DESTINATION",        23, "in",   None,    "dropdown, matches the rate table"),
    ("H", "TONS",                8, "in",   TONS,    "weight hauled"),
    ("I", "RATE /\nTON",        11, "calc", PESO,    "auto from rate table"),
    ("J", "FUEL\nSURCH. %",      9, "in",   PCT,     "blank = 18%"),
    ("K", "GROSS\nFEE",         12, "calc", PESO,    "tons x rate x (1+surcharge)"),
    ("L", "2% TAX",             10, "calc", PESO,    "withheld"),
    ("M", "NET FEE",            12, "calc", PESO,    "gross less tax"),
    ("N", "DRIVER\nPAY",        11, "calc", PESO,    "share of net fee"),
    ("O", "HELPER\nPAY",        11, "calc", PESO,    "share of net fee"),
    ("P", "ALLOWANCE",          11, "in",   PESO,    "blank = default"),
    ("Q", "REIMBURSABLE",       13, "in",   PESO,    "driver receipts to pay back"),
    ("R", "FUEL COST",          11, "in",   PESO,    "diesel for this trip"),
    ("S", "TOLL +\nPARKING",    11, "in",   PESO,    ""),
    ("T", "OTHER TRIP\nCOST",   12, "in",   PESO,    "weighing, grease, tip"),
    ("U", "DIRECT\nCOST",       11, "calc", PESO,    "fuel + toll + other"),
    ("V", "TRIP NET\nPROFIT",   13, "calc", PESO,    "fee \u2013 crew \u2013 costs"),
    ("W", "BILLING",            10, "in",   None,    "Paid / Unpaid"),
    ("X", "REMARKS",            26, "in",   None,    ""),
]
T_LAST = "X"
for letter, head, w, kind, fmt, nt in TRIP_COLS:
    T.column_dimensions[letter].width = w
T.column_dimensions["Y"].width = 2
T.column_dimensions["Z"].hidden = True

title_bar(T, T_LAST, "TRIPS  \u2014  one row per delivery  (this is the only sheet you fill in daily)",
          "Type only in the WHITE columns. Grey columns calculate themselves. "
          "Leave the fuel surcharge blank to use the 18% default \u2014 leave allowance blank to use the default.")

# header row 4
for letter, head, w, kind, fmt, nt in TRIP_COLS:
    c = T[f"{letter}4"]
    c.value = head
    c.font = fnt(8.5, True, WHITE)
    c.fill = fill(TEAL if kind == "in" else NAVY2)
    c.alignment = al("center", "center", wrap=True)
    c.border = Border(left=side(WHITE), right=side(WHITE), top=side(NAVY), bottom=side(NAVY))
T.row_dimensions[4].height = 38

# totals strip row 3
put(T, "A3", "ALL-TIME TOTALS  \u2192", font=fnt(8.5, True, NAVY), align=al("left", indent=1))
SUM_LETTERS = ["K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V"]
for letter in SUM_LETTERS:
    put(T, f"{letter}3", f"=SUM({letter}{DATA_R1}:{letter}{DATA_R2})", fmt=PESO0,
        font=fnt(8.5, True, NAVY), bg=GREY_L, align=al("right", indent=1))
put(T, "B3", f'=COUNT({DATA_R1}:{DATA_R2})&" trips logged"', font=fnt(8.5, True, NAVY),
    bg=GREY_L, align=al("center"))
put(T, "H3", f"=SUM(H{DATA_R1}:H{DATA_R2})", fmt=TONS, font=fnt(8.5, True, NAVY), bg=GREY_L,
    align=al("center"))
put(T, "A3", "ALL-TIME TOTALS  \u2192", font=fnt(8.5, True, NAVY), bg=GREY_L, align=al("left", indent=1))
T.row_dimensions[3].height = 18

# data area
for r in range(DATA_R1, DATA_R2 + 1):
    shade = WHITE if (r - DATA_R1) % 2 == 0 else "FAFBFC"
    for letter, head, w, kind, fmt, nt in TRIP_COLS:
        c = T[f"{letter}{r}"]
        c.fill = fill(shade if kind == "in" else GREY_L)
        c.font = fnt(9.5, kind == "calc", INK if kind == "in" else MUTED)
        c.border = Border(bottom=side("E8ECEF"))
        if fmt:
            c.number_format = fmt
        c.alignment = al("center" if letter in "ABCW" else "left" if letter in "BDEFGX" else "right")
    # formulas
    r_ = r
    T[f"I{r}"] = (f'=IF($G{r}="","",'
                  f'IF($F{r}="",IFERROR(INDEX(Setup!$C${RATE_R1}:$C${RATE_R2},'
                  f'MATCH("|"&UPPER(TRIM($G{r})),Setup!$Q${RATE_R1}:$Q${RATE_R2},0)),"NO RATE"),'
                  f'IFERROR(INDEX(Setup!$C${RATE_R1}:$C${RATE_R2},'
                  f'MATCH(UPPER(TRIM($F{r}))&"|"&UPPER(TRIM($G{r})),Setup!$P${RATE_R1}:$P${RATE_R2},0)),'
                  f'IFERROR(INDEX(Setup!$C${RATE_R1}:$C${RATE_R2},'
                  f'MATCH(UPPER(TRIM($G{r}))&"|"&UPPER(TRIM($F{r})),Setup!$P${RATE_R1}:$P${RATE_R2},0)),"NO RATE"))))')
    T[f"K{r}"] = (f'=IF(OR($G{r}="",$H{r}="",NOT(ISNUMBER($I{r}))),"",'
                  f'ROUND($H{r}*$I{r}*(1+IF($J{r}="",{SURCH},$J{r})),2))')
    T[f"L{r}"] = f'=IF($K{r}="","",ROUND($K{r}*Setup!$B$13,2))'
    T[f"M{r}"] = f'=IF($K{r}="","",$K{r}-$L{r})'
    T[f"N{r}"] = (f'=IF(OR($M{r}="",$D{r}=""),"",ROUND($M{r}*IFERROR(IF(VLOOKUP($D{r},Setup!$A${CREW_R1}:$C${CREW_R2},3,FALSE)="",'
                  f'{SHARE_DRV},VLOOKUP($D{r},Setup!$A${CREW_R1}:$C${CREW_R2},3,FALSE)),{SHARE_DRV}),2))')
    T[f"O{r}"] = (f'=IF(OR($M{r}="",$E{r}=""),"",ROUND($M{r}*IFERROR(IF(VLOOKUP($E{r},Setup!$A${CREW_R1}:$C${CREW_R2},3,FALSE)="",'
                  f'{SHARE_HLP},VLOOKUP($E{r},Setup!$A${CREW_R1}:$C${CREW_R2},3,FALSE)),{SHARE_HLP}),2))')
    T[f"U{r}"] = f'=IF($K{r}="","",SUM($R{r}:$T{r}))'
    T[f"V{r}"] = (f'=IF($K{r}="","",$M{r}-(N($N{r})+N($O{r})+IF($P{r}="",{ALLOW},$P{r})+N($Q{r}))-N($U{r}))')
    T[f"Z{r}"] = f'=IF($A{r}="","",DATE(YEAR($A{r}),MONTH($A{r}),1))'
    T[f"Z{r}"].font = fnt(8, False, MUTED)

T.freeze_panes = "C5"
T.sheet_view.zoomScale = 90

# ---- dropdowns
def add_dv(ws, formula, cells, title, msg, allow_blank=True, kind="list", op=None):
    if kind == "list":
        dv = DataValidation(type="list", formula1=formula, allow_blank=allow_blank, showDropDown=False)
    else:
        dv = DataValidation(type=kind, operator=op or "greaterThanOrEqual", formula1=formula,
                            allow_blank=allow_blank)
    dv.error = msg
    dv.errorTitle = title
    dv.errorStyle = "warning"
    dv.prompt = msg
    dv.promptTitle = title
    dv.showInputMessage = True
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(cells)
    return dv

add_dv(T, LST_TRUCK, f"C{DATA_R1}:C{DATA_R2}", "Truck", "Pick a truck from the Setup sheet.")
add_dv(T, LST_CREW,  f"D{DATA_R1}:D{DATA_R2}", "Driver", "Pick a driver from Setup \u2192 CREW.")
add_dv(T, LST_CREW,  f"E{DATA_R1}:E{DATA_R2}", "Helper", "Pick a helper from Setup \u2192 CREW.")
add_dv(T, LST_SRC,   f"F{DATA_R1}:F{DATA_R2}", "Source", "Pick an origin from Setup \u2192 SOURCE list.")
add_dv(T, LST_DEST,  f"G{DATA_R1}:G{DATA_R2}", "Destination",
       "Pick a destination from the Setup rate table.")
add_dv(T, '"Paid,Unpaid,Partial"', f"W{DATA_R1}:W{DATA_R2}", "Billing", "Paid / Unpaid / Partial")
add_dv(T, "0", f"H{DATA_R1}:H{DATA_R2}", "Tons", "Enter the weight in tons (numbers only).", kind="decimal")
add_dv(T, "0", f"J{DATA_R1}:J{DATA_R2}", "Fuel surcharge", "Enter a rate such as 18% or 0.18.", kind="decimal")

# ---- conditional formatting
T.conditional_formatting.add(f"A{DATA_R1}:A{DATA_R2}",
    FormulaRule(formula=[f'AND($A{DATA_R1}<>"",$A{DATA_R1}>TODAY())'], fill=fill(AMBER_L)))
T.conditional_formatting.add(f"V{DATA_R1}:V{DATA_R2}",
    CellIsRule(operator="lessThan", formula=["0"], font=fnt(9.5, True, RED), fill=fill(RED_L)))
T.conditional_formatting.add(f"V{DATA_R1}:V{DATA_R2}",
    CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=fnt(9.5, True, GREEN)))
T.conditional_formatting.add(f"I{DATA_R1}:I{DATA_R2}",
    FormulaRule(formula=[f'$I{DATA_R1}="NO RATE"'], fill=fill(RED_L), font=fnt(9.5, True, RED)))
T.conditional_formatting.add(f"W{DATA_R1}:W{DATA_R2}",
    FormulaRule(formula=[f'$W{DATA_R1}="Unpaid"'], fill=fill(AMBER_L), font=fnt(9.5, True, AMBER)))
T.conditional_formatting.add(f"W{DATA_R1}:W{DATA_R2}",
    FormulaRule(formula=[f'$W{DATA_R1}="Paid"'], fill=fill(GREEN_L), font=fnt(9.5, True, GREEN)))
T.auto_filter.ref = f"A4:{T_LAST}{DATA_R2}"
T.print_title_rows = "4:4"
T.page_setup.orientation = "landscape"
T.page_setup.fitToWidth = 1
T.page_setup.fitToHeight = 0
T.sheet_properties.pageSetUpPr.fitToPage = True

# ======================================================================
#  EXPENSES  (input #2)
# ======================================================================
E = ws_exp
EXP_COLS = [
    ("A", "DATE",              11, "in",   DATE, ""),
    ("B", "EXPENSE TYPE",      24, "in",   None, "dropdown from Setup"),
    ("C", "DESCRIPTION",       38, "in",   None, "e.g. diesel, PMS, 5 tires"),
    ("D", "TRUCK",             12, "in",   None, "leave blank for office costs"),
    ("E", "PAID TO /\nLENDER", 20, "in",   None, "supplier, lender or crew name"),
    ("F", "AMOUNT",            13, "in",   PESO, ""),
    ("G", "PAYMENT\nMETHOD",   15, "in",   None, "cash / gcash / bank"),
    ("H", "RECEIPT?",          11, "in",   None, "Y or N"),
    ("I", "COST GROUP",        15, "calc", None, "auto from Setup"),
    ("J", "",                   1, "hidden", None, ""),
]
E_LAST = "I"
for letter, head, w, kind, fmt, nt in EXP_COLS:
    E.column_dimensions[letter].width = w
E.column_dimensions["J"].hidden = True

title_bar(E, E_LAST, "EXPENSES  \u2014  one row per receipt  (fuel, repairs, toll, loans, rent, cash advances)",
          "Do NOT log crew pay here \u2014 the Payroll sheet computes it from Trips. "
          "Cash advances and payroll releases are logged here so the Payroll sheet can clear them.")

for letter, head, w, kind, fmt, nt in EXP_COLS:
    if kind == "hidden":
        continue
    c = E[f"{letter}4"]
    c.value = head
    c.font = fnt(8.5, True, WHITE)
    c.fill = fill(AMBER if kind == "in" else NAVY2)
    c.alignment = al("center", "center", wrap=True)
    c.border = Border(left=side(WHITE), right=side(WHITE), top=side(NAVY), bottom=side(NAVY))
E.row_dimensions[4].height = 38

put(E, "A3", "ALL-TIME TOTALS  \u2192", font=fnt(8.5, True, NAVY), bg=GREY_L, align=al("left", indent=1))
put(E, "B3", f'=COUNT({DATA_R1}:{DATA_R2})&" entries logged"', font=fnt(8.5, True, NAVY), bg=GREY_L, align=al("center"))
put(E, "F3", f"=SUM(F{DATA_R1}:F{DATA_R2})", fmt=PESO0, font=fnt(8.5, True, NAVY), bg=GREY_L, align=al("right", indent=1))
E.row_dimensions[3].height = 18

for r in range(DATA_R1, DATA_R2 + 1):
    shade = WHITE if (r - DATA_R1) % 2 == 0 else "FAFBFC"
    for letter, head, w, kind, fmt, nt in EXP_COLS:
        if kind == "hidden":
            continue
        c = E[f"{letter}{r}"]
        c.fill = fill(shade if kind == "in" else GREY_L)
        c.font = fnt(9.5, kind == "calc", INK if kind == "in" else MUTED)
        c.border = Border(bottom=side("E8ECEF"))
        if fmt:
            c.number_format = fmt
        c.alignment = al("center" if letter in "ADGH" else "left" if letter in "BCE" else "right")
    E[f"I{r}"] = (f'=IF($B{r}="","",IFERROR(INDEX(Setup!$B${EXP_R1}:$B${EXP_R2},'
                  f'MATCH($B{r},Setup!$A${EXP_R1}:$A${EXP_R2},0)),"CHECK TYPE"))')
    E[f"J{r}"] = f'=IF($A{r}="","",DATE(YEAR($A{r}),MONTH($A{r}),1))'

E.freeze_panes = "B5"
E.sheet_view.zoomScale = 90
add_dv(E, LST_TYPE, f"B{DATA_R1}:B{DATA_R2}", "Expense type", "Pick a type from Setup \u2192 EXPENSE TYPES.")
add_dv(E, LST_TRUCK, f"D{DATA_R1}:D{DATA_R2}", "Truck", "Pick a truck, or leave blank for office costs.")
add_dv(E, LST_METH, f"G{DATA_R1}:G{DATA_R2}", "Payment method", "Pick a payment method.")
add_dv(E, '"Y,N"', f"H{DATA_R1}:H{DATA_R2}", "Receipt", "Y or N")
add_dv(E, "0", f"F{DATA_R1}:F{DATA_R2}", "Amount", "Numbers only.", kind="decimal")
E.conditional_formatting.add(f"I{DATA_R1}:I{DATA_R2}",
    FormulaRule(formula=[f'$I{DATA_R1}="CHECK TYPE"'], fill=fill(RED_L), font=fnt(9.5, True, RED)))
E.conditional_formatting.add(f"H{DATA_R1}:H{DATA_R2}",
    FormulaRule(formula=[f'$H{DATA_R1}="N"'], fill=fill(AMBER_L), font=fnt(9.5, True, AMBER)))
E.auto_filter.ref = f"A4:{E_LAST}{DATA_R2}"
E.print_title_rows = "4:4"
E.page_setup.orientation = "landscape"
E.page_setup.fitToWidth = 1
E.page_setup.fitToHeight = 0
E.sheet_properties.pageSetUpPr.fitToPage = True

# ======================================================================
#  PAYROLL  (automatic: monthly settlement per crew member)
# ======================================================================
P = ws_pay
PAY_COLS = [
    ("A", "CREW NAME",            20, 20), ("B", "ROLE",          11, 11),
    ("C", "TRIPS",                 8,  8), ("D", "GROSS FEE HANDLED", 17, 17),
    ("E", "SHARE EARNED",         14, 14), ("F", "ALLOWANCE",     12, 12),
    ("G", "REIMBURSED",           12, 12), ("H", "TOTAL EARNED",  14, 14),
    ("I", "CASH ADVANCE",         14, 14), ("J", "PAYOUT RELEASED", 16, 16),
    ("K", "BALANCE\n(to pay / owes)", 17, 17), ("L", "STATUS",      18, 18),
]
for letter, head, w, _ in PAY_COLS:
    P.column_dimensions[letter].width = w
P.column_dimensions["M"].width = 2
P.column_dimensions["R"].hidden = True

title_bar(P, "L", "PAYROLL  \u2014  automatic, per month, per crew member",
          "Nothing to type here. Salary is computed from the driver/helper share on every trip; "
          "cash advances & payouts are picked up from the Expenses sheet.")

for letter, head, w, _ in PAY_COLS:
    c = P[f"{letter}4"]
    c.value = head
    c.font = fnt(8.5, True, WHITE)
    c.fill = fill(GREEN)
    c.alignment = al("center", "center", wrap=True)
    c.border = Border(left=side(WHITE), right=side(WHITE), top=side(GREEN), bottom=side(GREEN))
P.row_dimensions[4].height = 38
P.freeze_panes = "A5"

row = 5
PAY_MONTHS = 10 if TEST_MODE else 24
PAY_SUBTOTALS = []
for mi in range(PAY_MONTHS):
    mon = MONTHS[mi]
    blk = row
    P.merge_cells(start_row=blk, start_column=1, end_row=blk, end_column=12)
    cell = P.cell(row=blk, column=1, value=f"{mon.strftime('%B %Y').upper()}   \u00b7   crew earnings & settlement")
    cell.font = fnt(11, True, WHITE)
    cell.fill = fill(TEAL)
    cell.alignment = al("left", "center", indent=1)
    P.row_dimensions[blk].height = 24
    row += 1
    first_crew = row
    for ci in range(CREW_SLOTS):
        r = row
        src = CREW_R1 + ci
        P[f"R{r}"] = mon
        P[f"R{r}"].number_format = "yyyy-mm-dd"
        P[f"A{r}"] = f'=IF(Setup!$A{src}="","",Setup!$A{src})'
        P[f"B{r}"] = f'=IF($A{r}="","",Setup!$B{src})'
        guard = f'IF($A{r}="","",'
        P[f"C{r}"] = (f'={guard}IF(COUNTIFS({rng("Trips","$Z")},$R{r},{rng("Trips","$D")},$A{r})'
                      f'+COUNTIFS({rng("Trips","$Z")},$R{r},{rng("Trips","$E")},$A{r})=0,"",'
                      f'COUNTIFS({rng("Trips","$Z")},$R{r},{rng("Trips","$D")},$A{r})'
                      f'+COUNTIFS({rng("Trips","$Z")},$R{r},{rng("Trips","$E")},$A{r})))')
        g2 = f'IF($C{r}="","",'
        P[f"D{r}"] = (f'={g2}SUMIFS({rng("Trips","$K")},{rng("Trips","$Z")},$R{r},{rng("Trips","$D")},$A{r})'
                      f'+SUMIFS({rng("Trips","$K")},{rng("Trips","$Z")},$R{r},{rng("Trips","$E")},$A{r}))')
        P[f"E{r}"] = (f'={g2}SUMIFS({rng("Trips","$N")},{rng("Trips","$Z")},$R{r},{rng("Trips","$D")},$A{r})'
                      f'+SUMIFS({rng("Trips","$O")},{rng("Trips","$Z")},$R{r},{rng("Trips","$E")},$A{r}))')
        P[f"F{r}"] = f'={g2}SUMIFS({rng("Trips","$P")},{rng("Trips","$Z")},$R{r},{rng("Trips","$D")},$A{r}))'
        P[f"G{r}"] = f'={g2}SUMIFS({rng("Trips","$Q")},{rng("Trips","$Z")},$R{r},{rng("Trips","$D")},$A{r}))'
        P[f"I{r}"] = (f'={guard}IFERROR(1/(1/SUMIFS({EX_F},{EX_B},"Cash Advance",{EX_J},$R{r},'
                      f'{rng("Expenses","$E")},$A{r})),""))')
        P[f"J{r}"] = (f'={guard}IFERROR(1/(1/SUMIFS({EX_F},{EX_B},"Payroll Payout",{EX_J},$R{r},'
                      f'{rng("Expenses","$E")},$A{r})),""))')
        P[f"H{r}"] = (f'=IF(AND($C{r}="",$I{r}="",$J{r}=""),"",N($E{r})+N($F{r})+N($G{r}))')
        P[f"K{r}"] = f'=IF($H{r}="","",N($H{r})-N($I{r})-N($J{r}))'
        P[f"L{r}"] = (f'=IF($K{r}="","",IF($K{r}>0.5,"TO PAY",IF($K{r}<-0.5,"ADVANCE / OWES","SETTLED")))')
        for col, fmt, bold in [("C", '0', False), ("D", PESO0, False), ("E", PESO, True),
                               ("F", PESO, False), ("G", PESO, False), ("H", PESO, True),
                               ("I", PESO, False), ("J", PESO, False), ("K", PESO, True)]:
            c = P[f"{col}{r}"]
            c.number_format = fmt
            c.font = fnt(9.5, bold, INK if col != "K" else NAVY)
            c.alignment = al("center" if col == "C" else "left" if col in "AB" else "right")
        P[f"A{r}"].font = fnt(9.5, True, NAVY)
        P[f"A{r}"].alignment = al("left", indent=1)
        P[f"B{r}"].alignment = al("center")
        P[f"L{r}"].alignment = al("center")
        P[f"L{r}"].font = fnt(9, True, MUTED)
        row += 1
    # subtotal
    st = row
    PAY_SUBTOTALS.append(st)
    P.merge_cells(start_row=st, start_column=1, end_row=st, end_column=2)
    c = P.cell(row=st, column=1, value="MONTH TOTAL")
    c.font = fnt(9.5, True, NAVY); c.fill = fill(GREEN_L); c.alignment = al("left", indent=1)
    P.cell(row=st, column=2).fill = fill(GREEN_L)
    for col, fmt in [("C", '0'), ("D", PESO0), ("E", PESO), ("F", PESO), ("G", PESO),
                     ("H", PESO), ("I", PESO), ("J", PESO), ("K", PESO)]:
        c = P[f"{col}{st}"]
        c.value = (f'=IF(COUNT({col}{first_crew}:{col}{st-1})=0,"",SUM({col}{first_crew}:{col}{st-1}))')
        c.number_format = fmt
        c.font = fnt(9.5, True, NAVY)
        c.fill = fill(GREEN_L)
        c.border = Border(top=side(GREEN), bottom=side(GREY_M))
        c.alignment = al("center" if col == "C" else "right", "center")
    for col in "ABCDEFGHIJKL":
        P[f"{col}{st}"].border = Border(top=side(GREEN), bottom=side(GREY_M))
    row += 1

PAY_END = row - 1
P.conditional_formatting.add(f"K5:K{PAY_END}",
    CellIsRule(operator="greaterThan", formula=["0.5"], font=fnt(9.5, True, GREEN), fill=fill(GREEN_L)))
P.conditional_formatting.add(f"K5:K{PAY_END}",
    CellIsRule(operator="lessThan", formula=["-0.5"], font=fnt(9.5, True, RED), fill=fill(RED_L)))
P.conditional_formatting.add(f"L5:L{PAY_END}",
    FormulaRule(formula=['$L5="TO PAY"'], font=fnt(9, True, GREEN)))
P.conditional_formatting.add(f"L5:L{PAY_END}",
    FormulaRule(formula=['$L5="ADVANCE / OWES"'], font=fnt(9, True, RED)))
note(P, f"A{PAY_END+2}",
     "BALANCE = total earned \u2013 cash advance \u2013 payout released. Positive = you still owe the crew. "
     "Negative = the crew still owes you (unsettled advance).", size=8.5)
note(P, f"A{PAY_END+3}",
     "To cover a new year, copy the last month block (block header + crew rows + total row) and change the month key in the hidden column R.",
     size=8.5)

# ======================================================================
#  CALC extras: P&L lines + truck roll-up (columns Q..V, out of the way)
# ======================================================================
def pl(col):
    return (f'=IF($B$4="All",SUMIFS({rng("Trips",col)},{TR_Z},$B$3),'
            f'SUMIFS({rng("Trips",col)},{TR_Z},$B$3,{TR_C},$B$4))')

def grp(g):
    return (f'=IF($B$4="All",SUMIFS({EX_F},{EX_J},$B$3,{EX_I},"{g}"),'
            f'SUMIFS({EX_F},{EX_J},$B$3,{EX_I},"{g}",{EX_D},$B$4))')

PL_ROWS = [
    (7,  "Gross hauling fee",                       pl("$K")),
    (8,  "Withholding tax (2%)",                    pl("$L")),
    (9,  "Net hauling fee",                         '=Q7-Q8'),
    (10, "Driver share",                            pl("$N")),
    (11, "Helper share",                            pl("$O")),
    (12, "Driver allowance",                        pl("$P")),
    (13, "Reimbursable expenses",                   pl("$Q")),
    (14, "Gross profit",                            '=Q9-Q10-Q11-Q12-Q13'),
    (15, "Fuel",                                    pl("$R")),
    (16, "Toll & parking",                          pl("$S")),
    (17, "Other trip costs",                        pl("$T")),
    (18, "Total trip cost",                         '=SUM(Q15:Q17)'),
    (19, "Trip net profit",                         '=Q14-Q18'),
    (20, "Vehicle running costs (repairs, tires\u2026)", grp("Vehicle")),
    (21, "Overhead (rent, office, salaries\u2026)",   grp("Overhead")),
    (22, "Operating profit",                        '=Q19-Q20-Q21'),
    (23, "Loan amortisation paid",                  grp("Financing")),
    (24, "Cash profit after loans",                 '=Q22-Q23'),
    (25, "Margin on hauling fee",                   '=IF(Q7=0,"",Q24/Q7)'),
    (26, "Memo: advances & payouts (not a cost)", grp("Transfer")),
]
put(C_, "Q5", "P&L \u2014 selected month (filtered)", font=fnt(9, True, WHITE), bg=NAVY)
put(C_, "R5", "LINE", font=fnt(9, True, WHITE), bg=NAVY)
for r, lab, formula in PL_ROWS:
    put(C_, f"Q{r}", formula, fmt=PCT if r == 25 else PESO, font=fnt(9))
    put(C_, f"R{r}", lab, font=fnt(8.5, False, MUTED))

# truck roll-up
TRK_HDR = 30
put(C_, f"Q{TRK_HDR}", "TRUCK ROLL-UP (selected month)", font=fnt(9, True, WHITE), bg=NAVY)
for j, h in enumerate(["TRUCK", "TRIPS", "TONS", "GROSS FEE", "TRIP PROFIT", "MARGIN"], 1):
    put(C_, f"{get_column_letter(16 + j)}{TRK_HDR+1}", h, font=fnt(8, True, NAVY), bg=GREY_L, align=al("center"))
for i in range(TRUCK_SLOTS):
    r = TRK_HDR + 2 + i
    srow = TRUCK_R1 + i
    put(C_, f"Q{r}", f'=IF(Setup!$A{srow}="","",Setup!$A{srow})', font=fnt(9))
    put(C_, f"R{r}", f'=IF($Q{r}="","",COUNTIFS({TR_Z},$B$3,{TR_C},$Q{r}))', fmt='0', font=fnt(9))
    put(C_, f"S{r}", f'=IF($Q{r}="","",SUMIFS({rng("Trips","$H")},{TR_Z},$B$3,{TR_C},$Q{r}))', fmt=TONS, font=fnt(9))
    put(C_, f"T{r}", f'=IF($Q{r}="","",SUMIFS({rng("Trips","$K")},{TR_Z},$B$3,{TR_C},$Q{r}))', fmt=PESO, font=fnt(9))
    put(C_, f"U{r}", f'=IF($Q{r}="","",SUMIFS({rng("Trips","$V")},{TR_Z},$B$3,{TR_C},$Q{r}))', fmt=PESO, font=fnt(9))
    put(C_, f"V{r}", f'=IF($Q{r}="","",IF($T{r}=0,"",$U{r}/$T{r}))', fmt=PCT, font=fnt(9))
TRK_R1, TRK_R2 = TRK_HDR + 2, TRK_HDR + 1 + TRUCK_SLOTS

# ======================================================================
#  SAMPLE DATA  (delete the rows when you start for real)
# ======================================================================
RATE_MAP = {(a, b): c for a, b, c in RATES}
# lanes actually worked by the demo truck: (source, destination, tons, toll budget)
LANES = [
    ("BATAAN", "TARLAC", 18, 620), ("BATAAN", "BMPP", 20, 560),
    ("ORION", "BMPP", 22, 540), ("MABINI, BATANGAS", "BMPP", 20, 980),
    ("BATAAN", "ISABELA", 18, 1250), ("ORION", "TARLAC", 20, 610),
    ("MABINI, BATANGAS", "TARLAC", 18, 900), ("BATAAN", "ARMOR", 20, 520),
    ("ORION", "BULACAN", 22, 590), ("MABINI, BATANGAS", "MONCADA", 18, 430),
    ("BATAAN", "BULACAN", 20, 480), ("ORION", "ROSALES", 20, 700),
    ("MABINI, BATANGAS", "BATAAN", 18, 880), ("BATAAN", "BMEG BATAAN", 20, 120),
]
CREWS = [("R. Dela Cruz", "J. Santos"), ("M. Reyes", "A. Bautista")]
TR = []          # date, invoice, truck, driver, helper, source, dest, tons, surch, allow, reimb, fuel, toll, other, billing, remarks
inv = 1001
MONTH_TRIPS = [(2026, 8, 14), (2026, 9, 16), (2026, 10, 15)]
for trip_no, (y, m, n) in enumerate(MONTH_TRIPS):
    lane_pool = list(LANES)
    for i in range(n):
        src, dst, tons, toll = lane_pool[i % len(lane_pool)]
        tons = tons + ((i + trip_no) % 3) - 1          # small weight variation
        day = 1 + (i * 28) // max(1, n)                 # spread over the month
        day = min(day, 30 if m in (4, 6, 9, 11) else (28 if m == 2 else 31))
        driver, helper = CREWS[(i + trip_no) % 2]
        gross = tons * RATE_MAP[(src, dst)] * 1.18
        fuel = round(gross * (0.34 + 0.03 * ((i + trip_no) % 3)) / 100) * 100
        reimb = 450 if (i + trip_no) % 5 == 0 else 0
        allow = 400 if dst == "BMEG BATAAN" else 500
        billing = "Paid"
        if trip_no == 2 and i >= 9:
            billing = "Unpaid"
        elif trip_no == 1 and i >= 14:
            billing = "Partial"
        remark = ""
        if i == 4:
            remark = "long haul \u2014 diesel receipt attached"
        elif i == 9:
            remark = "helper not available, driver solo"
        elif trip_no == 2 and i == 12:
            remark = "surcharge raised to 20%"
        TR.append((dt.date(y, m, day), f"DR-{inv}", "TRUCK 1", driver, helper, src, dst, tons,
                   0.20 if (trip_no == 2 and i == 12) else None, allow, reimb, fuel, toll, 0,
                   billing, remark))
        inv += 1

for i, row in enumerate(TR):
    r = DATA_R1 + i
    (d, inv, trk, drv, hlp, src, dst, tons, surch, allow, reimb, fuel, toll, other, bill, rem) = row
    T[f"A{r}"] = d
    T[f"B{r}"] = inv
    T[f"C{r}"] = trk
    T[f"D{r}"] = drv
    T[f"E{r}"] = hlp
    T[f"F{r}"] = src
    T[f"G{r}"] = dst
    T[f"H{r}"] = tons
    if surch is not None:
        T[f"J{r}"] = surch
    T[f"P{r}"] = allow
    T[f"Q{r}"] = reimb
    T[f"R{r}"] = fuel
    T[f"S{r}"] = toll
    T[f"T{r}"] = other
    T[f"W{r}"] = bill
    T[f"X{r}"] = rem

EX = [      # date, type, description, truck, paid to, amount, method, receipt
    (dt.date(2026, 8, 5),  "Repairs & Maintenance", "PMS + brake adjustment", "TRUCK 1", "Casa Mechanic", 8500, "Cash", "Y"),
    (dt.date(2026, 8, 12), "Insurance", "Comprehensive premium (quarterly)", "TRUCK 1", "Fortune Insurance", 18500, "Bank Transfer", "Y"),
    (dt.date(2026, 8, 31), "Loan Payment", "August amortisation", "TRUCK 1", "BDO Leasing", 32000, "Bank Transfer", "Y"),
    (dt.date(2026, 8, 31), "Loan Payment", "August amortisation", "TRUCK 1", "Security Bank", 18500, "Bank Transfer", "Y"),
    (dt.date(2026, 8, 31), "Loan Payment", "August amortisation", "TRUCK 1", "Private Lender", 12500, "GCash", "N"),
    (dt.date(2026, 8, 31), "Rent", "Yard rent", "", "Mrs. Lim", 5000, "Cash", "N"),
    (dt.date(2026, 8, 31), "Payroll Payout", "August release", "", "R. Dela Cruz", 13000, "Cash", "N"),
    (dt.date(2026, 8, 31), "Payroll Payout", "August release", "", "M. Reyes", 13000, "Cash", "N"),
    (dt.date(2026, 9, 4),  "Tires & Batteries", "2 pcs rear tires", "TRUCK 1", "Goodyear Dealer", 21000, "Bank Transfer", "Y"),
    (dt.date(2026, 9, 9),  "Repairs & Maintenance", "Clutch lining + labour", "TRUCK 1", "Diesel Works", 6200, "Cash", "Y"),
    (dt.date(2026, 9, 10), "LTO Registration", "Registration + emission test", "TRUCK 1", "LTO", 4200, "Cash", "Y"),
    (dt.date(2026, 9, 15), "Cash Advance", "Advance — family emergency", "", "R. Dela Cruz", 5000, "Cash", "N"),
    (dt.date(2026, 9, 20), "Office / Staff Salary", "Dispatcher salary", "", "L. Garcia", 12000, "Cash", "N"),
    (dt.date(2026, 9, 30), "Loan Payment", "September amortisation", "TRUCK 1", "BDO Leasing", 32000, "Bank Transfer", "Y"),
    (dt.date(2026, 9, 30), "Loan Payment", "September amortisation", "TRUCK 1", "Security Bank", 18500, "Bank Transfer", "Y"),
    (dt.date(2026, 9, 30), "Loan Payment", "September amortisation", "TRUCK 1", "Private Lender", 12500, "GCash", "N"),
    (dt.date(2026, 9, 30), "Rent", "Yard rent", "", "Mrs. Lim", 5000, "Cash", "N"),
    (dt.date(2026, 9, 30), "Utilities", "Water + internet", "", "Meralco / PLDT", 2800, "GCash", "Y"),
    (dt.date(2026, 9, 30), "Truck Washing", "Wash + greasing", "TRUCK 1", "Jomar Carwash", 700, "Cash", "N"),
    (dt.date(2026, 10, 3), "Cash Advance", "Trip pocket money", "", "J. Santos", 3000, "Cash", "N"),
    (dt.date(2026, 10, 8), "Spare Parts & Lubricants", "Oil filters + coolant", "TRUCK 1", "Truck Parts Center", 3900, "Cash", "Y"),
    (dt.date(2026, 10, 10), "Truck Washing", "Wash + greasing", "TRUCK 1", "Jomar Carwash", 700, "Cash", "N"),
    (dt.date(2026, 10, 12), "Owner's Draw", "Owner's personal withdrawal", "", "Owner", 15000, "Cash", "N"),
    (dt.date(2026, 10, 15), "Loan Payment", "October amortisation", "TRUCK 1", "BDO Leasing", 32000, "Bank Transfer", "Y"),
    (dt.date(2026, 10, 15), "Loan Payment", "October amortisation", "TRUCK 1", "Security Bank", 18500, "Bank Transfer", "Y"),
    (dt.date(2026, 10, 16), "Rent", "Yard rent", "", "Mrs. Lim", 5000, "Cash", "N"),
    (dt.date(2026, 10, 17), "Utilities", "Water + internet", "", "Meralco / PLDT", 2800, "GCash", "Y"),
    (dt.date(2026, 10, 19), "Payroll Payout", "Mid-month release", "", "R. Dela Cruz", 10000, "GCash", "N"),
    (dt.date(2026, 10, 20), "Bank Charges", "InstaPay transfers", "", "BDO", 350, "Bank Transfer", "Y"),
]
for i, row in enumerate(EX):
    r = DATA_R1 + i
    d, typ, desc, trk, paid, amt, meth, rcpt = row
    E[f"A{r}"] = d
    E[f"B{r}"] = typ
    E[f"C{r}"] = desc
    E[f"D{r}"] = trk
    E[f"E{r}"] = paid
    E[f"F{r}"] = amt
    E[f"G{r}"] = meth
    E[f"H{r}"] = rcpt

# ======================================================================
#  CALC extras: filter lists + snap month
# ======================================================================
C_["B3"] = '=IF(Dashboard!$C$4="","",DATE(YEAR(Dashboard!$C$4),MONTH(Dashboard!$C$4),1))'
# previous month reference figures (dashboard "vs last month" strip)
put(C_, "B40", f'=SUMIFS({rng("Trips","$K")},{TR_Z},EDATE($B$3,-1))', fmt=PESO, font=fnt(9, False, MUTED))
put(C_, "B41", f'=COUNTIFS({TR_Z},EDATE($B$3,-1))', fmt='0', font=fnt(9, False, MUTED))
put(C_, "B42", f'=SUMIFS({rng("Trips","$H")},{TR_Z},EDATE($B$3,-1))', fmt=TONS, font=fnt(9, False, MUTED))
put(C_, "A40", "previous month fee", font=fnt(8, False, MUTED))
put(C_, "A41", "previous month trips", font=fnt(8, False, MUTED))
put(C_, "A42", "previous month tons", font=fnt(8, False, MUTED))

put(C_, "AA1", "MONTH LIST (dropdown)", font=fnt(8, True, MUTED))
for i in range(M_R2 - M_R1 + 1):
    put(C_, f"AA{2+i}", f"=$A${M_R1+i}", fmt="yyyy-mm-dd", font=fnt(8))
put(C_, "AB1", "All", font=fnt(8))
for i in range(TRUCK_SLOTS):
    put(C_, f"AB{2+i}", f'=IF(Setup!$A${TRUCK_R1+i}="","",Setup!$A${TRUCK_R1+i})', font=fnt(8))
MONTH_LIST = f"=Calc!$AA$2:$AA${1 + (M_R2 - M_R1 + 1)}"
TRUCK_LIST = f"=Calc!$AB$1:$AB${1 + TRUCK_SLOTS}"

# ======================================================================
#  DASHBOARD
# ======================================================================
D = ws_dash
widths(D, {"A": 2.5, "B": 17.5, "C": 17.5, "D": 17.5, "E": 17.5, "F": 17.5, "G": 17.5,
           "H": 17.5, "I": 17.5, "J": 17.5, "K": 17.5, "L": 17.5, "M": 17.5, "N": 2.5})

D.merge_cells("A1:M1")
put(D, "A1", "\U0001F69A  J&Z TRUCKING  \u00b7  COMMAND CENTER", font=fnt(20, True, WHITE), bg=NAVY,
    align=al("left", "center", indent=1))
D.row_dimensions[1].height = 44
D.merge_cells("A2:M2")
put(D, "A2", "Live figures from every trip, every receipt and every loan. Change the two filters below \u2014 the whole page follows.",
    font=fnt(9, False, "D6E4F0", italic=True), bg=NAVY2, align=al("left", "center", indent=1))
D.row_dimensions[2].height = 20

# nav row 3
put(D, "A3", "", bg=WHITE)
put(D, "B3", "GO TO  \u25b8", font=fnt(8, True, MUTED), align=al("right"))
NAV = [("C3", "P&L", f"#'Dashboard'!B47"), ("D3", "Trucks", f"#'Dashboard'!B70"),
       ("E3", "Loans", f"#'Dashboard'!H70"), ("F3", "Destinations", f"#'Dashboard'!B84"),
       ("G3", "Crew pay", f"#'Dashboard'!B99"), ("H3", "Trips log", "#'Trips'!A1"),
       ("I3", "Expenses", "#'Expenses'!A1"), ("J3", "Payroll", "#'Payroll'!A1"),
       ("K3", "Setup", "#'Setup'!A1"), ("L3", "Read me", "#'Read Me'!A1")]
for ref, lab, link in NAV:
    c = D[ref]
    c.value = lab
    c.font = fnt(8.5, True, TEAL, underline="single")
    c.alignment = al("center")
    c.hyperlink = link

# filters row 4
put(D, "B4", "MONTH", font=fnt(8, True, MUTED), align=al("right"))
put(D, "C4", MONTHS[9], fmt="mmmm yyyy", font=fnt(12, True, NAVY), bg=AMBER_L,
    border=box(AMBER), align=al("center"))
put(D, "D4", "\u25c0 change", font=fnt(7.5, False, MUTED, italic=True), align=al("left"))
put(D, "E4", "TRUCK", font=fnt(8, True, MUTED), align=al("right"))
put(D, "F4", "All", font=fnt(12, True, NAVY), bg=AMBER_L, border=box(AMBER), align=al("center"))
put(D, "G4", "all units", font=fnt(7.5, False, MUTED, italic=True), align=al("left"))
D.merge_cells("H4:M4")
put(D, "H4",
    '="vs previous month   \u25b8   fee "&IF(N(Calc!$B$40)=0,"n/a",'
    'IF(Calc!$B$7>=Calc!$B$40,"+","")&ROUND((Calc!$B$7/Calc!$B$40-1)*100,1)&"%")'
    '&"      trips "&IF(N(Calc!$B$41)=0,"n/a",IF(Calc!$B$10>=Calc!$B$41,"+","")&(Calc!$B$10-Calc!$B$41))'
    '&"      tons "&IF(N(Calc!$B$42)=0,"n/a",'
    'IF(Calc!$B$11>=Calc!$B$42,"+","")&ROUND((Calc!$B$11/Calc!$B$42-1)*100,1)&"%")'
    '&"      (pick a truck to see that unit alone)"',
    font=fnt(8.5, False, MUTED, italic=True), align=al("left", "center"))
D.row_dimensions[4].height = 30

add_dv(D, MONTH_LIST, "C4", "Reporting month", "Pick a month from the list (this drives the whole dashboard).")
add_dv(D, TRUCK_LIST, "F4", "Truck filter", "Pick 'All' or one truck.")

# ------------- KPI cards
CARD_ROWS = [7, 11]
CARDS = [
    ("GROSS HAULING FEE",      "Calc!$B$7",   PESO0, "total fee invoiced this month", NAVY),
    ("NET HAULING FEE",        "Calc!$Q$9",   PESO0, "after the 2% withholding tax", NAVY),
    ("CREW COST",              "Calc!$B$15",  PESO0, "driver + helper share + allowance", AMBER),
    ("TRIP NET PROFIT",        "Calc!$Q$19",  PESO0, "after fuel, toll & crew", TEAL),
    ("OPERATING PROFIT",       "Calc!$Q$22",  PESO0, "after repairs, rent & overhead", GREEN),
    ("CASH PROFIT AFTER LOANS","Calc!$Q$24",  PESO0, "what is really left in your pocket", GREEN),
    ("PROFIT MARGIN",          "Calc!$Q$25",  PCT,   "cash profit \u00f7 gross fee", GREEN),
    ("TRIPS DELIVERED",        "Calc!$B$10",  "0",   "this month", NAVY),
    ("TONS HAULED",            "Calc!$B$11",  TONS,  "total weight moved", NAVY),
    ("STILL TO PAY CREW",      "Calc!$B$30",  PESO0, "unsettled balance this month", AMBER),
    ("LOAN BALANCE",           "Calc!$B$LOT", PESO0, "principal still owed, all lenders", RED),
    ("UNPAID INVOICES",        "Calc!$B$31",  PESO0, "customer fees not yet collected", RED),
]
C_["B30"] = '=IFERROR(SUMIFS(Payroll!$K$5:$K$%d,Payroll!$R$5:$R$%d,$B$3,Payroll!$K$5:$K$%d,">0.5"),0)' % (PAY_END, PAY_END, PAY_END)
C_["B31"] = (f'=SUMIFS({rng("Trips","$K")},{TR_Z},$B$3)-SUMIFS({rng("Trips","$K")},{TR_Z},$B$3,{rng("Trips","$W")},"Paid")'
             if False else
             f'=SUMIFS({rng("Trips","$K")},{TR_Z},$B$3)-SUMIFS({rng("Trips","$K")},{TR_Z},$B$3,{rng("Trips","$W")},"Paid")')
CARD_ROWS_X = [1, 2]   # 6 cards per row
for idx, (label, ref, fmt, sub, colour) in enumerate(CARDS):
    row_label = 7 if idx < 6 else 11
    pos = idx % 6
    c1 = 2 + pos * 2                       # B, D, F, H, J, L
    c2 = c1 + 1
    value = ref.replace("Calc!$B$LOT", f"Calc!$B${L_TOT+2}")
    if idx == 9:                           # still to pay crew lives in Calc!B30
        value = "Calc!$B$30"
    # label
    D.merge_cells(start_row=row_label, start_column=c1, end_row=row_label, end_column=c2)
    cell = D.cell(row=row_label, column=c1, value=label)
    cell.font = fnt(7.5, True, MUTED)
    cell.fill = fill(GREY_B)
    cell.alignment = al("center", "center")
    for cc in (c1, c2):
        D.cell(row=row_label, column=cc).border = Border(top=side(GREY_M), left=side(GREY_M), right=side(GREY_M))
    # value
    vrow = row_label + 1
    D.merge_cells(start_row=vrow, start_column=c1, end_row=vrow, end_column=c2)
    vc = D.cell(row=vrow, column=c1, value=f"={value}")
    vc.number_format = fmt
    vc.font = fnt(19, True, colour)
    vc.alignment = al("center", "center")
    for cc in (c1, c2):
        D.cell(row=vrow, column=cc).border = Border(left=side(GREY_M), right=side(GREY_M))
    D.row_dimensions[vrow].height = 32
    # sub
    srow = row_label + 2
    D.merge_cells(start_row=srow, start_column=c1, end_row=srow, end_column=c2)
    sc = D.cell(row=srow, column=c1, value=sub)
    sc.font = fnt(7.5, False, MUTED, italic=True)
    sc.alignment = al("center", "center")
    for cc in (c1, c2):
        D.cell(row=srow, column=cc).border = Border(bottom=side(GREY_M), left=side(GREY_M), right=side(GREY_M))
    D.row_dimensions[srow].height = 16

# negative-value colouring on the money cards
for row_label in (8, 12):
    for pos in range(6):
        c1 = 2 + pos * 2
        col = get_column_letter(c1)
        D.conditional_formatting.add(f"{col}{row_label}",
            CellIsRule(operator="lessThan", formula=["0"], font=fnt(19, True, RED)))

# ------------- charts
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.chart.label import DataLabelList

def text_cats(chart, ref):
    """point the category axis at a text range (month names, destinations...)"""
    for ser in chart.series:
        ser.cat = AxDataSource(strRef=StrRef(f=ref))

def add_chart(anchor, chart, title):
    chart.title = title
    chart.style = 2
    chart.height = 7.4
    chart.width = 12.6
    D.add_chart(chart, anchor)

line = LineChart()
line.y_axis.numFmt = '"\u20b1"#,##0'
line.y_axis.majorGridlines = None
data = Reference(C_, min_col=5, min_row=M_R1 - 1, max_col=5, max_row=M_R2)
line.add_data(data, titles_from_data=True)
line.add_data(Reference(C_, min_col=12, min_row=M_R1 - 1, max_col=12, max_row=M_R2), titles_from_data=True)
line.add_data(Reference(C_, min_col=13, min_row=M_R1 - 1, max_col=13, max_row=M_R2), titles_from_data=True)
line.set_categories(Reference(C_, min_col=2, min_row=M_R1, max_row=M_R2))
text_cats(line, f"Calc!$B${M_R1}:$B${M_R2}")
add_chart("B16", line, "Monthly trend \u2014 fee, total cost, net profit")

bar = BarChart()
bar.type = "col"
bar.grouping = "clustered"
bar.y_axis.numFmt = '"\u20b1"#,##0'
bar.add_data(Reference(C_, min_col=2, min_row=G_R1 - 1, max_row=G_R2), titles_from_data=True)
bar.set_categories(Reference(C_, min_col=1, min_row=G_R1, max_row=G_R2))
text_cats(bar, f"Calc!$A${G_R1}:$A${G_R2}")
bar.dataLabels = DataLabelList()
bar.dataLabels.showVal = True
bar.legend = None
add_chart("H16", bar, "Booked expenses by cost group (this month)")

pie = PieChart()
pie.add_data(Reference(C_, min_col=4, min_row=D_R1 - 1, max_row=D_R1 + 6), titles_from_data=True)
pie.set_categories(Reference(C_, min_col=1, min_row=D_R1, max_row=D_R1 + 6))
text_cats(pie, f"Calc!$A${D_R1}:$A${D_R1 + 6}")
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
add_chart("B31", pie, "Where the revenue comes from (by destination)")

trk_bar = BarChart()
trk_bar.type = "bar"
trk_bar.y_axis.numFmt = '"\u20b1"#,##0'
trk_bar.add_data(Reference(C_, min_col=21, min_row=TRK_HDR + 1, max_row=TRK_R2), titles_from_data=True)
trk_bar.set_categories(Reference(C_, min_col=17, min_row=TRK_R1, max_row=TRK_R2))
text_cats(trk_bar, f"Calc!$Q${TRK_R1}:$Q${TRK_R2}")
trk_bar.dataLabels = DataLabelList()
trk_bar.dataLabels.showVal = True
trk_bar.legend = None
add_chart("H31", trk_bar, "Trip net profit per truck (this month)")

# ------------- P&L table
P_SEC, P_HR = 47, 48
def dash_section(row, c1, c2, text):
    D.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = D.cell(row=row, column=c1, value=text)
    c.font = fnt(9.5, True, WHITE)
    c.fill = fill(NAVY2)
    c.alignment = al("left", "center", indent=1)
    D.row_dimensions[row].height = 24

dash_section(P_SEC, 2, 7, "PROFIT & LOSS  \u00b7  selected month")
D.merge_cells(start_row=P_HR, start_column=2, end_row=P_HR, end_column=3)
put(D, f"B{P_HR}", "LINE ITEM", font=fnt(8, True, MUTED), align=al("left", indent=1))
put(D, f"D{P_HR}", "AMOUNT", font=fnt(8, True, MUTED), align=al("right", indent=1))
put(D, f"E{P_HR}", "% OF FEE", font=fnt(8, True, MUTED), align=al("right", indent=1))
D.merge_cells(start_row=P_HR, start_column=6, end_row=P_HR, end_column=7)
put(D, f"F{P_HR}", "HOW IT IS BUILT", font=fnt(8, True, MUTED), align=al("left", indent=1))
for cc in range(2, 8):
    D.cell(row=P_HR, column=cc).border = Border(bottom=side(GREY_M))

PL_NOTE = {
    7: "tons \u00d7 rate \u00d7 (1 + fuel surcharge)", 8: "2% withheld by the client",
    9: "the base for the crew shares", 10: "driver share of net fee", 11: "helper share of net fee",
    12: "cash given per trip", 13: "driver receipts to pay back", 14: "after crew and reimbursables",
    15: "diesel logged per trip", 16: "expressway + parking", 17: "weighing, grease, tips",
    18: "", 19: "before truck running costs", 20: "repairs, tires, insurance, LTO, wash",
    21: "rent, office salary, utilities, fees", 22: "the operating result",
    23: "principal + interest to all lenders", 24: "money actually left with you",
    25: "cash profit \u00f7 gross fee", 26: "advances are a transfer, not a cost",
}
SUB = {9: "sub", 14: "sub", 19: "sub", 22: "sub", 24: "big"}
for i, (calcrow, label, _f) in enumerate(PL_ROWS):
    r = P_SEC + 2 + i
    style = SUB.get(calcrow, "")
    bgc = WHITE if not style else (GREEN_L if style == "big" else GREY_B)
    bold = style != ""
    size = 11 if style == "big" else 9.5
    colour = GREEN if style == "big" else (NAVY if bold else INK)
    D.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    lc = D.cell(row=r, column=2, value=label)
    lc.font = fnt(size, bold, colour)
    lc.fill = fill(bgc)
    lc.alignment = al("left", "center", indent=2 if calcrow > 7 else 1)
    D.cell(row=r, column=3).fill = fill(bgc)
    ac = D.cell(row=r, column=4, value=f"=Calc!$Q${calcrow}")
    ac.number_format = PCT if calcrow == 25 else PESO
    ac.font = fnt(size, bold, colour)
    ac.fill = fill(bgc)
    ac.alignment = al("right", "center", indent=1)
    pc = D.cell(row=r, column=5, value=('' if calcrow == 25 else f'=IF(Calc!$Q$7=0,"",D{r}/Calc!$Q$7)'))
    pc.number_format = PCT
    pc.font = fnt(8.5, False, MUTED)
    pc.fill = fill(bgc)
    pc.alignment = al("right", "center", indent=1)
    D.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
    nc = D.cell(row=r, column=6, value=PL_NOTE.get(calcrow, ""))
    nc.font = fnt(7.5, False, MUTED, italic=True)
    nc.fill = fill(bgc)
    nc.alignment = al("left", "center", indent=1)
    D.cell(row=r, column=7).fill = fill(bgc)
    for cc in range(2, 8):
        top = side(NAVY2) if style else side("E8ECEF")
        D.cell(row=r, column=cc).border = Border(top=top, bottom=side("E8ECEF"))
    D.row_dimensions[r].height = 22 if style == "big" else 18

# colour the money in the P&L by sign
for r in range(P_SEC + 2, P_SEC + 2 + len(PL_ROWS)):
    D.conditional_formatting.add(f"D{r}", CellIsRule(operator="lessThan", formula=["0"],
                                  font=fnt(9.5, True, RED)))

# ------------- trucks + loans side by side (rows 70+)
def table_head_dash(row, c1, labels, widths=None):
    for i, lab in enumerate(labels):
        c = D.cell(row=row, column=c1 + i, value=lab)
        c.font = fnt(7.5, True, MUTED)
        c.alignment = al("center", "center", wrap=True)
        c.border = Border(bottom=side(GREY_M))

TRK_SEC = 70
dash_section(TRK_SEC, 2, 7, "TRUCK PERFORMANCE  \u00b7  selected month")
table_head_dash(TRK_SEC + 1, 2, ["TRUCK", "TRIPS", "TONS", "GROSS FEE", "TRIP PROFIT", "MARGIN"])
for i in range(TRUCK_SLOTS):
    r = TRK_SEC + 2 + i
    crow = TRK_HDR + 2 + i
    D.cell(row=r, column=2, value=f"=Calc!$Q${crow}").font = fnt(9, True, NAVY)
    D.cell(row=r, column=2).alignment = al("left", indent=1)
    D.cell(row=r, column=3, value=f"=Calc!$R${crow}").number_format = '0'
    D.cell(row=r, column=4, value=f"=Calc!$S${crow}").number_format = TONS
    D.cell(row=r, column=5, value=f"=Calc!$T${crow}").number_format = PESO
    D.cell(row=r, column=6, value=f"=Calc!$U${crow}").number_format = PESO
    D.cell(row=r, column=7, value=f"=Calc!$V${crow}").number_format = PCT
    for cc in range(3, 8):
        D.cell(row=r, column=cc).font = fnt(9)
        D.cell(row=r, column=cc).alignment = al("center" if cc in (3, 4, 7) else "right", indent=1)
    for cc in range(2, 8):
        D.cell(row=r, column=cc).border = Border(bottom=side("E8ECEF"))
    D.row_dimensions[r].height = 18
TRK_TOT = TRK_SEC + 2 + TRUCK_SLOTS
D.cell(row=TRK_TOT, column=2, value="TOTAL").font = fnt(9, True, WHITE)
D.cell(row=TRK_TOT, column=2).fill = fill(TEAL)
for cc, fmt in [(3, '0'), (4, TONS), (5, PESO), (6, PESO), (7, PCT)]:
    L = get_column_letter(cc)
    c = D.cell(row=TRK_TOT, column=cc, value=f"=SUM({L}{TRK_SEC+2}:{L}{TRK_TOT-1})")
    c.number_format = fmt
    c.font = fnt(9, True, WHITE)
    c.fill = fill(TEAL)
    c.alignment = al("center", "center")
D.conditional_formatting.add(f"F{TRK_SEC+2}:F{TRK_TOT}",
    CellIsRule(operator="lessThan", formula=["0"], font=fnt(9, True, RED)))

LO_SEC = 70
dash_section(LO_SEC, 8, 13, "LOAN BOOK  \u00b7  all lenders, self-updating")
D.merge_cells(start_row=LO_SEC + 1, start_column=8, end_row=LO_SEC + 1, end_column=9)
put(D, f"H{LO_SEC+1}", "LENDER  \u00b7  TRUCK", font=fnt(7.5, True, MUTED), align=al("left", indent=1))
table_head_dash(LO_SEC + 1, 10, ["AMORT.", "BALANCE", "MONTHS LEFT", "PAID %"])
for cc in (8, 9):
    D.cell(row=LO_SEC + 1, column=cc).border = Border(bottom=side(GREY_M))
for i in range(LOAN_SLOTS):
    r = LO_SEC + 2 + i
    srow = LOAN_R1 + i
    D.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
    lc = D.cell(row=r, column=8, value=f'=IF(Setup!$A{srow}="","",Setup!$A{srow}&"  \u00b7  "&Setup!$B{srow})')
    lc.font = fnt(9, True, NAVY)
    lc.alignment = al("left", indent=1)
    D.cell(row=r, column=10, value=f"=Setup!$E{srow}").number_format = PESO
    D.cell(row=r, column=11, value=f"=Setup!$K{srow}").number_format = PESO
    D.cell(row=r, column=12, value=f"=Setup!$J{srow}").number_format = '0'
    D.cell(row=r, column=13, value=f"=Setup!$M{srow}").number_format = PCT
    for cc in range(10, 14):
        D.cell(row=r, column=cc).font = fnt(9)
        D.cell(row=r, column=cc).alignment = al("center" if cc in (12, 13) else "right", indent=1)
    for cc in range(8, 14):
        D.cell(row=r, column=cc).border = Border(bottom=side("E8ECEF"))
    D.row_dimensions[r].height = 18
LO_TOT = LO_SEC + 2 + LOAN_SLOTS
D.merge_cells(start_row=LO_TOT, start_column=8, end_row=LO_TOT, end_column=9)
c = D.cell(row=LO_TOT, column=8, value="TOTAL")
c.font = fnt(9, True, WHITE); c.fill = fill(RED); c.alignment = al("left", indent=1)
D.cell(row=LO_TOT, column=9).fill = fill(RED)
for cc, fmt, formula in [(10, PESO, f"=SUM(J{LO_SEC+2}:J{LO_TOT-1})"),
                         (11, PESO, f"=SUM(K{LO_SEC+2}:K{LO_TOT-1})"),
                         (12, '0', f"=SUM(L{LO_SEC+2}:L{LO_TOT-1})"),
                         (13, PCT, f'=IF(SUM(Setup!$C${LOAN_R1}:$C${LOAN_R2})=0,0,1-K{LO_TOT}/SUM(Setup!$C${LOAN_R1}:$C${LOAN_R2}))')]:
    c = D.cell(row=LO_TOT, column=cc, value=formula)
    c.number_format = fmt
    c.font = fnt(9, True, WHITE)
    c.fill = fill(RED)
    c.alignment = al("center" if cc in (12, 13) else "right", indent=1)

# ------------- destinations + expense groups (rows 84+)
DST_SEC = 84
dash_section(DST_SEC, 2, 7, "TOP DESTINATIONS  \u00b7  selected month")
table_head_dash(DST_SEC + 1, 2, ["DESTINATION", "TRIPS", "TONS", "GROSS FEE", "TRIP PROFIT", "MARGIN"])
for i in range(10):
    r = DST_SEC + 2 + i
    crow = D_R1 + i
    D.cell(row=r, column=2, value=f"=Calc!$J${crow}").font = fnt(9, True, NAVY)
    D.cell(row=r, column=2).alignment = al("left", indent=1)
    D.cell(row=r, column=3, value=f'=IF(Calc!$J${crow}="","",Calc!$K${crow})').number_format = '0'
    D.cell(row=r, column=4, value=f'=IF(Calc!$J${crow}="","",Calc!$L${crow})').number_format = TONS
    D.cell(row=r, column=5, value=f'=IF(Calc!$J${crow}="","",Calc!$M${crow})').number_format = PESO
    D.cell(row=r, column=6, value=f'=IF(Calc!$J${crow}="","",Calc!$N${crow})').number_format = PESO
    D.cell(row=r, column=7, value=f'=IF(Calc!$J${crow}="","",Calc!$O${crow})').number_format = PCT
    for cc in range(3, 8):
        D.cell(row=r, column=cc).font = fnt(9)
        D.cell(row=r, column=cc).alignment = al("center" if cc in (3, 4, 7) else "right", indent=1)
    for cc in range(2, 8):
        D.cell(row=r, column=cc).border = Border(bottom=side("E8ECEF"))
    D.row_dimensions[r].height = 18
note(D, f"B{DST_SEC+12}", "Full list (20 destinations) is on the Calc sheet \u2014 add new lanes in Setup \u00a78.", size=7.5, span=6)

GRP_SEC = 84
dash_section(GRP_SEC, 8, 13, "WHERE THE MONEY WENT  \u00b7  booked expenses")
D.merge_cells(start_row=GRP_SEC + 1, start_column=8, end_row=GRP_SEC + 1, end_column=10)
put(D, f"H{GRP_SEC+1}", "COST GROUP", font=fnt(7.5, True, MUTED), align=al("left", indent=1))
table_head_dash(GRP_SEC + 1, 11, ["AMOUNT", "SHARE", "COUNTED IN PROFIT?"])
for cc in (8, 9, 10):
    D.cell(row=GRP_SEC + 1, column=cc).border = Border(bottom=side(GREY_M))
for i, g in enumerate(GROUPS):
    r = GRP_SEC + 2 + i
    crow = G_R1 + i
    D.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10)
    c = D.cell(row=r, column=8, value=g)
    c.font = fnt(9, True, NAVY)
    c.alignment = al("left", indent=1)
    D.cell(row=r, column=11, value=f"=Calc!$B${crow}").number_format = PESO
    D.cell(row=r, column=12, value=f'=IF(Calc!$B${L_TOT}=0,"",K{r}/Calc!$B${L_TOT})').number_format = PCT
    verdict = {"Vehicle": "yes \u2014 running cost", "Financing": "yes \u2014 after operating profit",
               "Transfer": "no \u2014 advances", "Overhead": "yes \u2014 overhead",
               "Other": "no \u2014 owner / capital"}[g]
    vc = D.cell(row=r, column=13, value=verdict)
    vc.font = fnt(8, False, MUTED, italic=True)
    vc.alignment = al("left", indent=1)
    for cc in (11,):
        D.cell(row=r, column=cc).font = fnt(9)
        D.cell(row=r, column=cc).alignment = al("right", indent=1)
    D.cell(row=r, column=12).font = fnt(9)
    D.cell(row=r, column=12).alignment = al("center")
    for cc in range(8, 14):
        D.cell(row=r, column=cc).border = Border(bottom=side("E8ECEF"))
    D.row_dimensions[r].height = 18
GR_TOT = GRP_SEC + 2 + len(GROUPS)
D.merge_cells(start_row=GR_TOT, start_column=8, end_row=GR_TOT, end_column=10)
c = D.cell(row=GR_TOT, column=8, value="TOTAL BOOKED")
c.font = fnt(9, True, WHITE); c.fill = fill(AMBER); c.alignment = al("left", indent=1)
for cc in (9, 10):
    D.cell(row=GR_TOT, column=cc).fill = fill(AMBER)
c = D.cell(row=GR_TOT, column=11, value=f"=Calc!$B${L_TOT}")
c.number_format = PESO; c.font = fnt(9, True, WHITE); c.fill = fill(AMBER)
c.alignment = al("right", indent=1)
c = D.cell(row=GR_TOT, column=12, value=1)
c.number_format = PCT; c.font = fnt(9, True, WHITE); c.fill = fill(AMBER)
c.alignment = al("center")
D.cell(row=GR_TOT, column=13).fill = fill(AMBER)
note(D, f"H{GR_TOT+1}", "Only Vehicle / Overhead / Financing reach the P&L. Advances, payouts, owner's draw and asset buys are excluded.",
     size=7.5, span=6)

# ------------- crew settlement (rows 98+)
CRW_SEC = 99
dash_section(CRW_SEC, 2, 13, "CREW SETTLEMENT  \u00b7  selected month  (auto \u2014 mirrors the Payroll sheet)")
table_head_dash(CRW_SEC + 1, 2, ["CREW NAME", "ROLE", "TRIPS", "GROSS FEE HANDLED", "SHARE",
                                 "ALLOWANCE + REIMB.", "CASH ADVANCE", "PAYOUT RELEASED",
                                 "BALANCE", "STATUS"])
for i in range(CREW_SLOTS):
    r = CRW_SEC + 2 + i
    crow = W_R1 + i
    D.cell(row=r, column=2, value=f"=Calc!$A${crow}").font = fnt(9, True, NAVY)
    D.cell(row=r, column=2).alignment = al("left", indent=1)
    g = f'IF(Calc!$A${crow}="","",'
    D.cell(row=r, column=3, value=f"=Calc!$B${crow}").font = fnt(8.5, False, MUTED)
    D.cell(row=r, column=4, value=f"={g}Calc!$C${crow})").number_format = '0'
    D.cell(row=r, column=5, value=f"={g}Calc!$D${crow})").number_format = PESO
    D.cell(row=r, column=6, value=f"={g}Calc!$E${crow})").number_format = PESO
    D.cell(row=r, column=7, value=f"={g}Calc!$F${crow}+Calc!$G${crow})").number_format = PESO
    D.cell(row=r, column=8, value=f"={g}Calc!$H${crow})").number_format = PESO
    D.cell(row=r, column=9, value=f"={g}Calc!$I${crow})").number_format = PESO
    D.cell(row=r, column=10, value=f"={g}Calc!$J${crow})").number_format = PESO
    D.cell(row=r, column=11, value=f'=IF(B{r}="","",IF(AND(Calc!$E${crow}=0,Calc!$H${crow}=0,Calc!$I${crow}=0),"",'
                                    f'IF(J{r}>0.5,"TO PAY",IF(J{r}<-0.5,"OWES ADVANCE","SETTLED"))))')
    for cc in range(3, 12):
        cell = D.cell(row=r, column=cc)
        cell.font = fnt(9, cc == 10, NAVY if cc == 10 else INK)
        cell.alignment = al("center" if cc in (3, 4, 11) else "right", indent=0 if cc in (3, 4, 11) else 1)
    D.cell(row=r, column=11).font = fnt(8.5, True, MUTED)
    for cc in range(2, 12):
        D.cell(row=r, column=cc).border = Border(bottom=side("E8ECEF"))
    D.row_dimensions[r].height = 18
CR_TOT = CRW_SEC + 2 + CREW_SLOTS
D.cell(row=CR_TOT, column=2, value="TOTAL").font = fnt(9, True, WHITE)
D.cell(row=CR_TOT, column=2).fill = fill(GREEN)
D.cell(row=CR_TOT, column=3).fill = fill(GREEN)
for cc, fmt in [(4, '0'), (5, PESO), (6, PESO), (7, PESO), (8, PESO), (9, PESO), (10, PESO)]:
    L = get_column_letter(cc)
    c = D.cell(row=CR_TOT, column=cc, value=f"=SUM({L}{CRW_SEC+2}:{L}{CR_TOT-1})")
    c.number_format = fmt
    c.font = fnt(9, True, WHITE)
    c.fill = fill(GREEN)
    c.alignment = al("center" if cc == 4 else "right", indent=1)
D.cell(row=CR_TOT, column=11).fill = fill(GREEN)
D.conditional_formatting.add(f"J{CRW_SEC+2}:J{CR_TOT}",
    CellIsRule(operator="lessThan", formula=["-0.5"], font=fnt(9, True, RED)))
D.conditional_formatting.add(f"J{CRW_SEC+2}:J{CR_TOT}",
    CellIsRule(operator="greaterThan", formula=["0.5"], font=fnt(9, True, GREEN)))

note(D, f"B{CR_TOT+2}",
     "Every number above is a live formula. Nothing on this page is typed by hand \u2014 only the two filters at the top.",
     size=8, span=12)
D.freeze_panes = "A6"
D.sheet_view.zoomScale = 90
D.page_setup.orientation = "landscape"
D.page_setup.fitToWidth = 1
D.page_setup.fitToHeight = 0
D.sheet_properties.pageSetUpPr.fitToPage = True

# ======================================================================
#  READ ME
# ======================================================================
RM = ws_read
widths(RM, {"A": 2, "B": 3, "C": 100, "D": 2})
RM.merge_cells("A1:D1")
put(RM, "A1", "\U0001F4D8  READ ME  \u2014  how this workbook runs your trucking business",
    font=fnt(18, True, WHITE), bg=NAVY, align=al("left", "center", indent=1))
RM.row_dimensions[1].height = 42
RM.merge_cells("A2:D2")
put(RM, "A2", "5-minute setup, then 30 seconds per trip. Written for one truck today and ten trucks later.",
    font=fnt(9.5, False, "D6E4F0", italic=True), bg=NAVY2, align=al("left", "center", indent=1))
RM.row_dimensions[2].height = 20

row = 4
def rm_heading(text):
    global row
    RM.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    c = RM.cell(row=row, column=2, value=text)
    c.font = fnt(11, True, WHITE)
    c.fill = fill(TEAL)
    c.alignment = al("left", "center", indent=1)
    RM.row_dimensions[row].height = 24
    row += 1

def rm_body(text, size=9.5, bold=False, color=INK, bg=None, height=None, italic=False):
    global row
    RM.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    c = RM.cell(row=row, column=2, value=text)
    c.font = fnt(size, bold, color, italic=italic)
    c.alignment = al("left", "top", wrap=True, indent=1)
    if bg:
        c.fill = fill(bg)
    import math
    lines = max(1, math.ceil(len(text) / 118)) + text.count("\n")
    RM.row_dimensions[row].height = height or (12.2 * lines + 4)
    row += 1

def rm_gap(h=6):
    global row
    RM.row_dimensions[row].height = h
    row += 1

rm_heading("1 \u00b7 WHAT THIS WORKBOOK DOES")
rm_body("It answers four questions every single day: Did this trip make money? Did this month make money? "
        "What do I still owe the crew? What do I still owe the bank?")
rm_body("You type trips and receipts. The workbook computes the hauling fee from your rate table, adds the fuel "
        "surcharge, withholds the 2% tax, computes the driver's and helper's share, deducts fuel / toll / allowance / "
        "reimbursables, and rolls everything up into a profit & loss, a per-truck result, a loan book and payroll.")
rm_gap()

rm_heading("2 \u00b7 THE THREE RULES")
rm_body("\u2460  SETUP is filled once \u2014 rates, crew, trucks, loans, expense types.\n"
        "\u2461  TRIPS = one row per delivery.  EXPENSES = one row per receipt.\n"
        "\u2462  Never type a number that the workbook can compute. White columns are yours, grey columns are mine.",
        bold=False)
rm_gap()

rm_heading("3 \u00b7 THE SHEETS")
for name, desc in [
    ("DASHBOARD", "Read-only. Two filters at the top (month + truck) drive the whole page: 12 KPI cards, 4 charts, "
                  "the P&L, per-truck result, loan book, destination ranking and the crew settlement."),
    ("TRIPS", "The only sheet you touch daily. Date, invoice #, truck, driver, helper, source, destination, tons and the "
              "actual cash you spent per trip. Everything else in that row is computed."),
    ("EXPENSES", "One row per receipt: fuel, repairs, tires, insurance, LTO, rent, utilities, loan amortisation, "
                 "cash advances, payroll releases, owner's draw."),
    ("PAYROLL", "Fully automatic, monthly, one block per month, one row per crew member. Salary = the share earned "
                "on every trip. It also clears cash advances and payouts so you know the real balance per person."),
    ("SETUP", "Your master data: business profile, tax/surcharge/share percentages, the J&Z rate table, trucks, crew, "
              "expense types, loans and the dropdown lists."),
    ("CALC", "Hidden engine. Do not edit, do not delete \u2014 it holds the totals the dashboard reads."),
]:
    rm_body(f"\u25aa  {name}\n{desc}", size=9)
rm_gap()

rm_heading("4 \u00b7 HOW THE HAULING FEE IS COMPUTED (matches your J&Z rate sheet)")
rm_body("RATE  =  the amount on your J&Z Bulk Rate sheet for that source \u2192 destination, in \u20b1 per TON.\n"
        "GROSS FEE  =  TONS  \u00d7  RATE  \u00d7  ( 1 + FUEL SURCHARGE ).  The surcharge defaults to 18% from Setup \u00a72; "
        "type a different % on a single trip (for example 20%) and only that trip changes.\n"
        "2% TAX  =  GROSS FEE \u00d7 2%, withheld by the client.\n"
        "NET FEE  =  GROSS FEE \u2212 2% TAX.  This is the base for the crew shares.\n"
        "DRIVER  =  NET FEE \u00d7 10%     HELPER  =  NET FEE \u00d7 5%\n"
        "The percentages live in Setup \u00a72 and can be overridden per person in Setup \u00a75 (leave the % blank to use the default).",
        size=9)
rm_body("Example \u2014 BATAAN \u2192 BMPP at \u20b1950/ton, 20 tons:  950 \u00d7 20 = 19,000  \u00d7 1.18 = \u20b122,420 gross.  "
        "Tax 2% = \u20b1448.40  \u2192  net \u20b121,971.60.  Driver 10% = \u20b12,197.16, helper 5% = \u20b11,098.58.",
        size=9, bold=True, color=NAVY, bg=AMBER_L)
rm_gap()

rm_heading("5 \u00b7 WHAT COUNTS AS A COST (and what does not)")
rm_body("Trip profit  =  GROSS FEE \u2212 2% TAX \u2212 driver share \u2212 helper share \u2212 allowance \u2212 reimbursables "
        "\u2212 fuel \u2212 toll & parking \u2212 other trip cost.")
rm_body("Operating profit  =  trip profit (all trips in the month)  \u2212  booked Vehicle expenses  \u2212  booked Overhead "
        "expenses.    Cash profit after loans  =  operating profit  \u2212  loan amortisation actually paid.")
rm_body("NOT costs: cash advances, payroll payouts, reimbursable payouts, owner's draw and asset purchases. They are "
        "money moving between your pocket and your crew \u2014 the real cost was already booked when the trip or the "
        "receipt was logged. They are grouped as TRANSFER / OTHER so they never inflate your expenses.", size=9)
rm_body("Log fuel and toll per trip in the TRIPS sheet. Use EXPENSES \u2192 \u2018Fuel \u2013 Bulk Purchase\u2019 only when you buy "
        "in bulk, otherwise you will count the same diesel twice.", size=9, bold=True, color=RED)
rm_gap()

rm_heading("6 \u00b7 YOUR ROUTINE")
rm_body("\u2022  EVERY TRIP (30 seconds): Trips \u2192 new row \u2192 date, invoice #, truck, driver, helper, source, destination, tons, "
        "allowance, fuel, toll. Done \u2014 fee, tax, shares and profit appear by themselves.\n"
        "\u2022  EVERY RECEIPT: Expenses \u2192 new row \u2192 date, type, description, truck, paid to, amount, method, receipt Y/N.\n"
        "\u2022  EVERY PAYDAY: log the release in Expenses with type \u2018Payroll Payout\u2019 and the crew name in PAID TO. "
        "Payroll clears the balance automatically. Cash advances work the same way with \u2018Cash Advance\u2019.\n"
        "\u2022  MONTHLY: open the Dashboard, set MONTH, read the P&L, check \u2018Cash profit after loans\u2019. If it is positive "
        "you can service the loans and still keep money. Then check STILL TO PAY CREW and UNPAID INVOICES, and collect.",
        size=9)
rm_gap()

rm_heading("7 \u00b7 GROWING THE BUSINESS")
for head, txt in [
    ("New truck", "Setup \u00a74 \u2192 add a row (name, plate, capacity, status, driver). It appears in every dropdown, "
                  "on the dashboard truck table and in the charts \u2014 no formula edits."),
    ("New driver / helper", "Setup \u00a75 \u2192 add name + role. Leave SHARE % blank to inherit 10% / 5%. Payroll picks it up "
                            "for every month automatically."),
    ("New destination", "Setup \u00a73 \u2192 add one row: SOURCE, DESTINATION, RATE. Also add the destination to Setup \u00a78 "
                        "so it appears in the ranking. Trips resolves the rate by source + destination; if the lane is "
                        "one-way it still matches on destination alone."),
    ("New loan", "Setup \u00a77 \u2192 lender, truck, principal, rate, monthly amortisation, term, start date. Payments are "
                 "counted from the Expenses sheet the moment you log them as \u2018Loan Payment\u2019 with that lender in PAID TO."),
    ("New year", "Copy the last month block on the Payroll sheet and change the hidden month key in column R. "
                 "The dashboard month list already covers 24 months."),
]:
    rm_body(f"\u25aa  {head}  \u2014  {txt}", size=9)
rm_gap()

rm_heading("8 \u00b7 USING IT IN GOOGLE SHEETS")
rm_body("File \u2192 Import \u2192 Upload \u2192 select this file \u2192 Import location: \u2018Replace spreadsheet\u2019. "
        "Formulas, dropdowns, conditional colours and the dashboard all convert. Two things behave differently:")
rm_body("\u2022  Charts: Google Sheets re-draws them on import. If a series arrives empty, click the chart \u2192 "
        "three dots \u2192 Edit chart \u2192 re-pick the range on the Calc tab.\n"
        "\u2022  The month filter: if the dropdown does not survive the import, type the date in the MONTH cell "
        "(any day of the month works \u2014 it snaps to the 1st).", size=9)
rm_body("Tip: keep the Calc tab hidden (right-click \u2192 Hide sheet). Deleting it breaks the dashboard.", size=9, color=RED)
rm_gap()

rm_heading("9 \u00b7 IF SOMETHING LOOKS WRONG")
for problem, fix in [
    ("\u2018NO RATE\u2019 in the RATE / TON column", "That destination is not in Setup \u00a73, or the spelling differs. "
     "Pick the destination from the dropdown, or add the lane to the rate table."),
    ("\u2018CHECK TYPE\u2019 in the COST GROUP column", "The expense type is not in Setup \u00a76. Pick from the dropdown."),
    ("A crew member has no pay", "The name in the trip must match Setup \u00a75 exactly (use the dropdown)."),
    ("A trip shows no allowance", "Blank allowance uses the default in Setup \u00a72 \u2014 type 0 if no allowance was given."),
    ("A cost is missing from the P&L", "Check its COST GROUP: Transfer and Other are deliberately excluded from profit."),
    ("A crew member still shows a balance", "The payout or advance must carry that person's name in PAID TO/CREDITOR."),
]:
    rm_body(f"\u25aa  {problem}\n     \u2192  {fix}", size=9)
rm_gap()

rm_heading("10 \u00b7 IMPORTANT")
rm_body(f"The workbook opens with {len(TR)} sample trips and {len(EX)} sample expenses so you can see the dashboard working. "
        "Before you start for real, delete the sample rows (select the filled rows on Trips and on Expenses \u2192 "
        "right-click \u2192 Delete rows / Clear contents) \u2014 or keep one month as a reference and delete the rest.", size=9, bold=True, color=RED)
rm_body("Log every peso. A trucking business rarely fails because of the rate \u2014 it fails because the diesel, the "
        "toll and the 5% helper share were never written down.", size=9, italic=True)

# ======================================================================
#  defined names, sheet protection, save
# ======================================================================
NAMES = {
    "Reporting_Month":   "Dashboard!$C$4",
    "Truck_Filter":      "Dashboard!$F$4",
    "Gross_Fee_Month":   "Dashboard!$B$8",
    "Cash_Profit_Month": "Dashboard!$L$8",
    "Margin_Month":      "Dashboard!$B$12",
    "Crew_To_Pay":       "Dashboard!$H$12",
    "Loan_Balance":      "Dashboard!$J$12",
    "Unpaid_Invoices":   "Dashboard!$L$12",
}
for nm, ref in NAMES.items():
    try:
        wb.defined_names.add(DefinedName(nm, attr_text=ref))
    except Exception as exc:
        print("name skipped:", nm, exc)

# light protection on the engine sheet (no password: visible in the Review tab)
ws_calc.protection.sheet = True
ws_calc.protection.enable()

wb.calculation.fullCalcOnLoad = True
wb.properties.title = "J&Z Trucking Command Center"
wb.properties.subject = "Trip profitability, expenses, payroll and loan tracking (Philippines)"
wb.properties.creator = "J&Z Trucking"
wb.properties.keywords = "trucking, P&L, payroll, loans, Philippines, dashboard"

out = "../JZ_TEST.xlsx" if TEST_MODE else "../JZ_Trucking_Command_Center.xlsx"
wb.active = 0
wb.save(out)
print("saved", out)
