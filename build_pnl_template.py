from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = Workbook()
ws = wb.active
ws.title = "P&L"
ws.sheet_view.showGridLines = False

GREEN = "1B5E20"
GREEN_L = "2E7D32"
BLUE = "1565C0"
LGRAY = "F2F2F2"
NET_BG = "C8E6C9"
DGRAY = "B0B0B0"

title_font = Font(bold=True, size=15, color="1B5E20")
section_font = Font(bold=True, size=11, color="FFFFFF")
subtotal_font = Font(bold=True, size=11)
line_font = Font(size=11)
label_align = Alignment(horizontal="left", vertical="center")
num_align = Alignment(horizontal="right", vertical="center")
indent_align = Alignment(horizontal="left", vertical="center", indent=2)
thin = Side(style="thin", color=DGRAY)
top_border = Border(top=thin)
bot_border = Border(bottom=thin)
both_border = Border(top=thin, bottom=thin)

# --- Title block ---
ws["A1"] = "Profit & Loss Statement"
ws["A1"].font = title_font
ws["A2"] = "Company Name"
ws["A2"].font = Font(italic=True, size=11, color="555555")
ws["C2"] = "Period:"
ws["C2"].alignment = Alignment(horizontal="right")
ws["B3"] = "Currency:"
ws["B3"].alignment = Alignment(horizontal="right")
ws["C3"] = "USD"
ws["C3"].alignment = Alignment(horizontal="right")

# --- Column widths ---
ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 16

# --- Header row ---
for col, h in enumerate(["Account", "Amount", "% of Revenue"], 1):
    c = ws.cell(row=4, column=col, value=h)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=BLUE)
    c.alignment = label_align if col == 1 else num_align

def section(row, title):
    for col in range(1, 4):
        c = ws.cell(row=row, column=col)
        c.fill = PatternFill("solid", fgColor=GREEN_L)
        c.font = section_font
        c.border = both_border
    ws.cell(row=row, column=1, value=title).alignment = label_align

def rev_line(row, label):
    for col in range(1, 4):
        c = ws.cell(row=row, column=col)
        c.fill = PatternFill("solid", fgColor=LGRAY)
        c.border = both_border
    ws.cell(row=row, column=1, value=label).alignment = label_align
    ws.cell(row=row, column=2).alignment = num_align
    ws.cell(row=row, column=2).number_format = '#,##0.00'

def exp_line(row, label):
    ws.cell(row=row, column=1, value=label).alignment = indent_align
    ws.cell(row=row, column=1).font = line_font
    ws.cell(row=row, column=2).alignment = num_align
    ws.cell(row=row, column=2).number_format = '#,##0.00'
    ws.cell(row=row, column=2).border = bot_border

def subtotal(row, label, formula, highlight=False):
    fc = PatternFill("solid", fgColor=NET_BG) if highlight else None
    for col in range(1, 4):
        c = ws.cell(row=row, column=col)
        c.font = subtotal_font
        c.border = top_border
        if fc:
            c.fill = fc
    ws.cell(row=row, column=1, value=label).alignment = label_align
    sc = ws.cell(row=row, column=2)
    sc.value = formula
    sc.number_format = '#,##0.00'
    sc.alignment = num_align

def zero_fill(row, col, formula=None):
    c = ws.cell(row=row, column=col)
    c.number_format = '#,##0.00'
    c.alignment = num_align
    if formula:
        c.value = formula

R = 5
def rows_layout():
    # REVENUE
    section(5, "REVENUE")
    rev_line(6, "Sales Revenue")
    rev_line(7, "Other Revenue")
    subtotal(8, "Total Revenue", "=SUM(B6:B7)")
    # COGS
    section(9, "COST OF GOODS SOLD")
    exp_line(10, "Cost of Goods Sold")
    subtotal(11, "GROSS PROFIT", "=B8-B10")
    # OPEX
    section(12, "OPERATING EXPENSES")
    exp_line(13, "Salaries & Wages")
    exp_line(14, "Rent")
    exp_line(15, "Marketing")
    exp_line(16, "Utilities")
    exp_line(17, "Insurance")
    exp_line(18, "IT & Software")
    exp_line(19, "Office Supplies")
    exp_line(20, "Travel & Entertainment")
    exp_line(21, "Professional Fees")
    subtotal(22, "Total Operating Expenses", "=SUM(B13:B21)")
    subtotal(23, "EBITDA", "=B11-B22")
    exp_line(24, "Depreciation & Amortization")
    subtotal(25, "OPERATING INCOME (EBIT)", "=B23-B24")
    exp_line(26, "Interest Expense")
    subtotal(27, "Income Before Tax", "=B25-B26")
    exp_line(28, "Income Tax")
    subtotal(29, "NET INCOME", "=B27-B28", highlight=True)

rows_layout()

# % of Revenue column (C): formula = amount / total revenue, formatted as percent
for row in range(6, 30):
    if row in (5, 8, 9, 12):
        if row == 8:
            c = ws.cell(row=row, column=3)
            c.value = 1
            c.number_format = '0.0%'
            c.alignment = num_align
        continue
    c = ws.cell(row=row, column=3)
    c.value = "=IF($B$8=0,\"\",B{0}/$B$8)".format(row)
    c.number_format = '0.0%'
    c.alignment = num_align

for r in range(5, 30):
    for col in (1, 2, 3):
        ws.cell(row=r, column=col).font = Font(size=11)

ws.freeze_panes = "A5"

# Sample data (leave blank for template? fill with zeros so formulas evaluate)
sample = {6: 100000, 7: 5000, 10: 30000, 13: 25000, 14: 8000,
          15: 6000, 16: 1500, 17: 1200, 18: 2000, 19: 800,
          20: 2500, 21: 1000, 24: 3500, 26: 1200, 28: 20000}
for row, val in sample.items():
    c = ws.cell(row=row, column=2)
    c.value = val
    c.number_format = '#,##0.00'
    c.alignment = num_align

wb.save("PNL_Template.xlsx")
print("Saved PNL_Template.xlsx")