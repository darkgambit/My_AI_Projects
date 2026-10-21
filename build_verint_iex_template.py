"""
build_verint_iex_template.py
============================
Builds "Verint_IEX_Raw_Template.xlsx" - a production-ready WFM workbook that
ingests raw Verint Workforce Management and IEX TotalView interval exports,
aggregates them with formulas, reconciles the two sources and renders a
dashboard.

Design rules
------------
* Raw sheets stay pure data (no formulas) so any export can be pasted in.
* One canonical schema is shared by both sources -> reconciliation is 1:1.
* All maths lives on Calc_Data (defined names, SUMIFS / SUMPRODUCT).
* Nothing outside Python's standard library + openpyxl is required.

Run:  python build_verint_iex_template.py
"""

import random
from copy import copy
from datetime import date, time, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "Verint_IEX_Raw_Template.xlsx"

# --------------------------------------------------------------------------
# Palette / global styles
# --------------------------------------------------------------------------
NAVY = "0F2B46"
NAVY_L = "1D4E76"
TEAL = "0E7C86"
ACCENT = "1F6FB2"
LIGHT = "EAF2F8"
BAND = "F7FAFC"
GREEN = "1E7B4D"
GREEN_L = "E3F2E8"
AMBER = "B7791F"
AMBER_L = "FFF4E0"
RED = "B23B3B"
RED_L = "FBE4E4"
GRAY = "5B6B7A"
BORDER = "C9D6E2"
WHITE = "FFFFFF"

FONT = "Calibri"

THIN = Side(style="thin", color=BORDER)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOTTOM = Border(bottom=THIN)

PCT = "0.0%"
PCT2 = "0.00%"
INT = "#,##0"
DEC1 = "#,##0.0"
DEC2 = "#,##0.00"
DATE_FMT = "yyyy-mm-dd"
TIME_FMT = "hh:mm"

TODAY = date(2026, 10, 1)  # template default period start

# --------------------------------------------------------------------------
# Canonical raw schema (identical for Verint and IEX)
# --------------------------------------------------------------------------
COLUMNS = [
    # (letter, header, width, number_format, group)
    ("A", "Date", 12, DATE_FMT, "key"),
    ("B", "Interval Start", 11, TIME_FMT, "key"),
    ("C", "Interval End", 11, TIME_FMT, "key"),
    ("D", "Interval", 19, "@", "key"),
    ("E", "Site", 14, "@", "key"),
    ("F", "Queue / Skill", 22, "@", "key"),
    ("G", "Channel", 12, "@", "key"),
    ("H", "Offered", 10, INT, "volume"),
    ("I", "Handled", 10, INT, "volume"),
    ("J", "Abandoned", 11, INT, "volume"),
    ("K", "Short Abandoned", 15, INT, "volume"),
    ("L", "ASA (sec)", 10, DEC1, "speed"),
    ("M", "AHT (sec)", 10, DEC1, "speed"),
    ("N", "Talk (sec)", 11, INT, "time"),
    ("O", "Hold (sec)", 11, INT, "time"),
    ("P", "ACW (sec)", 11, INT, "time"),
    ("Q", "Staffed (agents)", 15, DEC1, "staffing"),
    ("R", "Required (agents)", 16, DEC1, "staffing"),
    ("S", "Service Level %", 14, PCT, "speed"),
    ("T", "Max Wait (sec)", 13, INT, "speed"),
    ("U", "Source File", 26, "@", "meta"),
    ("V", "Notes", 26, "@", "meta"),
]
FIRST_DATA_ROW = 6
CAPACITY = 10_000  # data rows per source
LAST_DATA_ROW = FIRST_DATA_ROW + CAPACITY - 1  # 10,005

FIELD_NAMES = {
    "Date": "DATE",
    "Interval Start": "TIME",
    "Interval End": "END",
    "Interval": "INT",
    "Site": "SITE",
    "Queue / Skill": "QUEUE",
    "Channel": "CHAN",
    "Offered": "OFFERED",
    "Handled": "HANDLED",
    "Abandoned": "ABAND",
    "Short Abandoned": "SHORT",
    "ASA (sec)": "ASA",
    "AHT (sec)": "AHT",
    "Talk (sec)": "TALK",
    "Hold (sec)": "HOLD",
    "ACW (sec)": "ACW",
    "Staffed (agents)": "STAFFED",
    "Required (agents)": "REQ",
    "Service Level %": "SL",
    "Max Wait (sec)": "WAIT",
    "Source File": "SRC",
    "Notes": "NOTES",
}

QUEUES = [
    "CS_Voice_General",
    "CS_Voice_Billing",
    "Tech_Support",
    "Retention",
    "Sales_Inbound",
    "Chat_Support",
    "Backoffice_Email",
    "Escalations",
    "VIP_Line",
    "Overflow",
]
SITES = ["Manila", "Cebu", "Davao", "Bacolod"]
CHANNELS = ["Voice", "Chat", "Email", "Social", "Backoffice"]


# --------------------------------------------------------------------------
# Small style helpers
# --------------------------------------------------------------------------
def fill(color):
    return PatternFill("solid", fgColor=color)


def title_block(ws, text, subtitle, last_col="J"):
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = text
    c.font = Font(name=FONT, size=15, bold=True, color=WHITE)
    c.fill = fill(NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value = subtitle
    c.font = Font(name=FONT, size=9, italic=True, color=GRAY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18


def section_bar(ws, row, text, first_col="A", last_col="J"):
    ws.merge_cells(f"{first_col}{row}:{last_col}{row}")
    c = ws[f"{first_col}{row}"]
    c.value = text
    c.font = Font(name=FONT, size=10, bold=True, color=NAVY)
    c.fill = fill(LIGHT)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(ws[f"{first_col}{row}"].column,
                     ws[f"{last_col}{row}"].column + 1):
        ws.cell(row=row, column=col).border = BOTTOM
    ws.row_dimensions[row].height = 20


def header_cell(cell, bg=NAVY_L, color=WHITE, size=9, align="center", wrap=True):
    cell.font = Font(name=FONT, size=size, bold=True, color=color)
    cell.fill = fill(bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    cell.border = BOX


# --------------------------------------------------------------------------
# Sampling helper (deterministic demo data)
# --------------------------------------------------------------------------
def build_sample_rows(source, shift=0, seed=7):
    rng = random.Random(seed)
    rows = []
    start = time(8, 0)
    shape = [0.55, 0.75, 0.95, 1.0, 1.0, 0.95, 0.9, 0.85, 0.8, 0.85, 0.9, 0.7]
    for day in (date(2026, 10, 1), date(2026, 10, 2)):
        for qi, queue in enumerate(QUEUES[:4]):
            base = [42, 26, 18, 12][qi]
            for i in range(12):
                offered = max(3, int(round(base * shape[i] * rng.uniform(0.9, 1.1))))
                handled = max(0, int(round(offered * rng.uniform(0.85, 0.94))))
                abandoned = offered - handled
                aht = rng.uniform(300, 470)
                talk = int(round(handled * aht * 0.66))
                hold = int(round(handled * aht * 0.05))
                acw = int(round(handled * aht * 0.29))
                staffed = max(1, round(offered * aht / 1800 * 1.08, 1))
                required = round(staffed * rng.uniform(0.95, 1.15), 1)
                asa = rng.choice([8, 12, 18, 24, 31, 44, 58])
                sl = max(0.42, min(0.98, 1 - abandoned / offered * 1.6))
                interval_end = (timedelta(minutes=30) * (i + 1))
                end_dt = (timedelta(hours=8) + interval_end)
                rows.append({
                    "date": day,
                    "start": (timedelta(hours=start.hour, minutes=start.minute)
                              + timedelta(minutes=30 * i)),
                    "end": end_dt,
                    "label": f"{fmt_time(timedelta(hours=8) + timedelta(minutes=30*i))} - "
                             f"{fmt_time(end_dt)}",
                    "site": "Manila",
                    "queue": queue,
                    "channel": "Voice",
                    "offered": offered,
                    "handled": handled,
                    "abandoned": abandoned,
                    "short": int(round(abandoned * 0.35)),
                    "asa": asa,
                    "aht": round(aht, 1),
                    "talk": talk,
                    "hold": hold,
                    "acw": acw,
                    "staffed": staffed,
                    "required": required,
                    "sl": round(sl, 3),
                    "wait": rng.choice([120, 180, 240, 310, 420, 600, 780]),
                })
    # Verint is the reference pull; IEX repeats it with two deliberate discrepancies
    if source == "IEX":
        day1 = rows[20]
        day1["offered"] += 9
        day1["abandoned"] += 9
        day2 = rows[63]
        day2["offered"] = max(1, day2["offered"] - 40)
        day2["abandoned"] = max(0, day2["abandoned"] - 40)
        day2["handled"] = day2["offered"] - day2["abandoned"]
    for row in rows:
        row["abandoned"] = max(0, row["offered"] - row["handled"])
        row["short"] = min(row["abandoned"], row["short"])
        row["sl"] = round(max(0.3, min(0.99, row["sl"])), 3)
    return rows


def fmt_time(td_or_time):
    total_minutes = int(td_or_time.total_seconds() // 60) if isinstance(td_or_time, timedelta) \
        else td_or_time.hour * 60 + td_or_time.minute
    h, m = divmod(total_minutes, 60)
    suffix = "AM" if h < 12 else "PM"
    h12 = h % 12 or 12
    return f"{h12}:{m:02d} {suffix}"


# --------------------------------------------------------------------------
# Turn timedelta into an Excel-ready time value (fraction of a day)
# --------------------------------------------------------------------------
def as_time(td):
    return td.total_seconds() / 86400.0


# --------------------------------------------------------------------------
# Raw sheet builder
# --------------------------------------------------------------------------
def build_raw_sheet(wb, name, source_label, tab_color, sample_rows):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = tab_color
    ws.sheet_view.showGridLines = False

    last_col = COLUMNS[-1][0]
    title_block(
        ws,
        f"{source_label} RAW DATA",
        "Paste the raw interval export below the header row. Do not edit column order, "
        "insert formulas or merge cells - all calculations happen on Calc_Data.",
        last_col=last_col,
    )

    # source badge / instructions
    ws["A3"] = "Source"
    ws["A3"].font = Font(name=FONT, size=9, bold=True, color=GRAY)
    ws["B3"] = source_label
    ws["B3"].font = Font(name=FONT, size=9, bold=True, color=TEAL)

    ws["C3"] = "Rows 6-10,005 available  |  Paste as values  |  Date = date, Interval = hh:mm"
    ws["C3"].font = Font(name=FONT, size=9, italic=True, color=GRAY)
    ws.merge_cells("C3:H3")

    ws["J3"] = "Sample rows are included - clear rows 6:101 before live use"
    ws["J3"].font = Font(name=FONT, size=9, italic=True, color=AMBER)
    ws.merge_cells("J3:V3")

    # header row
    hdr = 5
    ws.row_dimensions[hdr].height = 34
    group_colors = {
        "key": NAVY_L,
        "volume": TEAL,
        "speed": "3E7CA6",
        "time": "3E7CA6",
        "staffing": "6B5B95",
        "meta": GRAY,
    }
    for letter, header, width, numfmt, group in COLUMNS:
        ws.column_dimensions[letter].width = width
        c = ws[f"{letter}{hdr}"]
        c.value = header
        header_cell(c, bg=group_colors[group])
        ws[f"{letter}{hdr}"].alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    # column level number formats so pasted values land formatted
    for letter, header, width, numfmt, group in COLUMNS:
        ws.column_dimensions[letter].number_format = numfmt

    # light formatting for the paste zone (kept to 200 rows for file size)
    for r in range(FIRST_DATA_ROW, FIRST_DATA_ROW + 200):
        for letter, *_ in COLUMNS:
            cell = ws[f"{letter}{r}"]
            cell.border = BOTTOM
            if r % 2 == 0:
                cell.fill = fill(BAND)

    # write sample rows
    for i, row in enumerate(sample_rows):
        r = FIRST_DATA_ROW + i
        values = {
            "A": row["date"],
            "B": as_time(row["start"]),
            "C": as_time(row["end"]),
            "D": row["label"],
            "E": row["site"],
            "F": row["queue"],
            "G": row["channel"],
            "H": row["offered"],
            "I": row["handled"],
            "J": row["abandoned"],
            "K": row["short"],
            "L": row["asa"],
            "M": row["aht"],
            "N": row["talk"],
            "O": row["hold"],
            "P": row["acw"],
            "Q": row["staffed"],
            "R": row["required"],
            "S": row["sl"],
            "T": row["wait"],
            "U": f"SAMPLE_{source_label}_2026-10.xlsx",
            "V": "Sample row - delete before live use",
        }
        for letter, value in values.items():
            ws[f"{letter}{r}"] = value

    # keep sample rows visually distinct from the empty paste zone
    for r in range(FIRST_DATA_ROW, FIRST_DATA_ROW + len(sample_rows)):
        for letter, *_ in COLUMNS:
            ws[f"{letter}{r}"].fill = fill(AMBER_L)

    ws.freeze_panes = f"A{FIRST_DATA_ROW}"
    ws.auto_filter.ref = f"A{hdr}:{last_col}{LAST_DATA_ROW}"

    # validations (warning style so real exports are never blocked)
    dv_site = DataValidation(type="list", formula1="LIST_SITES", allow_blank=True,
                             showErrorMessage=True, errorStyle="warning",
                             errorTitle="Unknown site",
                             error="This site is not in Lists!A. Continue only if the value is intentional.")
    dv_queue = DataValidation(type="list", formula1="LIST_QUEUES", allow_blank=True,
                              showErrorMessage=True, errorStyle="warning",
                              errorTitle="Unknown queue / skill",
                              error="This queue is not in Lists!B. Continue only if the value is intentional.")
    dv_chan = DataValidation(type="list", formula1="LIST_CHANNELS", allow_blank=True,
                             showErrorMessage=True, errorStyle="warning",
                             errorTitle="Unknown channel",
                             error="This channel is not in Lists!C. Continue only if the value is intentional.")
    for dv, col in ((dv_site, "E"), (dv_queue, "F"), (dv_chan, "G")):
        ws.add_data_validation(dv)
        dv.add(f"{col}{FIRST_DATA_ROW}:{col}{LAST_DATA_ROW}")

    return ws


# --------------------------------------------------------------------------
# Setup sheet
# --------------------------------------------------------------------------
def build_setup(wb):
    ws = wb.create_sheet("Setup")
    ws.sheet_properties.tabColor = ACCENT
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 52
    for col in "DEFGH":
        ws.column_dimensions[col].width = 3

    title_block(ws, "SETUP - PERIOD, FILTERS & SETTINGS",
                "Every report in this workbook reads from these cells. Change a filter "
                "here and all outputs refresh.", last_col="C")

    section_bar(ws, 4, "REPORTING PERIOD", last_col="C")

    rows = [
        ("Period Start", TODAY, DATE_FMT, "First day of the reporting month."),
        ("Period End", "=EOMONTH(B5,0)", DATE_FMT, "Auto-calculated - do not overwrite."),
        ("Interval Minutes", 30, "0", "Length of one interval: 15, 30 or 60."),
    ]
    r = 5
    for label, value, fmt, note in rows:
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
        c = ws[f"B{r}"]
        c.value = value
        c.number_format = fmt
        c.font = Font(name=FONT, size=11, color=NAVY)
        c.fill = fill(LIGHT)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = Font(name=FONT, size=9, italic=True, color=GRAY)
        r += 1

    section_bar(ws, 9, "FILTERS (applies to Calc_Data, Dashboard & Reconciliation)", last_col="C")
    filters = [
        ("Site", "ALL", "Use ALL for the whole account."),
        ("Queue / Skill", "ALL", "Use ALL or pick one queue from Lists."),
        ("Channel", "ALL", "Use ALL or pick one channel from Lists."),
        ("Report Source", "COMBINED", "VERINT, IEX or COMBINED on the dashboard."),
    ]
    r = 10
    for label, value, note in filters:
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
        c = ws[f"B{r}"]
        c.value = value
        c.font = Font(name=FONT, size=11, bold=True, color=TEAL)
        c.fill = fill(LIGHT)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = Font(name=FONT, size=9, italic=True, color=GRAY)
        r += 1

    dv_site = DataValidation(type="list", formula1="LIST_SITES_FILTER", allow_blank=False)
    dv_queue = DataValidation(type="list", formula1="LIST_QUEUES_FILTER", allow_blank=False)
    dv_chan = DataValidation(type="list", formula1="LIST_CHANNELS_FILTER", allow_blank=False)
    dv_src = DataValidation(type="list", formula1='"VERINT,IEX,COMBINED"', allow_blank=False)
    dv_int = DataValidation(type="list", formula1='"15,30,60"', allow_blank=False)
    for dv, ref in ((dv_site, "B10"), (dv_queue, "B11"), (dv_chan, "B12"),
                    (dv_src, "B13"), (dv_int, "B7")):
        ws.add_data_validation(dv)
        dv.add(ref)

    section_bar(ws, 15, "WORKBOOK INFO", last_col="C")
    info = [
        ("Account / Site Name", "Contact Center Operations", None, "Shown on the dashboard."),
        ("Time Zone", "GMT+8", None, "Label only - keep both sources in one time zone."),
        ("Prepared By", "WFM Team", None, "Owner of this workbook."),
        ("Last Updated", "=TODAY()", DATE_FMT, "Auto-stamped on open."),
        ("Reconciliation Tolerance", 0.02, PCT, "Variance above this is flagged REVIEW."),
    ]
    r = 16
    for label, value, fmt, note in info:
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
        c = ws[f"B{r}"]
        c.value = value
        if fmt:
            c.number_format = fmt
        c.font = Font(name=FONT, size=11, color=NAVY)
        c.fill = fill(LIGHT)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = Font(name=FONT, size=9, italic=True, color=GRAY)
        r += 1

    notes = [
        "HOW TO USE",
        "1. Set Period Start above, then paste the Verint export into Verint_Raw and the IEX export into IEX_Raw.",
        "2. Both raw sheets share one schema - if your export order differs, rearrange the columns before pasting.",
        "3. Clear the yellow sample rows (6:101) in both raw sheets once you are ready for live data.",
        "4. Check DQ_Checks for blanks, duplicates and impossible values before reading the dashboard.",
        "5. Use Reconciliation to prove Verint and IEX agree interval by interval.",
        "6. Capacity is 10,000 rows per source. For a larger month, extend the raw ranges and update the defined names.",
    ]
    r = 22
    for i, line in enumerate(notes):
        ws.merge_cells(f"A{r}:C{r}")
        c = ws[f"A{r}"]
        c.value = line
        c.font = Font(name=FONT, size=10 if i else 10, bold=(i == 0), color=NAVY if i == 0 else GRAY)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[r].height = 18 if i == 0 else 30
        r += 1

    return ws


# --------------------------------------------------------------------------
# Lists sheet
# --------------------------------------------------------------------------
def build_lists(wb):
    ws = wb.create_sheet("Lists")
    ws.sheet_properties.tabColor = GRAY
    ws.sheet_view.showGridLines = False
    title_block(ws, "REFERENCE LISTS", "Dropdown sources and queue master. Edit freely - "
                "filter lists always keep ALL in row 1.", last_col="E")

    headers = ["Sites", "Queues / Skills", "Channels", "Interval Minutes", "Sources"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=4, column=i, value=h)
        header_cell(c, size=10)
        ws.column_dimensions[get_column_letter(i)].width = 24

    for i, site in enumerate(SITES, start=5):
        ws.cell(row=i, column=1, value=site).border = BOTTOM
    for i, queue in enumerate(QUEUES, start=5):
        ws.cell(row=i, column=2, value=queue).border = BOTTOM
    for i, chan in enumerate(CHANNELS, start=5):
        ws.cell(row=i, column=3, value=chan).border = BOTTOM
    for i, minutes in enumerate([15, 30, 60], start=5):
        ws.cell(row=i, column=4, value=minutes).border = BOTTOM
    for i, src in enumerate(["VERINT", "IEX", "COMBINED"], start=5):
        ws.cell(row=i, column=5, value=src).border = BOTTOM

    # filter lists (with ALL) live in G:J so they never disturb the master lists
    ws["G4"] = "Site filter"
    ws["H4"] = "Queue filter"
    ws["I4"] = "Channel filter"
    for cell in ("G4", "H4", "I4"):
        header_cell(ws[cell], size=10)
    ws["G5"] = "ALL"
    ws["H5"] = "ALL"
    ws["I5"] = "ALL"
    for i, site in enumerate(SITES, start=6):
        ws[f"G{i}"] = site
    for i, queue in enumerate(QUEUES, start=6):
        ws[f"H{i}"] = queue
    for i, chan in enumerate(CHANNELS, start=6):
        ws[f"I{i}"] = chan
    for col in "GHI":
        ws.column_dimensions[col].width = 22

    return ws


# --------------------------------------------------------------------------
# Defined names
# --------------------------------------------------------------------------
def add_defined_names(wb):
    name_map = {}
    for sheet, prefix in (("Verint_Raw", "VER"), ("IEX_Raw", "IEX")):
        for letter, header, *_ in COLUMNS:
            suffix = FIELD_NAMES[header]
            name = f"{prefix}_{suffix}"
            ref = f"{sheet}!${letter}${FIRST_DATA_ROW}:${letter}${LAST_DATA_ROW}"
            wb.defined_names[name] = DefinedName(name, attr_text=ref)
            name_map[name] = ref

    extra = {
        "LIST_SITES": f"Lists!$A$5:$A${4 + len(SITES)}",
        "LIST_QUEUES": f"Lists!$B$5:$B${4 + len(QUEUES)}",
        "LIST_CHANNELS": f"Lists!$C$5:$C${4 + len(CHANNELS)}",
        "LIST_SITES_FILTER": f"Lists!$G$5:$G${5 + len(SITES)}",
        "LIST_QUEUES_FILTER": f"Lists!$H$5:$H${5 + len(QUEUES)}",
        "LIST_CHANNELS_FILTER": f"Lists!$I$5:$I${5 + len(CHANNELS)}",
    }
    for name, ref in extra.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    return name_map


# --------------------------------------------------------------------------
# Calc_Data
# --------------------------------------------------------------------------
KPI_START = 14
DAILY_HEADER = 31
DAILY_START = 32
DAILY_ROWS = 31
PROFILE_HEADER = 65
PROFILE_START = 66
PROFILE_ROWS = 48
QUEUE_HEADER = 117
QUEUE_START = 118

KPI_DEFS = [
    # (label, number format, combined strategy)
    ("Contacts Offered", INT),
    ("Contacts Handled", INT),
    ("Contacts Abandoned", INT),
    ("Abandon Rate", PCT),
    ("Service Level", PCT),
    ("Avg Speed of Answer (sec)", DEC1),
    ("Avg Handle Time (sec)", DEC1),
    ("Staffed Agent-Hours", DEC1),
    ("Required Agent-Hours", DEC1),
    ("Staffing Gap (Agent-Hours)", DEC1),
    ("Staffing Gap %", PCT),
    ("Occupancy", PCT),
    ("Intervals Captured", INT),
    ("Avg Staffed Agents / Interval", DEC1),
    ("Max Interval Wait (sec)", INT),
]


def crit_month(prefix):
    return (f'{prefix}_DATE,">="&$B$3,{prefix}_DATE,"<="&$B$4,'
            f'{prefix}_SITE,$D$3,{prefix}_QUEUE,$D$4,{prefix}_CHAN,$D$5')


def crit_date(prefix, cell):
    return (f'{prefix}_DATE,">="&{cell},{prefix}_DATE,"<="&{cell},'
            f'{prefix}_SITE,$D$3,{prefix}_QUEUE,$D$4,{prefix}_CHAN,$D$5')


def or_cond(prefix, field, cell):
    """Filter = cell OR cell is ALL, expressed arithmetically (no array > 0)."""
    match = f'({prefix}_{field}={cell})'
    is_all = f'({cell}="ALL")'
    return f'({is_all}+{match}-{is_all}*{match})'


def sp(prefix, value, weight, date_cell=None):
    """SUMPRODUCT numerator that honours ALL filters and wildcards-free slots."""
    d = (f'(({prefix}_DATE>={date_cell})*({prefix}_DATE<={date_cell}))' if date_cell
         else f'(({prefix}_DATE>=$B$3)*({prefix}_DATE<=$B$4))')
    site = f'((($B$6="ALL")+({prefix}_SITE=$B$6))>0)'
    queue = f'((($B$7="ALL")+({prefix}_QUEUE=$B$7))>0)'
    chan = f'((($B$8="ALL")+({prefix}_CHAN=$B$8))>0)'
    return f'SUMPRODUCT(({d})*({site})*({queue})*({chan})*{value}*{weight})'


def sumif(metric, crit):
    return f"SUMIFS({metric},{crit})"


def build_calc(wb):
    ws = wb.create_sheet("Calc_Data")
    ws.sheet_properties.tabColor = "8A93A5"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 36
    for col in "BCDEFGHIJKLMN":
        ws.column_dimensions[col].width = 13
    ws.column_dimensions["N"].width = 14

    title_block(ws, "CALCULATION ENGINE",
                "Do not edit. Every figure is derived from the raw sheets through filtered "
                "SUMIFS / SUMPRODUCT formulas.", last_col="N")

    # ---- control block ----
    ws["A3"] = "Control block"
    ws["A3"].font = Font(name=FONT, size=9, bold=True, color=GRAY)
    controls = [
        ("Period Start", "=Setup!$B$5", DATE_FMT),
        ("Period End", "=Setup!$B$6", DATE_FMT),
        ("Interval Minutes", "=Setup!$B$7", "0"),
        ("Site Filter", "=Setup!$B$10", "@"),
        ("Queue Filter", "=Setup!$B$11", "@"),
        ("Channel Filter", "=Setup!$B$12", "@"),
        ("Reconciliation Tolerance", "=Setup!$B$20", PCT),
        ("Report Source", "=Setup!$B$13", "@"),
    ]
    r = 3
    for label, formula, fmt in controls:
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = Font(name=FONT, size=9, color=GRAY)
        c = ws[f"B{r}"]
        c.value = formula
        c.number_format = fmt
        c.font = Font(name=FONT, size=9, bold=True, color=NAVY)
        c.alignment = Alignment(horizontal="center")
        r += 1

    # wildcard versions of the filters (SUMIFS needs "*" for ALL)
    ws["D3"] = '=IF($B$6="ALL","*",$B$6)'
    ws["D4"] = '=IF($B$7="ALL","*",$B$7)'
    ws["D5"] = '=IF($B$8="ALL","*",$B$8)'
    for cell in ("D3", "D4", "D5"):
        ws[cell].font = Font(name=FONT, size=9, color=GRAY)
        ws[cell].alignment = Alignment(horizontal="center")
    ws["E3"] = "<- SUMIFS criteria"
    ws["E3"].font = Font(name=FONT, size=8, italic=True, color=GRAY)

    # ---- KPI block ----
    section_bar(ws, KPI_START - 2, "1. PERIOD TOTALS BY SOURCE", last_col="E")
    ws[f"A{KPI_START-1}"] = "Metric"
    ws[f"C{KPI_START-1}"] = "VERINT"
    ws[f"D{KPI_START-1}"] = "IEX"
    ws[f"E{KPI_START-1}"] = "COMBINED"
    for cell in (f"A{KPI_START-1}", f"C{KPI_START-1}", f"D{KPI_START-1}", f"E{KPI_START-1}"):
        header_cell(ws[cell], bg=NAVY_L, size=9)
    ws[f"A{KPI_START-1}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    row_of = {label: KPI_START + i for i, (label, _) in enumerate(KPI_DEFS)}
    r_off = row_of["Contacts Offered"]
    r_hand = row_of["Contacts Handled"]
    r_aband = row_of["Contacts Abandoned"]
    r_abpct = row_of["Abandon Rate"]
    r_sl = row_of["Service Level"]
    r_asa = row_of["Avg Speed of Answer (sec)"]
    r_aht = row_of["Avg Handle Time (sec)"]
    r_staff = row_of["Staffed Agent-Hours"]
    r_req = row_of["Required Agent-Hours"]
    r_gap = row_of["Staffing Gap (Agent-Hours)"]
    r_gappct = row_of["Staffing Gap %"]
    r_occ = row_of["Occupancy"]
    r_int = row_of["Intervals Captured"]
    r_avgstaff = row_of["Avg Staffed Agents / Interval"]
    r_wait = row_of["Max Interval Wait (sec)"]

    for i, (label, fmt) in enumerate(KPI_DEFS):
        r = KPI_START + i
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = Font(name=FONT, size=9, color=NAVY)
        ws[f"A{r}"].border = BOTTOM
        for col in ("C", "D", "E"):
            c = ws[f"{col}{r}"]
            c.number_format = fmt
            c.font = Font(name=FONT, size=9, color="1A1A1A")
            c.alignment = Alignment(horizontal="right")
            c.border = BOTTOM

    v_formulas = {}
    i_formulas = {}

    for prefix, bucket in (("VER", v_formulas), ("IEX", i_formulas)):
        c = crit_month(prefix)
        bucket["offered"] = sumif(f"{prefix}_OFFERED", c)
        bucket["handled"] = sumif(f"{prefix}_HANDLED", c)
        bucket["abandoned"] = sumif(f"{prefix}_ABAND", c)
        bucket["sl_num"] = sp(prefix, f"{prefix}_SL", f"{prefix}_OFFERED")
        bucket["asa_num"] = sp(prefix, f"{prefix}_ASA", f"{prefix}_HANDLED")
        bucket["talk"] = sumif(f"{prefix}_TALK", c)
        bucket["hold"] = sumif(f"{prefix}_HOLD", c)
        bucket["acw"] = sumif(f"{prefix}_ACW", c)
        bucket["staffed"] = sumif(f"{prefix}_STAFFED", c)
        bucket["req"] = sumif(f"{prefix}_REQ", c)
        bucket["count"] = (f'SUMPRODUCT(({prefix}_DATE>=$B$3)*({prefix}_DATE<=$B$4)*'
                           f'{or_cond(prefix, "SITE", "$B$6")}*'
                           f'{or_cond(prefix, "QUEUE", "$B$7")}*'
                           f'{or_cond(prefix, "CHAN", "$B$8")})')
        bucket["waitmax"] = (f'SUMPRODUCT(MAX(({prefix}_DATE>=$B$3)*({prefix}_DATE<=$B$4)*'
                             f'{or_cond(prefix, "SITE", "$B$6")}*'
                             f'{or_cond(prefix, "QUEUE", "$B$7")}*'
                             f'{or_cond(prefix, "CHAN", "$B$8")}*{prefix}_WAIT))')

    for prefix, bucket, col in (("VER", v_formulas, "C"), ("IEX", i_formulas, "D")):
        ws[f"{col}{r_off}"] = f"={bucket['offered']}"
        ws[f"{col}{r_hand}"] = f"={bucket['handled']}"
        ws[f"{col}{r_aband}"] = f"={bucket['abandoned']}"
        ws[f"{col}{r_abpct}"] = f"=IFERROR({col}{r_aband}/{col}{r_off},0)"
        ws[f"{col}{r_sl}"] = f"=IFERROR({bucket['sl_num']}/{col}{r_off},0)"
        ws[f"{col}{r_asa}"] = f"=IFERROR({bucket['asa_num']}/{col}{r_hand},0)"
        ws[f"{col}{r_aht}"] = (
            f"=IFERROR(({bucket['talk']}+{bucket['hold']}+{bucket['acw']})/{col}{r_hand},0)"
        )
        ws[f"{col}{r_staff}"] = f"={bucket['staffed']}*$B$5/60"
        ws[f"{col}{r_req}"] = f"={bucket['req']}*$B$5/60"
        ws[f"{col}{r_gap}"] = f"={col}{r_req}-{col}{r_staff}"
        ws[f"{col}{r_gappct}"] = f"=IFERROR({col}{r_gap}/{col}{r_req},0)"
        ws[f"{col}{r_occ}"] = (
            f"=IFERROR((({bucket['talk']}+{bucket['hold']}+{bucket['acw']})/3600)/{col}{r_staff},0)"
        )
        ws[f"{col}{r_int}"] = f"={bucket['count']}"
        ws[f"{col}{r_avgstaff}"] = f"=IFERROR({bucket['staffed']}/{col}{r_int},0)"
        ws[f"{col}{r_wait}"] = f"=IFERROR({bucket['waitmax']},0)"

    # COMBINED column
    additive = {r_off, r_hand, r_aband, r_staff, r_req, r_int}
    for i, (label, fmt) in enumerate(KPI_DEFS):
        r = KPI_START + i
        if r in additive:
            ws[f"E{r}"] = f"=C{r}+D{r}"
        elif r == r_abpct:
            ws[f"E{r}"] = f"=IFERROR(E{r_aband}/E{r_off},0)"
        elif r == r_sl:
            ws[f"E{r}"] = f"=IFERROR((C{r}*C{r_off}+D{r}*D{r_off})/E{r_off},0)"
        elif r == r_asa:
            ws[f"E{r}"] = f"=IFERROR((C{r}*C{r_hand}+D{r}*D{r_hand})/E{r_hand},0)"
        elif r == r_aht:
            ws[f"E{r}"] = f"=IFERROR((C{r}*C{r_hand}+D{r}*D{r_hand})/E{r_hand},0)"
        elif r == r_gap:
            ws[f"E{r}"] = f"=E{r_req}-E{r_staff}"
        elif r == r_gappct:
            ws[f"E{r}"] = f"=IFERROR(E{r_gap}/E{r_req},0)"
        elif r == r_occ:
            ws[f"E{r}"] = f"=IFERROR((C{r}*C{r_staff}+D{r}*D{r_staff})/E{r_staff},0)"
        elif r == r_avgstaff:
            ws[f"E{r}"] = f"=IFERROR((C{r}*C{r_int}+D{r}*D{r_int})/E{r_int},0)"
        elif r == r_wait:
            ws[f"E{r}"] = f"=MAX(C{r},D{r})"
        else:
            ws[f"E{r}"] = f"=C{r}+D{r}"

    # ---- daily block ----
    section_bar(ws, DAILY_HEADER - 2, "2. DAILY TREND (period dates)", last_col="N")
    daily_cols = [
        ("Date", 12, DATE_FMT, False),
        ("Day", 8, "@", False),
        ("VERINT Offered", 11, INT, True),
        ("IEX Offered", 11, INT, True),
        ("VERINT Handled", 11, INT, True),
        ("IEX Handled", 11, INT, True),
        ("VERINT Abandoned", 12, INT, True),
        ("IEX Abandoned", 12, INT, True),
        ("VERINT Aband %", 11, PCT, True),
        ("IEX Aband %", 11, PCT, True),
        ("VERINT SL %", 11, PCT, True),
        ("IEX SL %", 11, PCT, True),
        ("VERINT AHT", 11, DEC1, True),
        ("IEX AHT", 11, DEC1, True),
    ]
    for i, (label, width, fmt, _) in enumerate(daily_cols, start=1):
        col = get_column_letter(i)
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, width)
        c = ws.cell(row=DAILY_HEADER, column=i, value=label)
        header_cell(c, bg=NAVY_L, size=8)

    for i in range(DAILY_ROWS):
        r = DAILY_START + i
        if i == 0:
            ws[f"A{r}"] = '=IF($B$3>$B$4,"",$B$3)'
        else:
            ws[f"A{r}"] = f'=IF(OR(A{r-1}="",A{r-1}+1>$B$4),"",A{r-1}+1)'
        ws[f"B{r}"] = f'=IF(A{r}="","",TEXT(A{r},"ddd"))'
        for prefix, c1, c2, c3, c4, c5, c6 in (
            ("VER", "C", "E", "G", "I", "K", "M"),
            ("IEX", "D", "F", "H", "J", "L", "N"),
        ):
            crit = crit_date(prefix, f"$A{r}")
            ws[f"{c1}{r}"] = f'=IF($A{r}="","",{sumif(f"{prefix}_OFFERED", crit)})'
            ws[f"{c2}{r}"] = f'=IF($A{r}="","",{sumif(f"{prefix}_HANDLED", crit)})'
            ws[f"{c3}{r}"] = f'=IF($A{r}="","",{sumif(f"{prefix}_ABAND", crit)})'
            ws[f"{c4}{r}"] = f'=IF({c1}{r}="","",IFERROR({c3}{r}/{c1}{r},0))'
            ws[f"{c5}{r}"] = (
                f'=IF({c1}{r}="","",IFERROR({sp(prefix, f"{prefix}_SL", f"{prefix}_OFFERED", f"$A{r}")}'
                f'/{c1}{r},0))'
            )
            ws[f"{c6}{r}"] = (
                f'=IF({c1}{r}="","",IFERROR(({sumif(f"{prefix}_TALK", crit)}'
                f'+{sumif(f"{prefix}_HOLD", crit)}+{sumif(f"{prefix}_ACW", crit)})/{c2}{r},0))'
            )
        for i2, (_, _, fmt, _) in enumerate(daily_cols, start=1):
            cell = ws.cell(row=r, column=i2)
            cell.number_format = fmt
            cell.font = Font(name=FONT, size=9, color="1A1A1A")
            cell.border = BOTTOM
        ws[f"A{r}"].alignment = Alignment(horizontal="center")
        ws[f"B{r}"].alignment = Alignment(horizontal="center")

    # ---- interval profile ----
    section_bar(ws, PROFILE_HEADER - 2, "3. INTERVAL PROFILE (average day shape)", last_col="N")
    profile_cols = [
        ("Interval Start", 12, TIME_FMT),
        ("Interval", 20, "@"),
        ("VERINT Offered", 11, DEC1),
        ("IEX Offered", 11, DEC1),
        ("VERINT Avg Staffed", 12, DEC1),
        ("IEX Avg Staffed", 12, DEC1),
        ("VERINT Avg Required", 12, DEC1),
        ("IEX Avg Required", 12, DEC1),
    ]
    for i, (label, width, fmt) in enumerate(profile_cols, start=1):
        c = ws.cell(row=PROFILE_HEADER, column=i, value=label)
        header_cell(c, bg=NAVY_L, size=8)
        ws.column_dimensions[get_column_letter(i)].width = max(
            ws.column_dimensions[get_column_letter(i)].width or 0, width
        )

    for i in range(PROFILE_ROWS):
        r = PROFILE_START + i
        if i == 0:
            ws[f"A{r}"] = "=0"
        else:
            ws[f"A{r}"] = f"=A{r-1}+$B$5/1440"
        ws[f"B{r}"] = f'=TEXT(A{r},"h:mm AM/PM")&" - "&TEXT(A{r}+$B$5/1440,"h:mm AM/PM")'
        for prefix, col_off, col_staff, col_req in (
            ("VER", "C", "E", "G"), ("IEX", "D", "F", "H")
        ):
            base = crit_month(prefix)
            time_crit = (f'{prefix}_TIME,">="&ROUND($A{r},6),'
                         f'{prefix}_TIME,"<"&ROUND($A{r}+$B$5/1440,6)')
            ws[f"{col_off}{r}"] = f'=IFERROR({sumif(f"{prefix}_OFFERED", base + "," + time_crit)},0)'
            for col, metric in ((col_staff, "STAFFED"), (col_req, "REQ")):
                ws[f"{col}{r}"] = (
                    f'=IFERROR({sumif(f"{prefix}_{metric}", base + "," + time_crit)}'
                    f'/($B$4-$B$3+1),0)'
                )
        for i2, (_, _, fmt) in enumerate(profile_cols, start=1):
            cell = ws.cell(row=r, column=i2)
            cell.number_format = fmt
            cell.font = Font(name=FONT, size=9, color="1A1A1A")
            cell.border = BOTTOM
        ws[f"A{r}"].alignment = Alignment(horizontal="center")

    # ---- queue mix ----
    section_bar(ws, QUEUE_HEADER - 2, "4. QUEUE / SKILL MIX (period, site and channel filters apply to every queue)", last_col="N")
    queue_cols = [
        ("Queue / Skill", 22, "@"),
        ("VERINT Offered", 11, INT),
        ("IEX Offered", 11, INT),
        ("VERINT Handled", 11, INT),
        ("IEX Handled", 11, INT),
        ("VERINT Aband %", 11, PCT),
        ("IEX Aband %", 11, PCT),
        ("VERINT SL %", 11, PCT),
        ("IEX SL %", 11, PCT),
        ("VERINT AHT", 11, DEC1),
        ("IEX AHT", 11, DEC1),
    ]
    for i, (label, width, fmt) in enumerate(queue_cols, start=1):
        c = ws.cell(row=QUEUE_HEADER, column=i, value=label)
        header_cell(c, bg=NAVY_L, size=8)
        ws.column_dimensions[get_column_letter(i)].width = max(
            ws.column_dimensions[get_column_letter(i)].width or 0, width
        )

    for i in range(len(QUEUES)):
        r = QUEUE_START + i
        ws[f"A{r}"] = f'=IF(Lists!B{i+5}="","",Lists!B{i+5})'
        for prefix, cols in (
            ("VER", ("B", "D", "F", "H", "J")),
            ("IEX", ("C", "E", "G", "I", "K")),
        ):
            off, hand, abpct, sl, aht = cols
            crit_q = (f'{prefix}_DATE,">="&$B$3,{prefix}_DATE,"<="&$B$4,'
                      f'{prefix}_SITE,$D$3,{prefix}_QUEUE,$A{r},{prefix}_CHAN,$D$5')
            ws[f"{off}{r}"] = f'=IF($A{r}="","",{sumif(f"{prefix}_OFFERED", crit_q)})'
            ws[f"{hand}{r}"] = f'=IF($A{r}="","",{sumif(f"{prefix}_HANDLED", crit_q)})'
            ws[f"{abpct}{r}"] = (
                f'=IF($A{r}="","",IFERROR({sumif(f"{prefix}_ABAND", crit_q)}/{off}{r},0))'
            )
            sp_q = (f'SUMPRODUCT(({prefix}_DATE>=$B$3)*({prefix}_DATE<=$B$4)*'
                    f'{or_cond(prefix, "SITE", "$B$6")}*'
                    f'({prefix}_QUEUE=$A{r})*'
                    f'{or_cond(prefix, "CHAN", "$B$8")}*'
                    f'{prefix}_SL*{prefix}_OFFERED)')
            ws[f"{sl}{r}"] = f'=IF($A{r}="","",IFERROR({sp_q}/{off}{r},0))'
            ws[f"{aht}{r}"] = (
                f'=IF($A{r}="","",IFERROR(({sumif(f"{prefix}_TALK", crit_q)}'
                f'+{sumif(f"{prefix}_HOLD", crit_q)}+{sumif(f"{prefix}_ACW", crit_q)})/{hand}{r},0))'
            )
        for i2, (_, _, fmt) in enumerate(queue_cols, start=1):
            cell = ws.cell(row=r, column=i2)
            cell.number_format = fmt
            cell.font = Font(name=FONT, size=9, color="1A1A1A")
            cell.border = BOTTOM
        ws[f"A{r}"].font = Font(name=FONT, size=9, color=NAVY)

    # subtle grid for readability
    for r in range(KPI_START, KPI_START + len(KPI_DEFS)):
        ws[f"C{r}"].alignment = Alignment(horizontal="right")

    return ws


# --------------------------------------------------------------------------
# Dashboard
# --------------------------------------------------------------------------
def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard")
    ws.sheet_properties.tabColor = GREEN
    ws.sheet_view.showGridLines = False
    for col in "ABCDEFGHIJKLMN":
        ws.column_dimensions[col].width = 12.5

    title_block(ws, "VERINT + IEX | WFM PERFORMANCE DASHBOARD", "", last_col="J")
    ws["A2"] = ('="Period "&TEXT(Calc_Data!$B$3,"dd mmm yyyy")&" to "&'
                'TEXT(Calc_Data!$B$4,"dd mmm yyyy")&"   |   Source: "&Setup!$B$13&'
                '"   |   Site: "&Setup!$B$10&"   |   Queue: "&Setup!$B$11&'
                '"   |   Channel: "&Setup!$B$12')
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color=GRAY)
    ws.merge_cells("A2:J2")

    # account line
    ws["A3"] = "=Setup!$B$16"
    ws["A3"].font = Font(name=FONT, size=11, bold=True, color=NAVY)
    ws.merge_cells("A3:J3")

    def kpi_card(cell_range, label, formula, fmt, accent=NAVY_L):
        top_left, bottom_right = cell_range.split(":")
        start_col = "".join(ch for ch in top_left if ch.isalpha())
        end_col = "".join(ch for ch in bottom_right if ch.isalpha())
        top_row = int("".join(ch for ch in top_left if ch.isdigit()))
        bottom_row = int("".join(ch for ch in bottom_right if ch.isdigit()))

        # label strip
        ws.merge_cells(f"{start_col}{top_row}:{end_col}{top_row}")
        c = ws[f"{start_col}{top_row}"]
        c.value = label
        c.font = Font(name=FONT, size=8, bold=True, color=WHITE)
        c.fill = fill(accent)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[top_row].height = 19

        # value block
        ws.merge_cells(f"{start_col}{top_row + 1}:{end_col}{bottom_row}")
        v = ws[f"{start_col}{top_row + 1}"]
        v.value = formula
        v.number_format = fmt
        v.font = Font(name=FONT, size=21, bold=True, color=NAVY)
        v.fill = fill("F4F8FB")
        v.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[top_row + 1].height = 26

        # frame
        for r in range(top_row, bottom_row + 1):
            for col_idx in range(ws[f"{start_col}{top_row}"].column,
                                 ws[f"{end_col}{top_row}"].column + 1):
                ws.cell(row=r, column=col_idx).border = BOX

    pick = 'MATCH(Setup!$B$13,Calc_Data!$C$13:$E$13,0)'
    kpi_row = {
        "offered": KPI_START, "handled": KPI_START + 1,
        "abandon_rate": KPI_START + 3, "sl": KPI_START + 4,
        "aht": KPI_START + 6, "staffed": KPI_START + 7,
        "required": KPI_START + 8, "gap": KPI_START + 9,
        "occ": KPI_START + 11, "intervals": KPI_START + 12,
    }

    def kpi_formula(r):
        return f"=INDEX(Calc_Data!$C${r}:$E${r},{pick})"

    section_bar(ws, 5, "HEADLINE KPIs", last_col="J")
    cards_top = [
        ("A7:B9", "CONTACTS OFFERED", kpi_formula(kpi_row["offered"]), INT, NAVY_L),
        ("C7:D9", "CONTACTS HANDLED", kpi_formula(kpi_row["handled"]), INT, NAVY_L),
        ("E7:F9", "ABANDON RATE", kpi_formula(kpi_row["abandon_rate"]), PCT, RED),
        ("G7:H9", "SERVICE LEVEL", kpi_formula(kpi_row["sl"]), PCT, TEAL),
        ("I7:J9", "AVG HANDLE TIME", kpi_formula(kpi_row["aht"]), DEC1, "3E7CA6"),
    ]
    cards_bottom = [
        ("A12:B14", "STAFFED AGENT-HOURS", kpi_formula(kpi_row["staffed"]), DEC1, "6B5B95"),
        ("C12:D14", "REQUIRED AGENT-HOURS", kpi_formula(kpi_row["required"]), DEC1, "6B5B95"),
        ("E12:F14", "STAFFING GAP (HOURS)", kpi_formula(kpi_row["gap"]), DEC1, AMBER),
        ("G12:H14", "OCCUPANCY", kpi_formula(kpi_row["occ"]), PCT, "3E7CA6"),
        ("I12:J14", "INTERVALS CAPTURED", kpi_formula(kpi_row["intervals"]), INT, GRAY),
    ]
    for rng, label, formula, fmt, accent in cards_top + cards_bottom:
        kpi_card(rng, label, formula, fmt, accent)

    # reconciliation strip
    section_bar(ws, 16, "SOURCE RECONCILIATION", last_col="J")
    ws["A17"] = "Days with variance"
    ws["A17"].font = Font(name=FONT, size=9, bold=True, color=GRAY)
    ws["B17"] = '=COUNTIF(Reconciliation!$S$6:$S$36,"REVIEW")'
    ws["B17"].font = Font(name=FONT, size=11, bold=True, color=RED)
    ws["B17"].alignment = Alignment(horizontal="center")

    ws["D17"] = "Variance tolerance"
    ws["D17"].font = Font(name=FONT, size=9, bold=True, color=GRAY)
    ws["E17"] = "=Setup!$B$20"
    ws["E17"].number_format = PCT
    ws["E17"].alignment = Alignment(horizontal="center")
    ws["E17"].font = Font(name=FONT, size=11, bold=True, color=NAVY)

    ws["G17"] = "Data quality alerts"
    ws["G17"].font = Font(name=FONT, size=9, bold=True, color=GRAY)
    ws["H17"] = '=COUNTIF(DQ_Checks!$E$6:$E$30,"FAIL")'
    ws["H17"].font = Font(name=FONT, size=11, bold=True, color=RED)
    ws["H17"].alignment = Alignment(horizontal="center")

    ws["I17"] = "Open Reconciliation >"
    ws["I17"].font = Font(name=FONT, size=9, bold=True, color=ACCENT)
    ws["I17"].hyperlink = "#Reconciliation!A1"

    # ---- charts ----
    def style_chart(chart, title, y_title, height=7.4, width=15.2):
        chart.title = title
        chart.y_axis.title = y_title
        chart.height = height
        chart.width = width
        chart.style = 2
        chart.y_axis.majorGridlines = None
        chart.legend.position = "b"

    daily_cats = Reference(wb["Calc_Data"], min_col=1, min_row=DAILY_START,
                           max_row=DAILY_START + DAILY_ROWS - 1)
    chart1 = LineChart()
    style_chart(chart1, "Daily contact volume - Verint vs IEX", "Contacts")
    chart1.add_data(Reference(wb["Calc_Data"], min_col=3, min_row=DAILY_HEADER,
                              max_row=DAILY_START + DAILY_ROWS - 1), titles_from_data=True)
    chart1.add_data(Reference(wb["Calc_Data"], min_col=4, min_row=DAILY_HEADER,
                              max_row=DAILY_START + DAILY_ROWS - 1), titles_from_data=True)
    chart1.set_categories(daily_cats)
    for s in chart1.series:
        s.smooth = False
    ws.add_chart(chart1, "A21")

    chart2 = LineChart()
    style_chart(chart2, "Daily service level & abandon rate", "Percent")
    for col in (11, 12, 9, 10):
        chart2.add_data(Reference(wb["Calc_Data"], min_col=col, min_row=DAILY_HEADER,
                                  max_row=DAILY_START + DAILY_ROWS - 1), titles_from_data=True)
    chart2.set_categories(daily_cats)
    ws.add_chart(chart2, "H21")

    chart3 = BarChart()
    chart3.type = "col"
    chart3.grouping = "clustered"
    style_chart(chart3, "Interval profile - contacts offered by interval", "Contacts")
    chart3.add_data(Reference(wb["Calc_Data"], min_col=3, min_row=PROFILE_HEADER,
                              max_row=PROFILE_START + PROFILE_ROWS - 1), titles_from_data=True)
    chart3.add_data(Reference(wb["Calc_Data"], min_col=4, min_row=PROFILE_HEADER,
                              max_row=PROFILE_START + PROFILE_ROWS - 1), titles_from_data=True)
    chart3.set_categories(Reference(wb["Calc_Data"], min_col=2, min_row=PROFILE_START,
                                    max_row=PROFILE_START + PROFILE_ROWS - 1))
    ws.add_chart(chart3, "A38")

    chart4 = BarChart()
    chart4.type = "bar"
    chart4.grouping = "clustered"
    style_chart(chart4, "Queue / skill mix - offered by source", "Contacts")
    chart4.add_data(Reference(wb["Calc_Data"], min_col=2, min_row=QUEUE_HEADER,
                              max_row=QUEUE_START + len(QUEUES) - 1), titles_from_data=True)
    chart4.add_data(Reference(wb["Calc_Data"], min_col=3, min_row=QUEUE_HEADER,
                              max_row=QUEUE_START + len(QUEUES) - 1), titles_from_data=True)
    chart4.set_categories(Reference(wb["Calc_Data"], min_col=1, min_row=QUEUE_START,
                                    max_row=QUEUE_START + len(QUEUES) - 1))
    ws.add_chart(chart4, "H38")

    # footer
    ws.merge_cells("A55:J55")
    ws["A55"] = ("Reads from Verint_Raw and IEX_Raw through Calc_Data. "
                 "Change the source, site, queue or channel on Setup.")
    ws["A55"].font = Font(name=FONT, size=8, italic=True, color=GRAY)
    ws["A55"].alignment = Alignment(horizontal="left", indent=1)

    return ws


# --------------------------------------------------------------------------
# Reconciliation
# --------------------------------------------------------------------------
def build_reconciliation(wb):
    ws = wb.create_sheet("Reconciliation")
    ws.sheet_properties.tabColor = AMBER
    ws.sheet_view.showGridLines = False
    title_block(ws, "VERINT / IEX RECONCILIATION",
                "Side-by-side daily totals with variance flags. Tolerance is set on Setup.", last_col="S")

    ws["A3"] = "Period"
    ws["B3"] = "=Calc_Data!$B$3"
    ws["B3"].number_format = DATE_FMT
    ws["D3"] = "Through"
    ws["E3"] = "=Calc_Data!$B$4"
    ws["E3"].number_format = DATE_FMT
    ws["G3"] = "Tolerance"
    ws["H3"] = "=Setup!$B$20"
    ws["H3"].number_format = PCT
    ws["J3"] = "Days flagged REVIEW"
    ws["K3"] = '=COUNTIF(S6:S36,"REVIEW")'
    for cell in ("A3", "D3", "G3", "J3"):
        ws[cell].font = Font(name=FONT, size=9, bold=True, color=GRAY)
    for cell in ("B3", "E3", "H3", "K3"):
        ws[cell].font = Font(name=FONT, size=10, bold=True, color=NAVY)
        ws[cell].alignment = Alignment(horizontal="center")

    headers = [
        ("Date", 12), ("Day", 8),
        ("V Offered", 10), ("IEX Offered", 11), ("Var", 8), ("Var %", 9),
        ("V Handled", 10), ("IEX Handled", 11), ("Var", 8),
        ("V Aband", 9), ("IEX Aband", 10), ("Var", 8),
        ("V SL", 8), ("IEX SL", 8), ("SL Diff", 9),
        ("V AHT", 8), ("IEX AHT", 8), ("AHT Diff", 9),
        ("Status", 11),
    ]
    hdr = 5
    for i, (label, width) in enumerate(headers, start=1):
        col = get_column_letter(i)
        ws.column_dimensions[col].width = width
        c = ws.cell(row=hdr, column=i, value=label)
        header_cell(c, size=8, bg=NAVY_L if i < 19 else TEAL)
    ws.row_dimensions[hdr].height = 30

    for i in range(DAILY_ROWS):
        r = 6 + i
        cr = DAILY_START + i
        ws[f"A{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!A{cr})'
        ws[f"A{r}"].number_format = DATE_FMT
        ws[f"B{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!B{cr})'
        ws[f"C{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!C{cr})'
        ws[f"D{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!D{cr})'
        ws[f"E{r}"] = f'=IF($A{r}="","",N(C{r})-N(D{r}))'
        ws[f"F{r}"] = f'=IF(OR($A{r}="",N(D{r})=0),"",E{r}/D{r})'
        ws[f"G{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!E{cr})'
        ws[f"H{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!F{cr})'
        ws[f"I{r}"] = f'=IF($A{r}="","",N(G{r})-N(H{r}))'
        ws[f"J{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!G{cr})'
        ws[f"K{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!H{cr})'
        ws[f"L{r}"] = f'=IF($A{r}="","",N(J{r})-N(K{r}))'
        ws[f"M{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!K{cr})'
        ws[f"N{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!L{cr})'
        ws[f"O{r}"] = f'=IF($A{r}="","",(N(M{r})-N(N{r}))*100)'
        ws[f"P{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!M{cr})'
        ws[f"Q{r}"] = f'=IF(Calc_Data!A{cr}="","",Calc_Data!N{cr})'
        ws[f"R{r}"] = f'=IF($A{r}="","",N(P{r})-N(Q{r}))'
        ws[f"S{r}"] = (
            f'=IF($A{r}="","",IF(AND(ABS(N(F{r}))<=$H$3,'
            f'ABS(N(I{r}))<=MAX(2,0.02*N(G{r}))),"MATCH","REVIEW"))'
        )
        for i2, (_, _) in enumerate(headers, start=1):
            cell = ws.cell(row=r, column=i2)
            cell.border = BOTTOM
            cell.font = Font(name=FONT, size=9)
        for col, fmt in (("C", INT), ("D", INT), ("E", INT), ("F", PCT), ("G", INT),
                         ("H", INT), ("I", INT), ("J", INT), ("K", INT), ("L", INT),
                         ("M", PCT), ("N", PCT), ("O", "0.0"), ("P", DEC1),
                         ("Q", DEC1), ("R", DEC1)):
            ws[f"{col}{r}"].number_format = fmt
        ws[f"A{r}"].alignment = Alignment(horizontal="center")
        ws[f"B{r}"].alignment = Alignment(horizontal="center")
        ws[f"S{r}"].alignment = Alignment(horizontal="center")
        ws[f"S{r}"].font = Font(name=FONT, size=9, bold=True)

    # totals
    tr = 6 + DAILY_ROWS
    ws[f"A{tr}"] = "TOTAL"
    ws[f"A{tr}"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
    for col in ("C", "D", "G", "H", "J", "K"):
        ws[f"{col}{tr}"] = f"=SUM({col}6:{col}{tr-1})"
        ws[f"{col}{tr}"].number_format = INT
    ws[f"E{tr}"] = f"=SUM(E6:E{tr-1})"
    ws[f"E{tr}"].number_format = INT
    ws[f"F{tr}"] = f'=IFERROR(E{tr}/D{tr},0)'
    ws[f"F{tr}"].number_format = PCT
    ws[f"I{tr}"] = f"=SUM(I6:I{tr-1})"
    ws[f"L{tr}"] = f"=SUM(L6:L{tr-1})"
    ws[f"M{tr}"] = f"=Calc_Data!C{KPI_START + 4}"
    ws[f"N{tr}"] = f"=Calc_Data!D{KPI_START + 4}"
    ws[f"O{tr}"] = f'=(M{tr}-N{tr})*100'
    ws[f"O{tr}"].number_format = "0.0"
    ws[f"M{tr}"].number_format = PCT
    ws[f"N{tr}"].number_format = PCT
    ws[f"P{tr}"] = f"=Calc_Data!C{KPI_START + 6}"
    ws[f"Q{tr}"] = f"=Calc_Data!D{KPI_START + 6}"
    ws[f"R{tr}"] = f'=P{tr}-Q{tr}'
    for col in ("P", "Q", "R"):
        ws[f"{col}{tr}"].number_format = DEC1
    ws[f"S{tr}"] = f'=COUNTIF(S6:S{tr-1},"REVIEW")&" flagged"'
    ws[f"S{tr}"].font = Font(name=FONT, size=9, bold=True, color=RED)
    for i2 in range(1, len(headers) + 1):
        cell = ws.cell(row=tr, column=i2)
        cell.fill = fill(LIGHT)
        cell.border = Border(top=Side(style="thin", color=NAVY), bottom=Side(style="double", color=NAVY))
        if i2 > 1:
            cell.font = Font(name=FONT, size=9, bold=True, color=NAVY)
            cell.alignment = Alignment(horizontal="right")

    # conditional formatting: variance above tolerance
    red_fill = fill(RED_L)
    amber_fill = fill(AMBER_L)
    ws.conditional_formatting.add(
        "F6:F36",
        CellIsRule(operator="greaterThan", formula=["$H$3"], fill=red_fill,
                   font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(
        "F6:F36",
        CellIsRule(operator="lessThan", formula=["-$H$3"], fill=red_fill,
                   font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(
        "S6:S36", FormulaRule(formula=['$S6="REVIEW"'], fill=red_fill, font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(
        "S6:S36", FormulaRule(formula=['$S6="MATCH"'], fill=fill(GREEN_L), font=Font(color=GREEN)))

    ws.freeze_panes = "A6"
    return ws


# --------------------------------------------------------------------------
# Data quality checks
# --------------------------------------------------------------------------
def build_dq(wb):
    ws = wb.create_sheet("DQ_Checks")
    ws.sheet_properties.tabColor = RED
    ws.sheet_view.showGridLines = False
    title_block(ws, "DATA QUALITY CHECKS",
                "Run these before reading any report. FAIL rows are data problems, not formula errors.",
                last_col="F")
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 58

    hdr = 5
    for i, label in enumerate(["#", "Check", "Source", "Result", "Status", "What to do"], start=1):
        c = ws.cell(row=hdr, column=i, value=label)
        header_cell(c, size=9)

    checks = []
    for prefix, label in (("VER", "Verint"), ("IEX", "IEX")):
        checks += [
            (f"Rows loaded in period ({label})",
             f'=COUNTIFS({prefix}_DATE,">="&Calc_Data!$B$3,{prefix}_DATE,"<="&Calc_Data!$B$4)',
             label, INT, "info", "Count of raw rows inside the reporting period."),
            (f"Missing date ({label})",
             f'=SUMPRODUCT(--ISBLANK({prefix}_DATE),--({prefix}_QUEUE<>""))',
             label, INT, "fail_gt0", "Every row needs a Date. Delete or fix blank-date rows."),
            (f"Missing queue / skill ({label})",
             f'=SUMPRODUCT(--ISBLANK({prefix}_QUEUE),--({prefix}_DATE<>""))',
             label, INT, "fail_gt0", "Fill Queue / Skill so filters and reconciliation work."),
            (f"Missing interval start ({label})",
             f'=SUMPRODUCT(--ISBLANK({prefix}_TIME),--({prefix}_DATE<>""))',
             label, INT, "fail_gt0", "Interval Start must be a time value (hh:mm)."),
            (f"Abandoned greater than offered ({label})",
             f'=SUMPRODUCT(--(({prefix}_ABAND>{prefix}_OFFERED)*({prefix}_OFFERED<>"")))',
             label, INT, "fail_gt0", "Check the export mapping - abandons cannot exceed offered."),
            (f"Negative volumes ({label})",
             f'=SUMPRODUCT(--((({prefix}_OFFERED<0)+({prefix}_HANDLED<0)+({prefix}_ABAND<0))>0))',
             label, INT, "fail_gt0", "Remove negative values; they break every total."),
            (f"Service level outside 0-100% ({label})",
             f'=SUMPRODUCT(--((({prefix}_SL<0)+({prefix}_SL>1))>0))',
             label, INT, "fail_gt0", "Store Service Level as a percentage (0.85 = 85%), not 85."),
            (f"Handled without handle time ({label})",
             f'=SUMPRODUCT(--(({prefix}_HANDLED>0)*({prefix}_AHT=0)))',
             label, INT, "fail_gt0", "AHT or talk/hold/ACW is missing - weighted AHT will read 0."),
            (f"Dates outside the selected period ({label})",
             f'=SUMPRODUCT(--({prefix}_DATE<>""),--({prefix}_DATE<Calc_Data!$B$3))'
             f'+SUMPRODUCT(--({prefix}_DATE<>""),--({prefix}_DATE>Calc_Data!$B$4))',
             label, INT, "info", "Rows exist outside the period; they are excluded from reports."),
            (f"Sample rows present - delete before live use ({label})",
             f'=COUNTIF({prefix}_SRC,"SAMPLE*")',
             label, INT, "info", "Template demo rows. Clear rows 6:101 once real data is loaded."),
        ]

    r = 6
    for i, (name, formula, source, fmt, rule, action) in enumerate(checks, start=1):
        ws[f"A{r}"] = i
        ws[f"B{r}"] = name
        ws[f"C{r}"] = source
        ws[f"D{r}"] = formula
        ws[f"D{r}"].number_format = fmt
        if rule == "fail_gt0":
            ws[f"E{r}"] = f'=IF(D{r}>0,"FAIL","PASS")'
        else:
            ws[f"E{r}"] = '="INFO"'
        ws[f"F{r}"] = action
        for col in "ABCDEF":
            cell = ws[f"{col}{r}"]
            cell.border = BOTTOM
            cell.font = Font(name=FONT, size=9, color="1A1A1A")
        ws[f"A{r}"].alignment = Alignment(horizontal="center")
        ws[f"C{r}"].alignment = Alignment(horizontal="center")
        ws[f"D{r}"].alignment = Alignment(horizontal="center")
        ws[f"E{r}"].alignment = Alignment(horizontal="center")
        ws[f"E{r}"].font = Font(name=FONT, size=9, bold=True)
        ws[f"F{r}"].font = Font(name=FONT, size=9, color=GRAY)
        r += 1

    last = r - 1
    ws.conditional_formatting.add(
        f"E6:E{last}", FormulaRule(formula=['$E6="FAIL"'], fill=fill(RED_L), font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(
        f"E6:E{last}", FormulaRule(formula=['$E6="PASS"'], fill=fill(GREEN_L), font=Font(color=GREEN, bold=True)))

    # summary
    sr = last + 2
    ws[f"B{sr}"] = "Checks failing"
    ws[f"B{sr}"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
    ws[f"D{sr}"] = f'=COUNTIF(E6:E{last},"FAIL")'
    ws[f"D{sr}"].font = Font(name=FONT, size=14, bold=True, color=RED)
    ws[f"D{sr}"].alignment = Alignment(horizontal="center")

    ws[f"B{sr+2}"] = ("Status meanings: PASS = no records matched the rule; FAIL = records need fixing; "
                      "INFO = informational count.")
    ws.merge_cells(f"B{sr+2}:F{sr+2}")
    ws[f"B{sr+2}"].font = Font(name=FONT, size=9, italic=True, color=GRAY)

    ws.freeze_panes = "A6"
    return ws


# --------------------------------------------------------------------------
# README
# --------------------------------------------------------------------------
def build_readme(wb):
    ws = wb.create_sheet("README", 0)
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines = False
    widths = {"A": 4, "B": 28, "C": 30, "D": 30, "E": 18, "F": 46, "G": 3, "H": 3}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    title_block(ws, "VERINT + IEX RAW DATA TEMPLATE",
                "One schema, two sources, one reconciled view of interval performance.",
                last_col="F")
    ws.row_dimensions[1].height = 38

    # ---- intro ----
    section_bar(ws, 4, "WHAT THIS WORKBOOK DOES", last_col="F")
    paragraphs = [
        "Captures raw Verint Workforce Management and IEX TotalView interval exports, keeps them "
        "untouched, then calculates the same KPIs for both sources and reconciles them day by day.",
        "Use it to answer three questions every WFM cycle: what happened (volume, handle time, service "
        "level), does the data agree across systems, and where is staffing short of requirement.",
    ]
    r = 5
    for text in paragraphs:
        ws.merge_cells(f"B{r}:F{r}")
        c = ws[f"B{r}"]
        c.value = text
        c.font = Font(name=FONT, size=10, color="1A1A1A")
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[r].height = 32
        r += 1

    # ---- sheet map ----
    section_bar(ws, r + 1, "SHEET MAP", last_col="F")
    r += 2
    sheet_map = [
        ("README", "This guide", "Read first.", NAVY),
        ("Setup", "Period, filters, tolerance", "Only place you change inputs.", ACCENT),
        ("Verint_Raw", "Paste the Verint export", "Data only - no formulas.", "C05621"),
        ("IEX_Raw", "Paste the IEX export", "Data only - no formulas.", "C05621"),
        ("Calc_Data", "Calculation engine", "Totals, daily trend, interval profile, queue mix.", "8A93A5"),
        ("Dashboard", "KPIs and charts", "Reads Calc_Data using the Setup source.", GREEN),
        ("Reconciliation", "Verint vs IEX by day", "Variance flags against the tolerance.", AMBER),
        ("DQ_Checks", "Data quality rules", "Fix FAIL rows before reporting.", RED),
        ("Lists", "Sites, queues, channels", "Feeds every dropdown and the queue mix.", GRAY),
    ]
    for name, purpose, note, color in sheet_map:
        c = ws[f"B{r}"]
        c.value = name
        c.font = Font(name=FONT, size=10, bold=True, color=WHITE)
        c.fill = fill(color)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c.border = BOX
        ws.merge_cells(f"C{r}:E{r}")
        ws[f"C{r}"] = purpose
        ws[f"C{r}"].font = Font(name=FONT, size=9, color=NAVY)
        ws[f"C{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws[f"F{r}"] = note
        ws[f"F{r}"].font = Font(name=FONT, size=9, color=GRAY)
        ws[f"F{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in "BCDEF":
            ws[f"{col}{r}"].border = BOTTOM
        r += 1

    # ---- workflow ----
    section_bar(ws, r + 1, "WORKFLOW", last_col="F")
    r += 2
    steps = [
        ("1", "Set the month", "Setup!B5 = first day of the reporting month. Period End calculates itself."),
        ("2", "Paste Verint", "Verint_Raw rows 6+ - align to the header row, paste as values."),
        ("3", "Paste IEX", "IEX_Raw rows 6+ - same schema, same order."),
        ("4", "Clear the sample", "Delete rows 6:101 on both raw sheets when you go live."),
        ("5", "Run DQ_Checks", "Every FAIL row is a real data problem - fix it at source."),
        ("6", "Read the Dashboard", "Pick VERINT, IEX or COMBINED in Setup!B13."),
        ("7", "Reconcile", "Reconciliation compares both pulls per day; investigate REVIEW rows."),
    ]
    for num, title, detail in steps:
        ws[f"B{r}"] = f"{num}.  {title}"
        ws[f"B{r}"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
        ws[f"B{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.merge_cells(f"C{r}:F{r}")
        ws[f"C{r}"] = detail
        ws[f"C{r}"].font = Font(name=FONT, size=9, color="1A1A1A")
        ws[f"C{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[r].height = 20
        for col in "BCDEF":
            ws[f"{col}{r}"].border = BOTTOM
        r += 1

    # ---- column dictionary ----
    section_bar(ws, r + 1, "COLUMN DICTIONARY  (identical schema on both raw sheets)", last_col="F")
    r += 2
    for i, h in enumerate(["Column", "Verint export name", "IEX TotalView name",
                           "Type / format", "Notes"], start=2):
        header_cell(ws.cell(row=r, column=i, value=h), bg=TEAL)
    ws.row_dimensions[r].height = 20
    r += 1
    dictionary = [
        ("Date", "Date / Interval Date", "Date / Start Date", "Date (yyyy-mm-dd)", "Interval start date, not the file date."),
        ("Interval Start", "Interval Start / Start Time", "Interval / Start Time", "Time (hh:mm)", "Time only - never a date-time value."),
        ("Interval End", "Interval End / End Time", "End Time", "Time (hh:mm)", "Optional but recommended."),
        ("Interval", "Interval / Interval Name", "Interval Label", "Text", "Optional display label."),
        ("Site", "Site / Location", "Location / Site", "Text", "Must match Lists column A."),
        ("Queue / Skill", "Queue / Queue Name", "Skill / Skill Name", "Text", "Must match Lists column B."),
        ("Channel", "Channel / Media Type", "Channel / Media", "Text", "Voice, Chat, Email, Social, Backoffice."),
        ("Offered", "Offered / Contacts Offered", "Offered / Total Offered", "Integer", "Contacts that arrived in the interval."),
        ("Handled", "Handled / Answered", "Handled / Answered", "Integer", "Contacts answered by an agent."),
        ("Abandoned", "Abandoned / Abandons", "Abandoned", "Integer", "Offered minus handled is the sanity check."),
        ("Short Abandoned", "Short Abandoned / Abn < 5s", "Short Abandoned", "Integer", "Optional short-abandon split."),
        ("ASA (sec)", "ASA / Avg Speed of Answer", "ASA", "Decimal", "Weighted by handled in all KPIs."),
        ("AHT (sec)", "AHT / Avg Handle Time", "AHT", "Decimal", "Cross-checked against talk + hold + ACW."),
        ("Talk (sec)", "Talk Time", "Talk Time", "Integer", "Total talk seconds in the interval."),
        ("Hold (sec)", "Hold Time", "Hold Time", "Integer", "Total hold seconds."),
        ("ACW (sec)", "ACW / Wrap Time", "ACW / After Call Work", "Integer", "Total after-call work seconds."),
        ("Staffed (agents)", "Staffed / Agents", "Staffed / Staff", "Decimal", "Average staffed in the interval."),
        ("Required (agents)", "Required / Needed", "Required / Req", "Decimal", "Workforce requirement for the interval."),
        ("Service Level %", "Service Level / SL %", "Service Level", "Percent (0.85)", "Stored as a decimal, formatted as %."),
        ("Max Wait (sec)", "Max Wait / Longest Wait", "Max Wait", "Integer", "Longest wait in the interval."),
        ("Source File", "Export file name", "Export file name", "Text", "Traceability - keep the export name."),
        ("Notes", "Notes", "Notes", "Text", "Anything the next analyst should know."),
    ]
    for col_name, v_name, i_name, typ, note in dictionary:
        ws[f"B{r}"] = col_name
        ws[f"B{r}"].font = Font(name=FONT, size=9, bold=True, color=NAVY)
        ws[f"C{r}"] = v_name
        ws[f"D{r}"] = i_name
        ws[f"E{r}"] = typ
        ws[f"F{r}"] = note
        for col in "CDEF":
            ws[f"{col}{r}"].font = Font(name=FONT, size=8, color="1A1A1A")
            ws[f"{col}{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws[f"B{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in "BCDEF":
            ws[f"{col}{r}"].border = BOTTOM
        ws.row_dimensions[r].height = 22
        r += 1

    # ---- rules ----
    section_bar(ws, r + 1, "RULES, LIMITS AND DEFINITIONS", last_col="F")
    r += 2
    rules = [
        ("Capacity", "10,000 rows per raw sheet (rows 6-10,005). Extend the ranges and update the defined names for larger months."),
        ("Don't", "Never insert rows above row 6, merge cells in the raw sheets, or type into Calc_Data - it breaks every reference."),
        ("Staffed agent-hours", "SUM(staffed agents) x interval minutes / 60, using the Setup period and filters."),
        ("Occupancy", "(talk + hold + ACW seconds) / 3600 divided by staffed agent-hours."),
        ("Service level KPI", "Offered-weighted average of the interval service level, not a plain average."),
        ("AHT KPI", "(total talk + hold + ACW) / handled contacts - the weighted number, not the average of AHTs."),
        ("Staffing gap", "Required agent-hours minus staffed agent-hours. Positive = understaffed."),
        ("Reconciliation flag", "A day is REVIEW when offered variance % or handled variance exceeds the Setup tolerance."),
        ("Report limits", "Daily trend = 31 days; interval profile = 48 intervals (24 h at 30 minutes). Extend the blocks for 15-minute intervals."),
    ]
    for label, text in rules:
        ws[f"B{r}"] = label
        ws[f"B{r}"].font = Font(name=FONT, size=9, bold=True, color=NAVY)
        ws[f"B{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.merge_cells(f"C{r}:F{r}")
        ws[f"C{r}"] = text
        ws[f"C{r}"].font = Font(name=FONT, size=9, color="1A1A1A")
        ws[f"C{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[r].height = 26
        for col in "BCDEF":
            ws[f"{col}{r}"].border = BOTTOM
        r += 1

    return ws


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    wb = Workbook()
    wb.remove(wb.active)

    # order matters for tab display
    build_readme(wb)
    build_setup(wb)
    build_lists(wb)
    build_raw_sheet(wb, "Verint_Raw", "VERINT", "C05621", build_sample_rows("VERINT"))
    build_raw_sheet(wb, "IEX_Raw", "IEX", "C05621", build_sample_rows("IEX", shift=0, seed=7))
    build_calc(wb)
    build_dashboard(wb)
    build_reconciliation(wb)
    build_dq(wb)

    add_defined_names(wb)

    # print setup for the read-only views
    for sheet_name in ("Dashboard", "Reconciliation", "DQ_Checks", "README"):
        sheet = wb[sheet_name]
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.sheet_properties.pageSetUpPr.fitToPage = True

    wb.calculation.fullCalcOnLoad = True

    # workbook level cosmetics
    wb.properties.title = "Verint + IEX Raw Template"
    wb.properties.subject = "WFM interval data capture and reconciliation"
    wb.properties.creator = "WFM Analytics"
    wb.properties.description = (
        "Raw interval template for Verint Workforce Management and IEX TotalView "
        "with formula-driven KPIs, reconciliation and data quality checks."
    )
    wb.active = 0
    wb.save(OUT)
    print(f"Saved {OUT}")


if __name__ == "__main__":
    main()
