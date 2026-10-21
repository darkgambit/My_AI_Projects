# -*- coding: utf-8 -*-
"""
Builds: LGU_Criminal_Case_Monitoring_System.xlsx
A fully functional, formula-driven criminal case tracker for a Philippine LGU.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule, DataBarRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DFont
from openpyxl.drawing.line import LineProperties
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.properties import PageSetupProperties

TODAY = dt.date(2026, 10, 21)
D = lambda n: TODAY + dt.timedelta(days=n)

# ---------------------------------------------------------------- palette
NAVY   = "0B2545"
NAVY_2 = "13315C"
NAVY_3 = "1B4079"
NAVY_4 = "2A4E7A"
CALC   = "4A5568"
GOLD   = "C8A44D"
GOLD_L = "FBF5E6"
INK    = "1F2937"
MUTED  = "6B7280"
LINE   = "D9E1EA"
BG     = "F4F7FA"
WHITE  = "FFFFFF"
GREEN  = "1E7B4F"
GREEN_L= "E7F4EC"
AMBER  = "B7791F"
AMBER_L= "FDF4E3"
RED    = "B42318"
RED_L  = "FCE9E7"
BLUE   = "1D4ED8"
BLUE_L = "E8EEFB"
TEAL   = "0F6E6A"

F = "Segoe UI"

def font(size=10, bold=False, color=INK, italic=False, name=F):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)

def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)

def side(style="thin", color=LINE):
    return Side(style=style, color=color)

def box(color=LINE, style="thin"):
    s = side(style, color)
    return Border(left=s, right=s, top=s, bottom=s)

def valign(h="left", v="center", wrap=False, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)

def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def title_band(ws, rng, text, sub=None, height=34):
    first = rng.split(":")[0]
    ws.merge_cells(rng)
    c = ws[first]
    c.value = text
    c.font = font(15, True, WHITE)
    c.fill = fill(NAVY)
    c.alignment = valign("left", "center", False, 1)
    ws.row_dimensions[c.row].height = height

wb = Workbook()
wb.calculation.fullCalcOnLoad = True

# ================================================================ LISTS
wsL = wb.active
wsL.title = "Lists"
wsL.sheet_properties.tabColor = MUTED
wsL.sheet_view.showGridLines = False

offenses = [
 ("Murder","Crimes Against Persons","RPC Art. 248"),
 ("Frustrated Murder","Crimes Against Persons","RPC Art. 248 in rel. to Art. 6"),
 ("Attempted Murder","Crimes Against Persons","RPC Art. 248 in rel. to Art. 6"),
 ("Homicide","Crimes Against Persons","RPC Art. 249"),
 ("Frustrated Homicide","Crimes Against Persons","RPC Art. 249 in rel. to Art. 6"),
 ("Attempted Homicide","Crimes Against Persons","RPC Art. 249 in rel. to Art. 6"),
 ("Parricide","Crimes Against Persons","RPC Art. 246"),
 ("Infanticide","Crimes Against Persons","RPC Art. 255"),
 ("Serious Physical Injuries","Crimes Against Persons","RPC Art. 263"),
 ("Less Serious Physical Injuries","Crimes Against Persons","RPC Art. 265"),
 ("Slight Physical Injuries","Crimes Against Persons","RPC Art. 266"),
 ("Grave Threats","Crimes Against Persons","RPC Art. 282"),
 ("Grave Coercion","Crimes Against Persons","RPC Art. 286"),
 ("Kidnapping and Serious Illegal Detention","Crimes Against Persons","RPC Art. 267"),
 ("Slight Illegal Detention","Crimes Against Persons","RPC Art. 268"),
 ("Abduction","Crimes Against Persons","RPC Art. 342"),
 ("Reckless Imprudence Resulting in Homicide","Crimes Against Persons","RPC Art. 365"),
 ("Reckless Imprudence Resulting in Physical Injuries","Crimes Against Persons","RPC Art. 365"),
 ("Robbery","Crimes Against Property","RPC Art. 294"),
 ("Robbery with Homicide","Crimes Against Property","RPC Art. 294 (1)"),
 ("Theft","Crimes Against Property","RPC Art. 308"),
 ("Qualified Theft","Crimes Against Property","RPC Art. 310"),
 ("Estafa","Crimes Against Property","RPC Art. 315"),
 ("Estafa thru Falsification","Crimes Against Property","RPC Art. 315 in rel. to Art. 171"),
 ("Malicious Mischief","Crimes Against Property","RPC Art. 327"),
 ("Arson","Crimes Against Property","RPC Art. 320"),
 ("Destructive Arson","Crimes Against Property","RPC Art. 320"),
 ("Trespass to Dwelling","Crimes Against Property","RPC Art. 280"),
 ("Usurpation of Real Property","Crimes Against Property","RPC Art. 312"),
 ("Anti-Cattle Rustling","Crimes Against Property","PD 533"),
 ("Anti-Fencing","Crimes Against Property","PD 1612"),
 ("Rape","Crimes Against Chastity","RPC Art. 266-A"),
 ("Attempted Rape","Crimes Against Chastity","RPC Art. 266-A in rel. to Art. 6"),
 ("Acts of Lasciviousness","Crimes Against Chastity","RPC Art. 336"),
 ("Sexual Harassment","Crimes Against Chastity","RA 7877"),
 ("Safe Spaces Act (GBSH)","Crimes Against Chastity","RA 11313"),
 ("Anti-Photo and Video Voyeurism","Crimes Against Chastity","RA 9995"),
 ("Concubinage","Crimes Against Chastity","RPC Art. 334"),
 ("Adultery","Crimes Against Chastity","RPC Art. 333"),
 ("Direct Assault","Crimes Against Public Order","RPC Art. 148"),
 ("Indirect Assault","Crimes Against Public Order","RPC Art. 149"),
 ("Resistance and Disobedience","Crimes Against Public Order","RPC Art. 151"),
 ("Alarm and Scandal","Crimes Against Public Order","RPC Art. 155"),
 ("Tumults and Other Disturbances","Crimes Against Public Order","RPC Art. 153"),
 ("Illegal Possession of Firearms","Crimes Against Public Order","RA 10591"),
 ("Illegal Possession of Explosives","Crimes Against Public Order","RA 9516"),
 ("Illegal Gambling","Crimes Against Public Order","PD 1602"),
 ("Illegal Cockfighting","Crimes Against Public Order","PD 449"),
 ("Illegal Numbers Game (Jueteng, etc.)","Crimes Against Public Order","RA 9287"),
 ("Driving Under the Influence","Crimes Against Public Order","RA 10586"),
 ("Reckless Driving","Crimes Against Public Order","RA 4136"),
 ("Anti-Distracted Driving","Crimes Against Public Order","RA 10913"),
 ("Perjury","Crimes Against Public Interest","RPC Art. 183"),
 ("Falsification of Public Documents","Crimes Against Public Interest","RPC Art. 171"),
 ("Falsification of Private Documents","Crimes Against Public Interest","RPC Art. 172"),
 ("Usurpation of Authority","Crimes Against Public Interest","RPC Art. 177"),
 ("Direct Bribery","Crimes Against Public Interest","RPC Art. 210"),
 ("Anti-Graft and Corrupt Practices","Crimes Against Public Interest","RA 3019"),
 ("Plunder","Crimes Against Public Interest","RA 7080"),
 ("Grave Oral Defamation","Crimes Against Public Interest","RPC Art. 358"),
 ("Libel","Crimes Against Public Interest","RPC Art. 355"),
 ("Anti-Wiretapping","Crimes Against Public Interest","RA 4200"),
 ("Illegal Sale of Dangerous Drugs","Special Laws - Dangerous Drugs","RA 9165 Sec. 5"),
 ("Illegal Possession of Dangerous Drugs","Special Laws - Dangerous Drugs","RA 9165 Sec. 11"),
 ("Illegal Use of Dangerous Drugs","Special Laws - Dangerous Drugs","RA 9165 Sec. 15"),
 ("Possession of Drug Paraphernalia","Special Laws - Dangerous Drugs","RA 9165 Sec. 12"),
 ("Maintenance of a Drug Den","Special Laws - Dangerous Drugs","RA 9165 Sec. 6"),
 ("Cultivation of Dangerous Drugs","Special Laws - Dangerous Drugs","RA 9165 Sec. 16"),
 ("Other Violation of RA 9165","Special Laws - Dangerous Drugs","RA 9165"),
 ("Violation of VAWC Act","Special Laws - Women and Children","RA 9262"),
 ("Child Abuse","Special Laws - Women and Children","RA 7610"),
 ("Trafficking in Persons","Special Laws - Women and Children","RA 9208"),
 ("Expanded Anti-Trafficking in Persons","Special Laws - Women and Children","RA 10364"),
 ("Anti-Child Pornography","Special Laws - Women and Children","RA 9775"),
 ("Anti-OSAEC and CSAEM","Special Laws - Women and Children","RA 11930"),
 ("Child Labor Violation","Special Laws - Women and Children","RA 9231"),
 ("Illegal Logging","Special Laws - Environment and Natural Resources","PD 705"),
 ("Illegal Fishing","Special Laws - Environment and Natural Resources","RA 8550"),
 ("Illegal Quarrying","Special Laws - Environment and Natural Resources","RA 7942"),
 ("Violation of Wildlife Act","Special Laws - Environment and Natural Resources","RA 9147"),
 ("Violation of Chainsaw Act","Special Laws - Environment and Natural Resources","RA 9175"),
 ("Violation of Ecological Solid Waste Act","Special Laws - Environment and Natural Resources","RA 9003"),
 ("Violation of Clean Air Act","Special Laws - Environment and Natural Resources","RA 8749"),
 ("Violation of Clean Water Act","Special Laws - Environment and Natural Resources","RA 9275"),
 ("Cybercrime (Illegal Access / Fraud)","Special Laws - Economic and Cyber","RA 10175"),
 ("Cyber Libel","Special Laws - Economic and Cyber","RA 10175 in rel. to RPC Art. 355"),
 ("Online Scam / Financial Account Scam","Special Laws - Economic and Cyber","RA 12065"),
 ("Bouncing Checks","Special Laws - Economic and Cyber","BP 22"),
 ("Illegal Recruitment","Special Laws - Economic and Cyber","RA 8042"),
 ("Violation of Consumer Act","Special Laws - Economic and Cyber","RA 7394"),
 ("Price Act Violation (Hoarding / Profiteering)","Special Laws - Economic and Cyber","RA 7581"),
 ("Violation of Intellectual Property Code","Special Laws - Economic and Cyber","RA 8293"),
 ("Anti-Money Laundering","Special Laws - Economic and Cyber","RA 9160"),
 ("Violation of Local Ordinance","Local Ordinances and Other Special Laws","Local Ordinance"),
 ("Violation of Barangay Ordinance","Local Ordinances and Other Special Laws","Barangay Ordinance"),
 ("Violation of National Building Code","Local Ordinances and Other Special Laws","PD 1096"),
 ("Violation of Fire Code","Local Ordinances and Other Special Laws","RA 9514"),
 ("Violation of Sanitation Code","Local Ordinances and Other Special Laws","PD 856"),
 ("Violation of Anti-Smoking Ordinance","Local Ordinances and Other Special Laws","RA 9211 / Local Ordinance"),
]

categories = ["Crimes Against Persons","Crimes Against Property","Crimes Against Chastity",
              "Crimes Against Public Order","Crimes Against Public Interest",
              "Special Laws - Dangerous Drugs","Special Laws - Women and Children",
              "Special Laws - Environment and Natural Resources","Special Laws - Economic and Cyber",
              "Local Ordinances and Other Special Laws"]

stages = [("For Filing",5),("Preliminary Investigation",20),("Filed in Court",35),
          ("Arraignment",45),("Pre-Trial",60),("Trial",75),("Submitted for Decision",88),
          ("Judgment Rendered",97),("Appeal / Review",97),("Archived",100)]
statuses = ["Active","For Action","Deferred","Dormant","Closed","Endorsed / Transferred"]
outcomes = ["Pending","Convicted","Acquitted","Dismissed","Amicably Settled","Archived / Referred","N/A"]
priorities = ["High","Medium","Low"]
courts = ["Katarungang Pambarangay","Office of the Prosecutor","Regional Trial Court",
          "Metropolitan Trial Court","Municipal Trial Court","Municipal Circuit Trial Court",
          "Sandiganbayan","Court of Appeals","Supreme Court"]
custody = ["In Custody - No Bail","In Custody - With Bail","Released on Bail",
           "Released on Recognizance","At Large","N/A (Not in Custody)"]
hear_types = ["Arraignment","Pre-Trial","Trial","Promulgation","Motion Hearing","Inquest",
              "Preliminary Investigation","Mediation / Conciliation"]
hear_results = ["Held as Scheduled","Postponed - Prosecution","Postponed - Defense",
                "Postponed - Court","Reset","Cancelled","Submitted for Resolution"]
personnel = ["Atty. R. Villanueva","Atty. M. Fernandez","Atty. J. Aquino","Atty. C. Bautista",
             "Atty. S. Navarro","Investigator P. Reyes","Investigator L. Gomez","Unassigned"]
barangays = ["Barangay 1","Barangay 2","Barangay 3","Barangay 4","Barangay 5","Barangay 6",
             "Barangay 7","Barangay 8","Barangay 9","Barangay 10","Poblacion","Outside Jurisdiction"]
docs = ["Affidavit of Complaint","Sworn Statement(s)","Police / Incident Report","Evidence and Exhibits",
        "Resolution / Information","Court Order / Docket","Bail / Bond Records","Judgment / Order"]

# ---- layout: Blocks -------------------------------------------------
wsL["A1"] = "REFERENCE LISTS  |  DATA DICTIONARY"
wsL.merge_cells("A1:R1")
wsL["A1"].font = font(14, True, WHITE); wsL["A1"].fill = fill(NAVY)
wsL["A1"].alignment = valign("left", "center", False, 1)
wsL.row_dimensions[1].height = 30
wsL["A2"] = ("Dropdown sources for the entire workbook. Add or rename items here and every dropdown "
             "updates automatically. Do not delete a block's header row.")
wsL.merge_cells("A2:R2")
wsL["A2"].font = font(9, False, MUTED, True); wsL["A2"].alignment = valign("left", "center", False, 1)

def block_header(ws, cell_range, text, color=NAVY_2):
    first = cell_range.split(":")[0]
    a, b = cell_range.split(":")
    if a != b:
        ws.merge_cells(cell_range)
    c = ws[first]; c.value = text
    c.font = font(9, True, WHITE); c.fill = fill(color)
    c.alignment = valign("center", "center")
    ws.row_dimensions[c.row].height = 22

# Offense master
block_header(wsL, "A4:D4", "OFFENSE MASTER  (edit freely - dropdown + auto-category + legal basis)")
for i, h in enumerate(["OFFENSE / VIOLATION", "CATEGORY (auto)", "LEGAL BASIS (auto)", "USED IN DROPDOWN"], start=1):
    c = wsL.cell(row=5, column=i, value=h)
    c.font = font(8, True, NAVY); c.fill = fill(BG); c.alignment = valign("center", "center", True)
    c.border = box()
row = 6
for o in offenses:
    wsL.cell(row=row, column=1, value=o[0]).font = font(9)
    wsL.cell(row=row, column=2, value=o[1]).font = font(9, False, MUTED)
    wsL.cell(row=row, column=3, value=o[2]).font = font(9, False, MUTED)
    wsL.cell(row=row, column=4, value="Yes").font = font(9, False, MUTED)
    row += 1
OFF_LAST = row - 1
for r in range(6, OFF_LAST + 1):
    for col in range(1, 5):
        wsL.cell(row=r, column=col).border = box("EDF1F6")

# helper blocks (side by side)
def simple_block(top_row, col, title, items, header, width=None, color=NAVY_2, extra=None):
    """items: list of values; extra: list of lists for additional columns"""
    c1 = get_column_letter(col)
    c2 = get_column_letter(col + 1)
    block_header(wsL, f"{c1}{top_row}:{c2 if extra else c1}{top_row}", title, color)
    hc = wsL.cell(row=top_row + 1, column=col, value=header)
    hc.font = font(8, True, NAVY); hc.fill = fill(BG); hc.alignment = valign("center", "center", True); hc.border = box()
    if extra:
        hd = wsL.cell(row=top_row + 1, column=col + 1, value=extra[0])
        hd.font = font(8, True, NAVY); hd.fill = fill(BG); hd.alignment = valign("center", "center", True); hd.border = box()
    for i, it in enumerate(items):
        c = wsL.cell(row=top_row + 2 + i, column=col, value=it)
        c.font = font(9); c.border = box("EDF1F6")
        if extra:
            v = extra[1][i]
            d = wsL.cell(row=top_row + 2 + i, column=col + 1, value=v)
            d.font = font(9, False, MUTED); d.border = box("EDF1F6")
            d.alignment = valign("center", "center")
    return top_row + 2, top_row + 1 + len(items)

CAT_S, CAT_E = simple_block(4, 6, "CATEGORIES", categories, "OFFENSE CATEGORY")
STG_S, STG_E = simple_block(4, 8, "CASE STAGES", [s[0] for s in stages], "CASE STAGE",
                            extra=("PROGRESS %", [s[1] for s in stages]))
PRI_S, PRI_E = simple_block(4, 10, "PRIORITY", priorities, "PRIORITY")
STA_S, STA_E = simple_block(4, 11, "CASE STATUS", statuses, "CASE STATUS")
OUT_S, OUT_E = simple_block(4, 12, "OUTCOME", outcomes, "OUTCOME")
CRT_S, CRT_E = simple_block(4, 13, "COURTS / FORUMS", courts, "COURT / FORUM")
CUS_S, CUS_E = simple_block(4, 14, "CUSTODY / BAIL", custody, "CUSTODY STATUS")
PER_S, PER_E = simple_block(4, 15, "PERSONNEL", personnel, "PROSECUTOR / INVESTIGATOR")
BRG_S, BRG_E = simple_block(4, 16, "BARANGAYS (rename to actual)", barangays, "BARANGAY / LOCATION")
HTY_S, HTY_E = simple_block(4, 17, "HEARING TYPE", hear_types, "HEARING TYPE")
HRS_S, HRS_E = simple_block(4, 18, "HEARING RESULT", hear_results, "HEARING RESULT")

# Document list block lower
DOC_ROW = OFF_LAST + 3          # placed below the offense master so blocks never overlap
block_header(wsL, f"A{DOC_ROW}:B{DOC_ROW}", "DOCUMENT CHECKLIST ITEMS (used as column headers in the register)")
for i, d in enumerate(docs):
    c = wsL.cell(row=DOC_ROW + 1 + i, column=1, value=d); c.font = font(9); c.border = box("EDF1F6")
    c2 = wsL.cell(row=DOC_ROW + 1 + i, column=2, value="Y / N / N-A"); c2.font = font(9, False, MUTED)
    c2.alignment = valign("center", "center"); c2.border = box("EDF1F6")

# Settings block
SET_ROW = DOC_ROW + len(docs) + 3
block_header(wsL, f"A{SET_ROW}:D{SET_ROW}", "OFFICE PROFILE AND SETTINGS  (feeds the dashboard header and ageing rules)")
settings = [
    ("Local Government Unit (LGU)", "Local Government Unit of San Isidro", "text"),
    ("Province / District", "Province of Nueva Ecija", "text"),
    ("Office", "Office of the City Prosecutor", "text"),
    ("Prepared by / Designation", "Case Monitoring Officer", "text"),
    ("Contact Number / Email", "(044) 000-0000  |  monitoring@lgu.gov.ph", "text"),
    ("Reporting Period Label", "CY 2026 - All Criminal Cases", "text"),
    ("Ageing: Monitor threshold (days)", 90, "num"),
    ("Ageing: For Follow-up threshold (days)", 180, "num"),
    ("Ageing: Critical threshold (days)", 365, "num"),
    ("Upcoming hearing window (days)", 7, "num"),
]
wsL.merge_cells(f"A{SET_ROW + 1}:B{SET_ROW + 1}")
h1 = wsL.cell(row=SET_ROW + 1, column=1, value="SETTING")
h1.font = font(8, True, NAVY); h1.fill = fill(BG); h1.alignment = valign("center", "center", False, 1)
h1.border = box()
h2 = wsL.cell(row=SET_ROW + 1, column=3, value="VALUE  (edit these)")
h2.font = font(8, True, NAVY); h2.fill = fill(BG); h2.alignment = valign("center", "center")
h2.border = box()

SETTINGS_CELLS = {}
for i, (label, value, kind) in enumerate(settings):
    r = SET_ROW + 2 + i
    wsL.merge_cells(f"A{r}:B{r}")
    lc = wsL.cell(row=r, column=1, value=label); lc.font = font(9, True, NAVY)
    lc.fill = fill(BG); lc.border = box(); lc.alignment = valign("left", "center", False, 1)
    vc = wsL.cell(row=r, column=3, value=value); vc.font = font(9, True, GOLD if kind == "text" else NAVY)
    vc.border = box(); vc.fill = fill(GOLD_L if kind == "text" else WHITE)
    vc.alignment = valign("center", "center")
    if kind == "num":
        vc.number_format = "0"
    SETTINGS_CELLS[label] = f"Lists!$C${r}"

LGU_CELL = SETTINGS_CELLS["Local Government Unit (LGU)"]
PROV_CELL = SETTINGS_CELLS["Province / District"]
OFFICE_CELL = SETTINGS_CELLS["Office"]
PREP_CELL = SETTINGS_CELLS["Prepared by / Designation"]
PERIOD_CELL = SETTINGS_CELLS["Reporting Period Label"]
T_MON = SETTINGS_CELLS["Ageing: Monitor threshold (days)"]
T_FUP = SETTINGS_CELLS["Ageing: For Follow-up threshold (days)"]
T_CRIT = SETTINGS_CELLS["Ageing: Critical threshold (days)"]
T_WIN = SETTINGS_CELLS["Upcoming hearing window (days)"]

set_widths(wsL, {"A": 46, "B": 34, "C": 46, "D": 14, "E": 3,
                 "F": 34, "G": 12, "H": 26, "I": 12, "J": 12, "K": 20, "L": 19,
                 "M": 27, "N": 26, "O": 24, "P": 26, "Q": 24, "R": 28})
wsL.sheet_view.zoomScale = 90

# ================================================================ CASE REGISTER
wsR = wb.create_sheet("Case Register")
wsR.sheet_properties.tabColor = NAVY

REG_FIRST = 3
REG_LAST = 602          # capacity: 600 cases
REG_HDR = 2

BANDS = {
    "id":    ("CASE IDENTIFICATION", NAVY),
    "prog":  ("CASE PROGRESS, COURT AND ASSIGNMENT", NAVY_2),
    "dates": ("KEY DATES", NAVY_3),
    "cust":  ("CUSTODY AND BAIL", NAVY_4),
    "docs":  ("DOCUMENT COMPLETENESS  (Y = on file, N = missing, N/A = not applicable)", GOLD),
    "calc":  ("AUTO-CALCULATED  (do not type here)", CALC),
    "notes": ("", CALC),
    "help":  ("", CALC),
}

COLS = [
    ("case_no",      "CASE NO.",                                           "id",   15, "text"),
    ("date_rec",     "DATE RECORDED",                                      "id",   12, "date"),
    ("title",        "CASE TITLE (People of the Philippines v. ...)",       "id",   34, "textw"),
    ("complainant",  "COMPLAINANT / VICTIM",                               "id",   24, "textw"),
    ("respondent",   "RESPONDENT / ACCUSED",                               "id",   24, "textw"),
    ("offense",      "OFFENSE / VIOLATION",                                "id",   30, "choicew"),
    ("category",     "OFFENSE CATEGORY",                                   "id",   26, "auto"),
    ("stage",        "CASE STAGE",                                         "prog", 20, "choice"),
    ("stage_since",  "STAGE SINCE",                                        "prog", 12, "date"),
    ("priority",     "PRIORITY",                                           "prog", 10, "choicec"),
    ("status",       "CASE STATUS",                                        "prog", 14, "choicec"),
    ("outcome",      "OUTCOME",                                            "prog", 16, "choicec"),
    ("court",        "COURT / FORUM",                                      "prog", 22, "choice"),
    ("docket",       "COURT DOCKET NO.",                                   "prog", 15, "text"),
    ("personnel",    "ASSIGNED PROSECUTOR / INVESTIGATOR",                 "prog", 23, "choice"),
    ("barangay",     "BARANGAY / LOCATION",                                "prog", 18, "choice"),
    ("incident",     "DATE OF INCIDENT",                                   "dates",12, "date"),
    ("filed",        "DATE FILED / REFERRED",                              "dates",12, "date"),
    ("filed_court",  "DATE FILED IN COURT",                                "dates",12, "date"),
    ("next_hearing", "NEXT HEARING",                                       "dates",12, "date"),
    ("last_hearing", "LAST HEARING",                                       "dates",12, "date"),
    ("closed",       "DATE CLOSED / RESOLVED",                             "dates",12, "date"),
    ("custody",      "CUSTODY / BAIL STATUS",                              "cust", 20, "choice"),
    ("detained_on",  "DATE OF DETENTION",                                  "cust", 12, "date"),
    ("doc1",         "1. Affidavit of Complaint",                          "docs", 13, "yn"),
    ("doc2",         "2. Sworn Statement(s)",                              "docs", 13, "yn"),
    ("doc3",         "3. Police / Incident Report",                        "docs", 13, "yn"),
    ("doc4",         "4. Evidence and Exhibits",                           "docs", 13, "yn"),
    ("doc5",         "5. Resolution / Information",                        "docs", 13, "yn"),
    ("doc6",         "6. Court Order / Docket",                            "docs", 13, "yn"),
    ("doc7",         "7. Bail / Bond Records",                             "docs", 13, "yn"),
    ("doc8",         "8. Judgment / Order",                                "docs", 13, "yn"),
    ("days_pending", "DAYS PENDING (TOTAL)",                               "calc", 11, "autonum"),
    ("days_stage",   "DAYS IN CURRENT STAGE",                              "calc", 11, "autonum"),
    ("ageing",       "STAGE AGEING",                                       "calc", 14, "autoc"),
    ("progress",     "CASE PROGRESS",                                      "calc", 11, "autopct"),
    ("docs_pct",     "DOCUMENTS COMPLETE",                                 "calc", 11, "autopct"),
    ("detention_days","DAYS IN CUSTODY",                                   "calc", 11, "autonum"),
    ("remarks",      "REMARKS / NEXT ACTION",                              "notes",40, "textw"),
    ("key_age",      "HELPER - AGE KEY",                                   "help", 8,  "auto"),
    ("key_hearing",  "HELPER - HEARING KEY",                               "help", 8,  "auto"),
    ("key_rank",     "HELPER - RANK KEY",                                  "help", 8,  "auto"),
]

COL = {k: get_column_letter(i + 1) for i, (k, *_rest) in enumerate(COLS)}
LAST_COL = get_column_letter(len(COLS))
LAST_COL_IDX = len(COLS)
R = lambda key, r=REG_FIRST: f"${COL[key]}{r}"

# --- band row (row 1)
wsR.row_dimensions[1].height = 20
groups = []
for k, h, grp, w, kind in COLS:
    if not groups or groups[-1][0] != grp:
        groups.append([grp, COL[k], COL[k]])
    else:
        groups[-1][2] = COL[k]

for grp, c1, c2 in groups:
    if c1 != c2:
        wsR.merge_cells(f"{c1}1:{c2}1")
    c = wsR[f"{c1}1"]
    label, color = BANDS[grp]
    c.value = label
    c.fill = fill(color)
    c.font = font(8.5, True, NAVY if grp == "docs" else WHITE)
    c.alignment = valign("center", "center")
    c.border = Border(left=Side(style="thin", color=WHITE), right=Side(style="thin", color=WHITE))

# --- header row (row 2)
wsR.row_dimensions[REG_HDR].height = 46
for k, h, grp, w, kind in COLS:
    c = wsR[f"{COL[k]}{REG_HDR}"]
    c.value = h
    c.fill = fill(NAVY if grp != "docs" else "B08F3C")
    c.font = font(8, True, WHITE)
    c.alignment = valign("center", "center", True)
    c.border = box(WHITE)
    wsR.column_dimensions[COL[k]].width = w

# --- data rows
numfmt_date = "dd mmm yyyy"
for r in range(REG_FIRST, REG_LAST + 1):
    wsR.row_dimensions[r].height = 27
    for k, h, grp, w, kind in COLS:
        c = wsR[f"{COL[k]}{r}"]
        c.border = Border(bottom=side("thin", "E4EAF1"), right=side("thin", "EDF1F6"))
        c.font = font(9)
        c.alignment = valign("left", "center")
        if kind == "date":
            c.number_format = numfmt_date
            c.alignment = valign("center", "center")
        elif kind in ("textw", "choicew"):
            c.alignment = valign("left", "center", True)
        elif kind == "choicec":
            c.alignment = valign("center", "center", True)
        elif kind == "yn":
            c.alignment = valign("center", "center")
            c.font = font(9, True, "374151")
        elif kind in ("auto", "autonum", "autopct", "autoc"):
            c.fill = fill("EFF2F6")
            c.font = font(9, italic=False, color="374151")
            if kind == "autonum":
                c.number_format = "0"; c.alignment = valign("center", "center")
            elif kind == "autopct":
                c.number_format = "0%"; c.alignment = valign("center", "center")
            elif kind == "autoc":
                c.alignment = valign("center", "center")
            else:
                c.alignment = valign("left", "center", True)

# --- formulas (written for every row of the register) -------------------
C = COL
def put(key, builder):
    for r in range(REG_FIRST, REG_LAST + 1):
        wsR[f"{C[key]}{r}"] = builder(r)

put("category", lambda r: (
    f'=IF({C["case_no"]}{r}="","",IFERROR(INDEX(OffCat,MATCH({C["offense"]}{r},OffList,0)),"Unlisted"))'))

put("days_pending", lambda r: (
    f'=IF({C["case_no"]}{r}="","",IF({C["closed"]}{r}<>"",{C["closed"]}{r}-{C["date_rec"]}{r},TODAY()-{C["date_rec"]}{r}))'))

put("days_stage", lambda r: (
    f'=IF(OR({C["case_no"]}{r}="",{C["stage_since"]}{r}=""),"",'
    f'MAX(0,IF({C["closed"]}{r}<>"",{C["closed"]}{r}-{C["stage_since"]}{r},TODAY()-{C["stage_since"]}{r})))'))

put("ageing", lambda r: (
    f'=IF({C["case_no"]}{r}="","",'
    f'IF(OR({C["status"]}{r}="Closed",{C["closed"]}{r}<>""),"Closed",'
    f'IF({C["days_stage"]}{r}>=AgeCrit,"Critical",'
    f'IF({C["days_stage"]}{r}>=AgeFu,"For Follow-up",'
    f'IF({C["days_stage"]}{r}>=AgeMonitor,"Monitor","Within Standard")))))'))

put("progress", lambda r: (
    f'=IF({C["case_no"]}{r}="","",IFERROR(INDEX(StgPct,MATCH({C["stage"]}{r},StgList,0)),0))'))

put("docs_pct", lambda r: (
    f'=IF({C["case_no"]}{r}="","",IFERROR(COUNTIF({C["doc1"]}{r}:{C["doc8"]}{r},"Y")/'
    f'(COUNTIF({C["doc1"]}{r}:{C["doc8"]}{r},"Y")+COUNTIF({C["doc1"]}{r}:{C["doc8"]}{r},"N")),0))'))

put("detention_days", lambda r: (
    f'=IF(OR({C["case_no"]}{r}="",{C["detained_on"]}{r}=""),"",'
    f'MAX(0,IF({C["closed"]}{r}<>"",{C["closed"]}{r},TODAY())-{C["detained_on"]}{r}))'))

put("key_age", lambda r: (
    f'=IF({C["case_no"]}{r}="","",LOWER({C["ageing"]}{r})&"|"&{C["days_stage"]}{r}&"|"&{C["progress"]}{r}&"|"&{C["docs_pct"]}{r})'))

put("key_hearing", lambda r: (
    f'=IF(OR({C["case_no"]}{r}="",{C["next_hearing"]}{r}=""),"",{C["next_hearing"]}{r}-TODAY())'))

put("key_rank", lambda r: (
    f'=IF(OR({C["case_no"]}{r}="",{C["closed"]}{r}<>""),"",'
    f'(TODAY()-{C["date_rec"]}{r})+ROW()/1000000)'))

# --- sample records (26) : replace or delete, formulas keep working -------
SAMPLES = [
 ("CC-2026-0001", D(-182), "People of the Philippines v. Rolando Dela Cruz y Santos",
  "Elena Dela Cruz", "Rolando Dela Cruz y Santos", "Murder", "Trial", D(-45), "High", "Active", "Pending",
  "Regional Trial Court", "RTC-2026-1123", "Atty. R. Villanueva", "Barangay 3", D(-210), D(-182), D(-150),
  D(4), D(-26), None, "In Custody - No Bail", D(-182), "YYYYYNYN",
  "Prosecution presented 3rd witness; firearm evidence admitted."),
 ("CC-2026-0002", D(-140), "People of the Philippines v. Jayson Mercado y Lim",
  "PDEA Regional Office", "Jayson Mercado y Lim", "Illegal Sale of Dangerous Drugs", "Pre-Trial", D(-70),
  "High", "Active", "Pending", "Regional Trial Court", "RTC-2026-1188", "Atty. M. Fernandez", "Barangay 7",
  D(-160), D(-140), D(-120), D(12), D(-18), None, "In Custody - No Bail", D(-140), "YYYYYYYN",
  "Compliance with Sec. 21 RA 9165 stipulation; pre-trial order to follow."),
 ("CC-2026-0003", D(-80), "People of the Philippines v. Arnel Bautista y Reyes",
  "Ma. Cristina Lim", "Arnel Bautista y Reyes", "Theft", "Arraignment", D(-30), "Medium", "Active", "Pending",
  "Metropolitan Trial Court", "MeTC-2026-0455", "Atty. J. Aquino", "Barangay 1", D(-95), D(-80), D(-45),
  D(-6), D(-37), None, "Released on Bail", None, "YYNYYNYN",
  "Arraignment reset twice; notice of hearing re-served on counsel."),
 ("CC-2026-0004", D(-40), "People of the Philippines v. Danilo Perez y Cruz",
  "AAA (name withheld)", "Danilo Perez y Cruz", "Violation of VAWC Act", "Filed in Court", D(-10), "High",
  "Active", "Pending", "Municipal Trial Court", "MTC-2026-0672", "Atty. C. Bautista", "Barangay 5", D(-70),
  D(-40), D(-10), D(20), None, None, "Released on Bail", None, "YYNYYYYN",
  "Warrant of arrest served; court ordered psychiatric evaluation."),
 ("CC-2026-0005", D(-260), "People of the Philippines v. Victoriano Salazar y Ang",
  "BBB (name withheld)", "Victoriano Salazar y Ang", "Rape", "Trial", D(-120), "High", "Active", "Pending",
  "Regional Trial Court", "RTC-2026-0901", "Atty. S. Navarro", "Barangay 2", D(-300), D(-260), D(-240),
  D(2), D(-28), None, "In Custody - No Bail", D(-300), "YYYYYYYN",
  "Victim testimony completed; counsel to file formal offer of evidence."),
 ("CC-2026-0006", D(-25), "People of the Philippines v. Lilia Manalo y Diaz",
  "Fernando Uy", "Lilia Manalo y Diaz", "Estafa", "Preliminary Investigation", D(-25), "Medium",
  "For Action", "Pending", "Office of the Prosecutor", "", "Investigator P. Reyes", "Barangay 9", D(-60),
  D(-25), None, None, None, None, "N/A (Not in Custody)", None, "YYNYYNNN",
  "Awaiting counter-affidavit from respondent; clarify 15-day reglementary period."),
 ("CC-2026-0007", D(-300), "People of the Philippines v. Emmanuel Garcia y Tan",
  "Rodolfo Santos", "Emmanuel Garcia y Tan", "Serious Physical Injuries", "Judgment Rendered", D(-35),
  "Medium", "Closed", "Convicted", "Municipal Trial Court", "MTC-2025-0912", "Atty. J. Aquino", "Barangay 4",
  D(-320), D(-300), D(-280), None, D(-40), D(-30), "Released on Bail", None, "YYYYYYYY",
  "Accused found guilty; penalty imposed. Certified true copy on file."),
 ("CC-2026-0008", D(-170), "People of the Philippines v. Rodelito Ramos y Bautista",
  "Ramon Ramos", "Rodelito Ramos y Bautista", "Reckless Imprudence Resulting in Homicide", "Trial", D(-55),
  "Medium", "Active", "Pending", "Municipal Trial Court", "MTC-2026-0210", "Atty. R. Villanueva", "Barangay 6",
  D(-200), D(-170), D(-150), D(9), D(-20), None, "Released on Bail", None, "YYNYYNYY",
  "Private prosecutor presented LTO records; next setting for defense evidence."),
 ("CC-2026-0009", D(-380), "People of the Philippines v. Henry Ocampo y Villar",
  "Josefina Tan", "Henry Ocampo y Villar", "Qualified Theft", "Submitted for Decision", D(-205), "High",
  "Active", "Pending", "Regional Trial Court", "RTC-2026-0777", "Atty. M. Fernandez", "Barangay 8", D(-400),
  D(-380), D(-350), None, D(-210), None, "Released on Bail", None, "YYYYYYYN",
  "Case submitted for decision. Follow up with branch clerk of court."),
 ("CC-2025-0010", D(-560), "People of the Philippines v. Armando Sison y Uy",
  "Security Bank Corporation", "Armando Sison y Uy", "Robbery", "Trial", D(-410), "High", "Active", "Pending",
  "Regional Trial Court", "RTC-2025-1450", "Atty. S. Navarro", "Barangay 3", D(-600), D(-560), D(-520),
  D(15), D(-35), None, "In Custody - No Bail", D(-600), "YYYYNNYN",
  "Long-pending. Missing bail records and judgment docs; request from RTC records."),
 ("CC-2025-0011", D(-150), "People of the Philippines v. Carlos Miguel y Fernandez",
  "BBB (name withheld)", "Carlos Miguel y Fernandez", "Violation of VAWC Act", "Filed in Court", D(-90),
  "High", "Active", "Pending", "Municipal Trial Court", "MTC-2025-0888", "Atty. C. Bautista", "Barangay 10",
  D(-200), D(-150), D(-90), D(15), D(-30), None, "Released on Bail", None, "YYNYYNYN",
  "Temporary protection order issued and extended for 30 days."),
 ("CC-2025-0012", D(-120), "People of the Philippines v. Ismael Tupaz y Garcia",
  "PNP Station 2", "Ismael Tupaz y Garcia", "Illegal Possession of Firearms", "Arraignment", D(-60),
  "Medium", "Active", "Pending", "Regional Trial Court", "RTC-2025-1290", "Atty. R. Villanueva",
  "Barangay 7", D(-180), D(-120), D(-75), D(7), D(-15), None, "In Custody - With Bail", D(-180),
  "YYYYYNYN", "Bail approved at PHP 120,000. Arraignment set with counsel present."),
 ("CC-2025-0013", D(-250), "People of the Philippines v. Teresita Cruz y Dela Peña",
  "DSWD Field Office", "Teresita Cruz y Dela Peña", "Child Abuse", "Pre-Trial", D(-150), "High", "Active",
  "Pending", "Regional Trial Court", "RTC-2025-1102", "Atty. M. Fernandez", "Barangay 5", D(-300),
  D(-250), D(-200), D(25), D(-25), None, "Released on Bail", None, "YYYYNYNY",
  "Child witness presented via video conference; pre-trial partly terminated."),
 ("CC-2025-0014", D(-350), "People of the Philippines v. Nicanor Abad y Lim",
  "Metro Hardware Supply", "Nicanor Abad y Lim", "Anti-Fencing", "Archived", D(-70), "Low", "Closed",
  "Dismissed", "Metropolitan Trial Court", "MeTC-2025-0760", "Atty. J. Aquino", "Barangay 2", D(-400),
  D(-350), D(-300), None, None, D(-60), "At Large", None, "YYNYNNNN",
  "Dismissed on demurrer to evidence. Prosecutor did not appeal."),
 ("CC-2025-0015", D(-5), "People of the Philippines v. Joel Manalo y Reyes",
  "Barangay 6 Council", "Joel Manalo y Reyes", "Illegal Gambling", "For Filing", D(-5), "Low", "For Action",
  "Pending", "Office of the Prosecutor", "", "Investigator L. Gomez", "Barangay 6", D(-12), D(-5), None,
  None, None, None, "N/A (Not in Custody)", None, "YNNNNNNN",
  "New referral. Complete sworn statements and police blotter before filing."),
 ("CC-2025-0016", D(-70), "People of the Philippines v. Kristine Joy Salcedo",
  "Marvin Chua", "Kristine Joy Salcedo", "Cyber Libel", "Filed in Court", D(-35), "Medium", "Active",
  "Pending", "Regional Trial Court", "RTC-2025-1401", "Atty. S. Navarro", "Barangay 1", D(-120), D(-70),
  D(-35), D(35), None, None, "Released on Bail", None, "YYNYYNNN",
  "Preservation order on social media posts granted; awaiting URL certification."),
 ("CC-2025-0017", D(-220), "People of the Philippines v. Bonifacio Alonzo y Perez",
  "BFAR Municipal Office", "Bonifacio Alonzo y Perez", "Illegal Fishing", "Trial", D(-125), "Medium",
  "Active", "Pending", "Municipal Trial Court", "MTC-2025-0819", "Investigator P. Reyes", "Poblacion",
  D(-250), D(-220), D(-190), D(19), D(-30), None, "Released on Recognizance", None, "YYNYYNYN",
  "Boat and gear forfeited in favor of government; trial on the merits ongoing."),
 ("CC-2025-0018", D(-250), "People of the Philippines v. Alfonso Trinidad",
  "Barangay 4 Residents", "Alfonso Trinidad", "Violation of Local Ordinance", "Archived", D(-200), "Low",
  "Closed", "Archived / Referred", "Katarungang Pambarangay", "", "Atty. C. Bautista", "Barangay 4",
  D(-300), D(-250), None, None, None, D(-200), "N/A (Not in Custody)", None, "YNYNNNNN",
  "Settled before the Lupong Tagapamayapa; amicable settlement executed."),
 ("CC-2025-0019", D(-850), "People of the Philippines v. Federico Ledesma y Sy",
  "Heirs of Lorenzo Uy", "Federico Ledesma y Sy", "Homicide", "Appeal / Review", D(-300), "High", "Active",
  "Pending", "Court of Appeals", "CA-GR-2025-1234", "Atty. R. Villanueva", "Barangay 9", D(-900),
  D(-850), D(-800), None, D(-300), None, "In Custody - No Bail", D(-900), "YYYYYYYN",
  "Appeal pending. Manifestation with compliance filed; monitor CA resolution."),
 ("CC-2025-0020", D(-300), "People of the Philippines v. Lauro Domingo y Aquino",
  "Pedro Aquino", "Lauro Domingo y Aquino", "Frustrated Murder", "Trial", D(-85), "High", "Active",
  "Pending", "Regional Trial Court", "RTC-2025-1355", "Atty. M. Fernandez", "Barangay 3", D(-350),
  D(-300), D(-260), D(18), D(-22), None, "In Custody - No Bail", D(-300), "YYYYYNYN",
  "Medico-legal report admitted; defense to present rebuttal witness."),
 ("CC-2026-0021", D(-40), "People of the Philippines v. Marilou Sanchez y Reyes",
  "Inter-Agency Council Against Trafficking", "Marilou Sanchez y Reyes", "Trafficking in Persons",
  "Preliminary Investigation", D(-40), "High", "For Action", "Pending", "Office of the Prosecutor", "",
  "Atty. S. Navarro", "Barangay 8", D(-90), D(-40), None, None, None, None, "Released on Bail", None,
  "YYNYYNNN", "PI ongoing; victim-witness protection referral requested from DOJ."),
 ("CC-2026-0022", D(-100), "People of the Philippines v. Ruben Castillo y Tan",
  "Golden Grain Traders", "Ruben Castillo y Tan", "Bouncing Checks", "Filed in Court", D(-20), "Low",
  "Active", "Pending", "Metropolitan Trial Court", "MeTC-2026-0511", "Atty. J. Aquino", "Barangay 2",
  D(-150), D(-100), D(-20), D(5), None, None, "Released on Bail", None, "YYNYYNYN",
  "Court directed parties to explore compromise; reset for status conference."),
 ("CC-2026-0023", D(-200), "People of the Philippines v. Edgar Villamor y Cruz",
  "Bureau of Fire Protection", "Edgar Villamor y Cruz", "Arson", "Trial", D(-65), "High", "Active",
  "Pending", "Regional Trial Court", "RTC-2026-1002", "Atty. R. Villanueva", "Barangay 6", D(-240),
  D(-200), D(-170), D(40), D(-12), None, "In Custody - With Bail", D(-240), "YYYYYNYN",
  "Fire investigator testimony taken; next setting for documentary exhibits."),
 ("CC-2026-0024", D(-240), "People of the Philippines v. Rebecca Lim y Ong",
  "Nena Fernandez", "Rebecca Lim y Ong", "Illegal Recruitment", "Judgment Rendered", D(-50), "Medium",
  "Closed", "Amicably Settled", "Regional Trial Court", "RTC-2025-0999", "Atty. C. Bautista", "Barangay 10",
  D(-280), D(-240), D(-210), None, None, D(-45), "At Large", None, "YYYYYYYY",
  "Accused jumped bail; conviction in absentia recorded, records endorsed to court."),
 ("CC-2026-0025", D(-75), "People of the Philippines v. Rodrigo Basco y Naval",
  "CCC (name withheld)", "Rodrigo Basco y Naval", "Acts of Lasciviousness", "Arraignment", D(-28), "High",
  "Active", "Pending", "Municipal Trial Court", "MTC-2026-0701", "Atty. M. Fernandez", "Barangay 5",
  D(-110), D(-75), D(-40), D(11), D(-28), None, "In Custody - No Bail", D(-75), "YYNYYNYN",
  "Accused entered not guilty plea; pre-trial conference scheduled."),
 ("CC-2026-0026", D(-650), "People of the Philippines v. San Isidro Agro-Industrial Corp. (resp. reps.)",
  "DENR - EMB Region III", "San Isidro Agro-Industrial Corp.", "Violation of Clean Water Act",
  "Trial", D(-520), "Low", "Dormant", "Pending", "Regional Trial Court", "RTC-2024-0512",
  "Investigator L. Gomez", "Outside Jurisdiction", D(-700), D(-650), D(-600), None, D(-520), None,
  "At Large", None, "YYNYNNNN", "Dormant - accused cannot be located. Recommend archival review."),
]

for i, s in enumerate(SAMPLES):
    r = REG_FIRST + i
    vals = {
        "case_no": s[0], "date_rec": s[1], "title": s[2], "complainant": s[3], "respondent": s[4],
        "offense": s[5], "stage": s[6], "stage_since": s[7], "priority": s[8], "status": s[9],
        "outcome": s[10], "court": s[11], "docket": s[12] or None, "personnel": s[13],
        "barangay": s[14], "incident": s[15], "filed": s[16], "filed_court": s[17],
        "next_hearing": s[18], "last_hearing": s[19], "closed": s[20], "custody": s[21],
        "detained_on": s[22], "remarks": s[24],
    }
    for j, ch in enumerate(s[23]):
        vals[f"doc{j+1}"] = {"Y": "Y", "N": "N", "-": "N/A"}.get(ch, "N")
    for k, v in vals.items():
        if v is not None:
            wsR[f"{C[k]}{r}"] = v
    # zebra-free: rely on table style

# --- data validation -----------------------------------------------------
DV_RANGE_LAST = REG_LAST
def add_list_dv(key, name, title, msg):
    dv = DataValidation(type="list", formula1=name, allow_blank=True, showErrorMessage=True)
    dv.errorTitle = "Invalid entry"
    dv.error = msg
    dv.promptTitle = title
    dv.prompt = "Choose from the list. To add new options, edit the Lists sheet."
    dv.showInputMessage = True
    wsR.add_data_validation(dv)
    dv.add(f"{C[key]}{REG_FIRST}:{C[key]}{REG_LAST}")

add_list_dv("offense", "OffList", "Offense / Violation", "Pick an offense from the List sheet, or add it there first.")
add_list_dv("stage", "StgList", "Case stage", "Pick a case stage from the List sheet.")
add_list_dv("priority", "PriList", "Priority", "Choose High, Medium or Low.")
add_list_dv("status", "StatusList", "Case status", "Pick a case status from the List sheet.")
add_list_dv("outcome", "OutcomeList", "Outcome", "Pick an outcome from the List sheet.")
add_list_dv("court", "CourtList", "Court / forum", "Pick a court or forum from the List sheet.")
add_list_dv("custody", "CustodyList", "Custody status", "Pick a custody status from the List sheet.")
add_list_dv("personnel", "PersonnelList", "Assigned personnel", "Pick an assignee from the List sheet.")
add_list_dv("barangay", "BrgyList", "Barangay", "Pick a barangay from the List sheet.")

date_dv = DataValidation(type="date", operator="between", formula1="DATE(2000,1,1)",
                           formula2="DATE(2100,12,31)", allow_blank=True, showErrorMessage=True,
                           errorTitle="Invalid date", error="Enter a real date (e.g. 21/10/2026).")
wsR.add_data_validation(date_dv)
for dk in ("date_rec", "stage_since", "incident", "filed", "filed_court", "next_hearing",
           "last_hearing", "closed", "detained_on"):
    date_dv.add(f"{C[dk]}{REG_FIRST}:{C[dk]}{REG_LAST}")

doc_dv = DataValidation(type="list", formula1='"Y,N,N/A"', allow_blank=True, showErrorMessage=True)
doc_dv.errorTitle = "Y / N / N/A only"
doc_dv.error = "Enter Y (on file), N (missing) or N/A (not applicable)."
doc_dv.promptTitle = "Document checklist"
doc_dv.prompt = "Y = on file, N = missing, N/A = not applicable."
doc_dv.showInputMessage = True
wsR.add_data_validation(doc_dv)
doc_dv.add(f"{C['doc1']}{REG_FIRST}:{C['doc8']}{REG_LAST}")

# --- conditional formatting ----------------------------------------------
def cf_range(key, last=REG_LAST):
    return f"{C[key]}{REG_FIRST}:{C[key]}{last}"

age_col = cf_range("ageing")
wsR.conditional_formatting.add(age_col, FormulaRule(formula=[f'{C["ageing"]}{REG_FIRST}="Closed"'],
    fill=fill("E8ECF1"), font=font(8.5, True, MUTED), stopIfTrue=True))
wsR.conditional_formatting.add(age_col, FormulaRule(formula=[f'{C["ageing"]}{REG_FIRST}="Critical"'],
    fill=fill(RED_L), font=font(8.5, True, RED), border=box("F0B4AE"), stopIfTrue=True))
wsR.conditional_formatting.add(age_col, FormulaRule(formula=[f'{C["ageing"]}{REG_FIRST}="For Follow-up"'],
    fill=fill(AMBER_L), font=font(8.5, True, AMBER), border=box("EBCF9A"), stopIfTrue=True))
wsR.conditional_formatting.add(age_col, FormulaRule(formula=[f'{C["ageing"]}{REG_FIRST}="Monitor"'],
    fill=fill(BLUE_L), font=font(8.5, True, BLUE), stopIfTrue=True))
wsR.conditional_formatting.add(age_col, FormulaRule(formula=[f'{C["ageing"]}{REG_FIRST}="Within Standard"'],
    fill=fill(GREEN_L), font=font(8.5, True, GREEN), stopIfTrue=True))

# case number accent when case is critical / dormant
wsR.conditional_formatting.add(cf_range("case_no"), FormulaRule(
    formula=[f'AND({C["ageing"]}{REG_FIRST}="Critical")'],
    fill=fill("FFF1F0"), font=font(9.5, True, RED)))
wsR.conditional_formatting.add(cf_range("case_no"), FormulaRule(
    formula=[f'{C["ageing"]}{REG_FIRST}="Closed"'],
    font=font(9.5, True, MUTED)))

# priority
pri = cf_range("priority")
wsR.conditional_formatting.add(pri, FormulaRule(formula=[f'{C["priority"]}{REG_FIRST}="High"'],
    fill=fill(RED_L), font=font(9, True, RED)))
wsR.conditional_formatting.add(pri, FormulaRule(formula=[f'{C["priority"]}{REG_FIRST}="Medium"'],
    fill=fill(AMBER_L), font=font(9, True, AMBER)))
wsR.conditional_formatting.add(pri, FormulaRule(formula=[f'{C["priority"]}{REG_FIRST}="Low"'],
    fill=fill("EFF2F6"), font=font(9, True, MUTED)))

# status
st = cf_range("status")
wsR.conditional_formatting.add(st, FormulaRule(formula=[f'{C["status"]}{REG_FIRST}="Active"'],
    fill=fill(BLUE_L), font=font(9, True, BLUE)))
wsR.conditional_formatting.add(st, FormulaRule(formula=[f'{C["status"]}{REG_FIRST}="For Action"'],
    fill=fill(AMBER_L), font=font(9, True, AMBER)))
wsR.conditional_formatting.add(st, FormulaRule(formula=[f'OR({C["status"]}{REG_FIRST}="Dormant",{C["status"]}{REG_FIRST}="Deferred")'],
    fill=fill("F3F4F6"), font=font(9, True, "4B5563")))
wsR.conditional_formatting.add(st, FormulaRule(formula=[f'{C["status"]}{REG_FIRST}="Closed"'],
    fill=fill("E8ECF1"), font=font(9, True, MUTED)))
wsR.conditional_formatting.add(st, FormulaRule(formula=[f'{C["status"]}{REG_FIRST}="Endorsed / Transferred"'],
    fill=fill("EDE9FE"), font=font(9, True, "5B21B6")))

# outcome
oc = cf_range("outcome")
wsR.conditional_formatting.add(oc, FormulaRule(formula=[f'{C["outcome"]}{REG_FIRST}="Convicted"'],
    fill=fill(GREEN_L), font=font(9, True, GREEN)))
wsR.conditional_formatting.add(oc, FormulaRule(formula=[f'{C["outcome"]}{REG_FIRST}="Acquitted"'],
    fill=fill(RED_L), font=font(9, True, RED)))
wsR.conditional_formatting.add(oc, FormulaRule(formula=[f'{C["outcome"]}{REG_FIRST}="Dismissed"'],
    fill=fill("F3F4F6"), font=font(9, True, "4B5563")))

# hearings
nh = cf_range("next_hearing")
wsR.conditional_formatting.add(nh, FormulaRule(
    formula=[f'AND({C["key_hearing"]}{REG_FIRST}<>"",{C["key_hearing"]}{REG_FIRST}<0)'],
    fill=fill(RED), font=font(9, True, WHITE), stopIfTrue=True))
wsR.conditional_formatting.add(nh, FormulaRule(
    formula=[f'AND({C["key_hearing"]}{REG_FIRST}<>"",{C["key_hearing"]}{REG_FIRST}>=0,{C["key_hearing"]}{REG_FIRST}<=AgeWindow)'],
    fill=fill(GOLD), font=font(9, True, WHITE), stopIfTrue=True))

# documents
doc_rng = f"{C['doc1']}{REG_FIRST}:{C['doc8']}{REG_LAST}"
wsR.conditional_formatting.add(doc_rng, FormulaRule(formula=[f'{C["doc1"]}{REG_FIRST}="N"'],
    fill=fill(RED_L), font=font(9, True, RED)))
wsR.conditional_formatting.add(doc_rng, FormulaRule(formula=[f'{C["doc1"]}{REG_FIRST}="Y"'],
    fill=fill(GREEN_L), font=font(9, True, GREEN)))
wsR.conditional_formatting.add(doc_rng, FormulaRule(formula=[f'{C["doc1"]}{REG_FIRST}="N/A"'],
    font=font(9, False, "9CA3AF")))

# data bars on calculated measures
for key, color in [("days_pending", NAVY_3), ("days_stage", AMBER),
                   ("progress", GREEN), ("docs_pct", TEAL), ("detention_days", NAVY_4)]:
    wsR.conditional_formatting.add(cf_range(key),
        DataBarRule(start_type="num", start_value=0, end_type="max", color=color, showValue=True))

# --- table, panes, print ------------------------------------------------
tbl = Table(displayName="tblCases", ref=f"A{REG_HDR}:{LAST_COL}{REG_LAST}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True,
                                    showColumnStripes=False, showFirstColumn=False, showLastColumn=False)
tbl.totalsRowShown = False
wsR.add_table(tbl)

wsR.freeze_panes = "D3"
wsR.sheet_view.showGridLines = False
wsR.sheet_view.zoomScale = 85
wsR.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
wsR.print_title_rows = "1:2"
wsR.page_setup.orientation = "landscape"
wsR.page_setup.paperSize = wsR.PAPERSIZE_A3
wsR.page_setup.fitToWidth = 1
wsR.page_setup.fitToHeight = 0
wsR.print_options.horizontalCentered = True
wsR.print_area = f"A1:{COL['remarks']}{REG_LAST}"
for helper in ("key_age", "key_hearing", "key_rank"):
    wsR.column_dimensions[COL[helper]].hidden = True
wsR.oddFooter.center.text = "LGU Criminal Case Monitoring System  |  Page &P of &N"
wsR.oddFooter.center.size = 8
wsR.oddFooter.center.color = "808080"

# (usage notes live on the User Guide sheet and in each sheet header)

# --- defined names -------------------------------------------------------
def dn(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

dn("OffList", f"Lists!$A$6:$A${OFF_LAST}")
dn("OffCat", f"Lists!$B$6:$B${OFF_LAST}")
dn("StgList", f"Lists!$H${STG_S}:$H${STG_E}")
dn("StgPct", f"Lists!$I${STG_S}:$I${STG_E}")
dn("PriList", f"Lists!$J${PRI_S}:$J${PRI_E}")
dn("StatusList", f"Lists!$K${STA_S}:$K${STA_E}")
dn("OutcomeList", f"Lists!$L${OUT_S}:$L${OUT_E}")
dn("CourtList", f"Lists!$M${CRT_S}:$M${CRT_E}")
dn("CustodyList", f"Lists!$N${CUS_S}:$N${CUS_E}")
dn("PersonnelList", f"Lists!$O${PER_S}:$O${PER_E}")
dn("BrgyList", f"Lists!$P${BRG_S}:$P${BRG_E}")
dn("HearTypeList", f"Lists!$Q${HTY_S}:$Q${HTY_E}")
dn("HearResultList", f"Lists!$R${HRS_S}:$R${HRS_E}")
dn("CatList", f"Lists!$F${CAT_S}:$F${CAT_E}")
dn("LGUName", LGU_CELL)
dn("Province", PROV_CELL)
dn("OfficeName", OFFICE_CELL)
dn("PreparedBy", PREP_CELL)
dn("PeriodLabel", PERIOD_CELL)
dn("AgeMonitor", T_MON)
dn("AgeFu", T_FUP)
dn("AgeCrit", T_CRIT)
dn("AgeWindow", T_WIN)

def reg_name(name, key):
    dn(name, f"'Case Register'!${COL[key]}${REG_FIRST}:${COL[key]}${REG_LAST}")

reg_name("RegCaseNo", "case_no")
reg_name("RegTitle", "title")
reg_name("RegOffense", "offense")
reg_name("RegCat", "category")
reg_name("RegStage", "stage")
reg_name("RegPriority", "priority")
reg_name("RegStatus", "status")
reg_name("RegOutcome", "outcome")
reg_name("RegCourt", "court")
reg_name("RegPersonnel", "personnel")
reg_name("RegBrgy", "barangay")
reg_name("RegNextHearing", "next_hearing")
reg_name("RegAge", "ageing")
reg_name("RegProgress", "progress")
reg_name("RegDaysStage", "days_stage")
reg_name("RegDaysPending", "days_pending")
reg_name("RegDocsPct", "docs_pct")
reg_name("RegCustody", "custody")
reg_name("RegRankKey", "key_rank")
reg_name("RegFiled", "filed")

# ================================================================ ANALYSIS (chart data)
wsA = wb.create_sheet("Analysis")
wsA.sheet_properties.tabColor = MUTED
wsA.sheet_view.showGridLines = False
wsA["A1"] = "ANALYSIS TABLES  |  live counts that feed the dashboard charts"
wsA.merge_cells("A1:E1")
wsA["A1"].font = font(13, True, WHITE); wsA["A1"].fill = fill(NAVY)
wsA["A1"].alignment = valign("left", "center", False, 1)
wsA.row_dimensions[1].height = 28
wsA["A2"] = "Everything here is a formula. You never need to edit this sheet."
wsA.merge_cells("A2:E2"); wsA["A2"].font = font(9, False, MUTED, True)
set_widths(wsA, {"A": 4, "B": 40, "C": 12, "D": 14, "E": 30})

def ablock(top, title, header, items, counter, fmt="0"):
    c = wsA.cell(row=top, column=2, value=title)
    c.font = font(10, True, NAVY)
    h1 = wsA.cell(row=top + 1, column=2, value=header); h1.font = font(8, True, MUTED)
    h2 = wsA.cell(row=top + 1, column=3, value="COUNT"); h2.font = font(8, True, MUTED)
    h2.alignment = valign("center")
    for i, it in enumerate(items):
        r = top + 2 + i
        a = wsA.cell(row=r, column=2, value=it); a.font = font(9)
        b = wsA.cell(row=r, column=3, value=counter.format(item=it, r=r))
        b.font = font(9, True, NAVY); b.number_format = fmt
        b.alignment = valign("center")
        for cell in (a, b):
            cell.border = Border(bottom=side("thin", "EDF1F6"))
    return top + 2, top + 1 + len(items)

STG_A_S, STG_A_E = ablock(4, "CASES BY STAGE", "CASE STAGE", [s[0] for s in stages],
                          '=COUNTIF(RegStage,"{item}")')
CAT_A_S, CAT_A_E = ablock(STG_A_E + 3, "CASES BY OFFENSE CATEGORY", "CATEGORY", categories,
                          '=COUNTIF(RegCat,"{item}")')
STA_A_S, STA_A_E = ablock(CAT_A_E + 3, "CASES BY STATUS", "CASE STATUS", statuses,
                          '=COUNTIF(RegStatus,"{item}")')

# monthly intake
MO_TOP = STA_A_E + 3
wsA.cell(row=MO_TOP, column=2, value="MONTHLY CASE INTAKE (12 MONTHS)").font = font(10, True, NAVY)
wsA.cell(row=MO_TOP + 1, column=2, value="MONTH").font = font(8, True, MUTED)
wsA.cell(row=MO_TOP + 1, column=3, value="FILED").font = font(8, True, MUTED)
wsA.cell(row=MO_TOP + 1, column=3).alignment = valign("center")
months = []
y, m = TODAY.year, TODAY.month
for k in range(11, -1, -1):
    mm = m - k
    yy = y
    while mm <= 0:
        mm += 12; yy -= 1
    months.append(dt.date(yy, mm, 1))
for i, start in enumerate(months):
    r = MO_TOP + 2 + i
    nxt = dt.date(start.year + (start.month // 12), (start.month % 12) + 1, 1)
    a = wsA.cell(row=r, column=2, value=start); a.number_format = "mmm yyyy"; a.font = font(9)
    b = wsA.cell(row=r, column=3, value=f'=COUNTIFS(RegFiled,">="&$B{r},RegFiled,"<"&DATE({nxt.year},{nxt.month},1))')
    b.font = font(9, True, NAVY); b.alignment = valign("center")
    for cell in (a, b):
        cell.border = Border(bottom=side("thin", "EDF1F6"))
MO_A_S, MO_A_E = MO_TOP + 2, MO_TOP + 1 + len(months)

CRT_A_S, CRT_A_E = ablock(MO_A_E + 3, "CASES BY COURT / FORUM", "COURT / FORUM", courts,
                          '=COUNTIF(RegCourt,"{item}")')
PER_A_S, PER_A_E = ablock(CRT_A_E + 3, "CASELOAD PER PERSONNEL", "PERSONNEL", personnel,
                          '=COUNTIF(RegPersonnel,"{item}")')
BRG_A_S, BRG_A_E = ablock(PER_A_E + 3, "CASES BY BARANGAY", "BARANGAY", barangays,
                          '=COUNTIF(RegBrgy,"{item}")')

# hearing + document summary block
SUM_TOP = BRG_A_E + 3
wsA.cell(row=SUM_TOP, column=2, value="HEARING AND DOCUMENT TOTALS").font = font(10, True, NAVY)
summary_rows = [
    ("Hearings logged", "=COUNTA('Hearing Calendar'!$A$3:$A$302)"),
    ("Hearings held as scheduled", "=COUNTIF('Hearing Calendar'!$H$3:$H$302,\"Held as Scheduled\")"),
    ("Hearings postponed / reset", "=COUNTIF('Hearing Calendar'!$H$3:$H$302,\"Postponed - Prosecution\")"
        "+COUNTIF('Hearing Calendar'!$H$3:$H$302,\"Postponed - Defense\")"
        "+COUNTIF('Hearing Calendar'!$H$3:$H$302,\"Postponed - Court\")"
        "+COUNTIF('Hearing Calendar'!$H$3:$H$302,\"Reset\")"),
    ("Hearings cancelled", "=COUNTIF('Hearing Calendar'!$H$3:$H$302,\"Cancelled\")"),
    ("Document items marked missing (N)", f"=COUNTIF('Case Register'!{COL['doc1']}{REG_FIRST}:{COL['doc8']}{REG_LAST},\"N\")"),
    ("Document items on file (Y)", f"=COUNTIF('Case Register'!{COL['doc1']}{REG_FIRST}:{COL['doc8']}{REG_LAST},\"Y\")"),
]
for i, (lbl, f) in enumerate(summary_rows):
    r = SUM_TOP + 1 + i
    a = wsA.cell(row=r, column=2, value=lbl); a.font = font(9)
    b = wsA.cell(row=r, column=3, value=f); b.font = font(9, True, NAVY); b.alignment = valign("center")
    for cell in (a, b):
        cell.border = Border(bottom=side("thin", "EDF1F6"))

# ================================================================ DASHBOARD
wsD = wb.create_sheet("Case Dashboard")
wsD.sheet_properties.tabColor = GOLD
wsD.sheet_view.showGridLines = False
wsD.sheet_view.zoomScale = 90

from openpyxl.drawing.text import RegularTextRun, RichTextProperties
from openpyxl.chart.title import Title
from openpyxl.chart.text import Text
from openpyxl.chart.axis import ChartLines

set_widths(wsD, {"A": 2.5, "B": 13.5, "C": 13.5, "D": 13.5, "E": 13.5, "F": 13.5, "G": 13.5,
                 "H": 13.5, "I": 13.5, "J": 13.5, "K": 13.5, "L": 13.5, "M": 13.5, "N": 2.5})

BG_ROWS = 53
for r in range(1, BG_ROWS + 1):
    for cidx in range(1, 15):
        wsD.cell(row=r, column=cidx).fill = fill(BG)

def dfont(sz=900, b=False, color=INK):
    return CharacterProperties(sz=sz, b=b, solidFill=color, latin=DFont(typeface="Segoe UI"))

def txt(paras):
    return RichText(bodyPr=RichTextProperties(), p=paras)

def one_para(text, sz=900, b=False, color=INK):
    cp = dfont(sz, b, color)
    return Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(t=text, rPr=cp)], endParaRPr=cp)

# header
wsD.merge_cells("B1:M1")
wsD["B1"] = '=UPPER(LGUName)&"   \u2022   "&UPPER(Province)'
wsD["B1"].font = font(10, True, "BFD3E6")
wsD["B1"].fill = fill(NAVY)
wsD["B1"].alignment = valign("left", "center", False, 1)
for cidx in range(15, 2, -1):
    pass
for rr in ("B1",):
    pass
wsD.row_dimensions[1].height = 24

wsD.merge_cells("B2:M2")
wsD["B2"] = "CRIMINAL CASE MONITORING DASHBOARD"
wsD["B2"].font = font(21, True, WHITE)
wsD["B2"].fill = fill(NAVY)
wsD["B2"].alignment = valign("left", "center", False, 1)
wsD.row_dimensions[2].height = 40

wsD.merge_cells("B3:M3")
wsD["B3"] = ('=PeriodLabel&"     \u2022     As of "&TEXT(TODAY(),"dd mmmm yyyy")&'
             '"     \u2022     Prepared by: "&PreparedBy&"     \u2022     "&OfficeName')
wsD["B3"].font = font(9, False, "D6E2EE")
wsD["B3"].fill = fill(NAVY)
wsD["B3"].alignment = valign("left", "center", False, 1)
wsD.row_dimensions[3].height = 22

wsD.merge_cells("B4:M4")
wsD["B4"].fill = fill(GOLD)
wsD.row_dimensions[4].height = 4

for c in range(2, 14):
    wsD.cell(row=1, column=c).fill = fill(NAVY)
    wsD.cell(row=2, column=c).fill = fill(NAVY)
    wsD.cell(row=3, column=c).fill = fill(NAVY)

def section(row, text, span=("B", "M")):
    wsD.merge_cells(f"{span[0]}{row}:{span[1]}{row}")
    c = wsD[f"{span[0]}{row}"]
    c.value = text
    c.font = font(9, True, NAVY_3)
    c.alignment = valign("left", "bottom")
    c.fill = fill(BG)

def kpi(top, col, label, formula, sub, color, fmt="0"):
    a, b = get_column_letter(col), get_column_letter(col + 1)
    for rr, txt_, ft, al, h in [
        (top, label.upper(), font(7.5, True, MUTED), valign("left", "center", True, 1), 22),
        (top + 1, formula, font(24, True, color, name=F), valign("left", "center", False, 1), 36),
        (top + 2, sub, font(7.5, False, "8A94A6"), valign("left", "center", True, 1), 21),
    ]:
        wsD.merge_cells(f"{a}{rr}:{b}{rr}")
        wsD.row_dimensions[rr].height = h
        c = wsD[f"{a}{rr}"]
        c.value = txt_
        c.font = ft
        c.alignment = al
        c.fill = fill(WHITE)
        if rr == top:
            c.border = Border(left=side("thin", LINE), right=side("thin", LINE),
                              top=Side(style="medium", color=color))
        elif rr == top + 2:
            c.border = Border(left=side("thin", LINE), right=side("thin", LINE),
                              bottom=side("thin", LINE))
        else:
            c.border = Border(left=side("thin", LINE), right=side("thin", LINE))
    wsD[f"{a}{top + 1}"].number_format = fmt

section(6, "CASE INVENTORY AT A GLANCE")
kpi(7, 2, "Total cases monitored", "=COUNTA(RegCaseNo)", "All records in the register", NAVY)
kpi(7, 4, "Active cases", '=COUNTIF(RegStatus,"Active")', "Currently being acted upon", BLUE)
kpi(7, 6, "For action", '=COUNTIF(RegStatus,"For Action")', "Awaiting filing / assignment", AMBER)
kpi(7, 8, "Resolved / closed", '=COUNTIF(RegStatus,"Closed")', "Closed, dismissed or decided", GREEN)
kpi(7, 10, "High priority", '=COUNTIF(RegPriority,"High")', "Flagged for urgent attention", RED)
kpi(7, 12, "Average case progress", "=IFERROR(AVERAGE(RegProgress),0)", "Based on current stage", TEAL, "0%")

section(11, "AGEING AND RISK EXPOSURE  (thresholds are set on the Lists sheet)")
kpi(12, 2, "Critical", '=COUNTIF(RegAge,"Critical")',
    '=CONCATENATE("In stage over ",AgeCrit," days")', RED)
kpi(12, 4, "For follow-up", '=COUNTIF(RegAge,"For Follow-up")',
    '=CONCATENATE("In stage over ",AgeFu," days")', AMBER)
kpi(12, 6, "Under monitoring", '=COUNTIF(RegAge,"Monitor")',
    '=CONCATENATE("In stage over ",AgeMonitor," days")', BLUE)
kpi(12, 8, "Within standard", '=COUNTIF(RegAge,"Within Standard")', "Within the set threshold", GREEN)
kpi(12, 10, "Overdue hearings", '=COUNTIFS(RegNextHearing,"<"&TODAY(),RegStatus,"<>Closed")',
    "Next hearing date already passed", RED)
kpi(12, 12, "Average days in stage", "=IFERROR(ROUND(AVERAGE(RegDaysStage),0),0)",
    "Mean days in the current stage", NAVY)

# --- chart helpers -------------------------------------------------------
def new_chart(title, subtitle=None):
    pass

def style(ch, title, legend="b", height=7.0, width=9.7):
    cp = dfont(1050, True, NAVY)
    ch.title = Title(tx=Text(rich=txt([one_para(title, 1050, True, NAVY)])))
    ch.height = height
    ch.width = width
    ch.style = 2
    ch.graphical_properties = GraphicalProperties(solidFill=WHITE)
    ch.graphical_properties.ln = LineProperties(solidFill=LINE)
    for ax in (getattr(ch, "x_axis", None), getattr(ch, "y_axis", None)):
        if ax is None:
            continue
        ax.txPr = txt([one_para("", 800, False, "6B7280")])
        ax.spPr = GraphicalProperties()
        ax.spPr.ln = LineProperties(solidFill=LINE)
    if legend:
        ch.legend.position = legend
        ch.legend.overlay = False
        ch.legend.txPr = txt([one_para("", 800, False, "6B7280")])
    else:
        ch.legend = None

def cat_ref(sheet, first, last):
    return Reference(sheet, min_col=2, min_row=first, max_row=last)

def val_ref(sheet, first, last, title=None):
    if title:
        return Reference(sheet, min_col=3, min_row=first - 1, max_row=last)
    return Reference(sheet, min_col=3, min_row=first, max_row=last)

# 1. cases by stage (horizontal bars)
ch1 = BarChart(); ch1.type = "bar"; ch1.grouping = "clustered"; ch1.gapWidth = 45
ch1.add_data(val_ref(wsA, STG_A_S, STG_A_E), titles_from_data=False)
ch1.set_categories(cat_ref(wsA, STG_A_S, STG_A_E))
style(ch1, "Cases by stage", legend=None, height=8.4, width=9.7)
ch1.y_axis.majorGridlines = None
ch1.x_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill="E8EDF3")))
ch1.dataLabels = DataLabelList(); ch1.dataLabels.showVal = True
ch1.series[0].graphicalProperties = GraphicalProperties(solidFill=NAVY_3)
ch1.series[0].graphicalProperties.ln = LineProperties(noFill=True)
wsD.add_chart(ch1, "B16")

# 2. status doughnut
ch2 = DoughnutChart(); ch2.holeSize = 58
ch2.add_data(val_ref(wsA, STA_A_S, STA_A_E), titles_from_data=False)
ch2.set_categories(cat_ref(wsA, STA_A_S, STA_A_E))
style(ch2, "Case status mix", legend="r", height=8.4, width=9.7)
ch2.dataLabels = DataLabelList(); ch2.dataLabels.showVal = True
ch2.dataLabels.txPr = txt([one_para("", 800, True, WHITE)])
from openpyxl.chart.marker import DataPoint
STATUS_COLORS = [NAVY, GOLD, "9CA3AF", "6B7280", GREEN, "5B21B6"]
ch2.series[0].data_points = [
    DataPoint(idx=i, spPr=GraphicalProperties(solidFill=c, ln=LineProperties(solidFill=WHITE)))
    for i, c in enumerate(STATUS_COLORS)]
wsD.add_chart(ch2, "F16")

# 3. monthly intake line
ch3 = LineChart()
ch3.add_data(val_ref(wsA, MO_A_S, MO_A_E), titles_from_data=False)
ch3.set_categories(cat_ref(wsA, MO_A_S, MO_A_E))
style(ch3, "Monthly case intake (12 months)", legend=None, height=8.4, width=9.7)
ch3.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill="E8EDF3")))
ch3.y_axis.scaling.min = 0
s = ch3.series[0]
s.graphicalProperties = GraphicalProperties()
s.graphicalProperties.ln = LineProperties(solidFill=GOLD, w=28000)
from openpyxl.chart.marker import Marker
s.marker = Marker(symbol="circle", size=6)
s.smooth = False
ch3.x_axis.txPr = txt([one_para("", 750, False, "6B7280")])
ch3.x_axis.tickLblSkip = 1
wsD.add_chart(ch3, "J16")

# 4. offense category
ch4 = BarChart(); ch4.type = "bar"; ch4.grouping = "clustered"; ch4.gapWidth = 45
ch4.add_data(val_ref(wsA, CAT_A_S, CAT_A_E), titles_from_data=False)
ch4.set_categories(cat_ref(wsA, CAT_A_S, CAT_A_E))
style(ch4, "Volume by offense category", legend=None, height=8.4, width=9.7)
ch4.y_axis.majorGridlines = None
ch4.x_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill="E8EDF3")))
ch4.dataLabels = DataLabelList(); ch4.dataLabels.showVal = True
ch4.series[0].graphicalProperties = GraphicalProperties(solidFill=NAVY_2)
ch4.series[0].graphicalProperties.ln = LineProperties(noFill=True)
wsD.add_chart(ch4, "B33")

# 5. caseload per personnel
ch5 = BarChart(); ch5.type = "col"; ch5.grouping = "clustered"; ch5.gapWidth = 55
ch5.add_data(val_ref(wsA, PER_A_S, PER_A_E), titles_from_data=False)
ch5.set_categories(cat_ref(wsA, PER_A_S, PER_A_E))
style(ch5, "Caseload per prosecutor / investigator", legend=None, height=8.4, width=9.7)
ch5.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill="E8EDF3")))
ch5.dataLabels = DataLabelList(); ch5.dataLabels.showVal = True
ch5.series[0].graphicalProperties = GraphicalProperties(solidFill=TEAL)
ch5.series[0].graphicalProperties.ln = LineProperties(noFill=True)
ch5.x_axis.txPr = txt([one_para("", 750, False, "6B7280")])
wsD.add_chart(ch5, "F33")

# 6. barangay
ch6 = BarChart(); ch6.type = "bar"; ch6.grouping = "clustered"; ch6.gapWidth = 45
ch6.add_data(val_ref(wsA, BRG_A_S, BRG_A_E), titles_from_data=False)
ch6.set_categories(cat_ref(wsA, BRG_A_S, BRG_A_E))
style(ch6, "Geographic distribution by barangay", legend=None, height=8.4, width=9.7)
ch6.y_axis.majorGridlines = None
ch6.x_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill="E8EDF3")))
ch6.dataLabels = DataLabelList(); ch6.dataLabels.showVal = True
ch6.series[0].graphicalProperties = GraphicalProperties(solidFill=GOLD)
ch6.series[0].graphicalProperties.ln = LineProperties(noFill=True)
wsD.add_chart(ch6, "J33")

wsD.merge_cells("B51:M51")
wsD["B51"] = ("Legend:  Critical / For follow-up / Monitor ageing is calculated from DAYS IN CURRENT STAGE against the "
              "thresholds on the Lists sheet.  Red hearing dates are already overdue; gold dates fall inside the "
              "hearing window.  All figures refresh from the Case Register, Hearing Calendar and Document Compliance sheets.")
wsD["B51"].font = font(7.5, False, "7A8699", True)
wsD["B51"].alignment = valign("left", "top", True)
wsD.row_dimensions[51].height = 24

wsD.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
wsD.page_setup.orientation = "landscape"
wsD.page_setup.paperSize = wsD.PAPERSIZE_A3
wsD.page_setup.fitToWidth = 1
wsD.page_setup.fitToHeight = 1
wsD.print_area = "A1:N52"

# ================================================================ HEARING CALENDAR
wsH = wb.create_sheet("Hearing Calendar")
wsH.sheet_properties.tabColor = TEAL
wsH.sheet_view.showGridLines = False
wsH.sheet_view.zoomScale = 90
H_LAST = 302

HCOLS = [
    ("case_no",  "CASE NO.", 15, "text"),
    ("title",    "CASE TITLE", 34, "auto"),
    ("hdate",    "HEARING DATE", 13, "date"),
    ("htime",    "TIME", 11, "text"),
    ("htype",    "HEARING TYPE", 24, "choice"),
    ("court",    "COURT / FORUM", 22, "auto"),
    ("handler",  "HANDLER", 24, "auto"),
    ("result",   "RESULT", 22, "choice"),
    ("nextset",  "NEXT SETTING", 13, "date"),
    ("days",     "DAYS FROM TODAY", 11, "autonum"),
    ("flag",     "STATUS", 14, "autoc"),
    ("notes",    "ORDERS / NOTES", 46, "textw"),
]
HC = {k: get_column_letter(i + 1) for i, (k, *_x) in enumerate(HCOLS)}
HLAST_COL = get_column_letter(len(HCOLS))

wsH.merge_cells(f"A1:{HLAST_COL}1")
wsH["A1"] = "HEARING CALENDAR   |   encode every setting, then update the result after each hearing"
wsH["A1"].font = font(11.5, True, WHITE)
wsH["A1"].fill = fill(TEAL)
wsH["A1"].alignment = valign("left", "center", False, 1)
for cidx in range(2, len(HCOLS) + 1):
    wsH.cell(row=1, column=cidx).fill = fill(TEAL)
wsH.row_dimensions[1].height = 28

wsH.row_dimensions[2].height = 34
for k, h, w, kind in HCOLS:
    c = wsH[f"{HC[k]}2"]
    c.value = h
    c.font = font(8.5, True, WHITE)
    c.fill = fill(NAVY_3 if kind != "auto" else CALC)
    c.alignment = valign("center", "center", True)
    c.border = box(WHITE)
    wsH.column_dimensions[HC[k]].width = w

for r in range(3, H_LAST + 1):
    wsH.row_dimensions[r].height = 26
    for k, h, w, kind in HCOLS:
        c = wsH[f"{HC[k]}{r}"]
        c.border = Border(bottom=side("thin", "E4EAF1"), right=side("thin", "EDF1F6"))
        c.font = font(9)
        c.alignment = valign("left", "center")
        if kind == "date":
            c.number_format = "dd mmm yyyy"; c.alignment = valign("center", "center")
        elif kind == "textw":
            c.alignment = valign("left", "center", True)
        elif kind in ("auto", "autonum", "autoc"):
            c.fill = fill("EFF2F6")
            c.alignment = valign("center", "center") if kind != "auto" else valign("left", "center", True)
            if kind == "autonum":
                c.number_format = "0"

def putH(key, builder):
    for r in range(3, H_LAST + 1):
        wsH[f"{HC[key]}{r}"] = builder(r)

putH("title", lambda r: f'=IFERROR(IF(${HC["case_no"]}{r}="","",INDEX(RegTitle,MATCH(${HC["case_no"]}{r},RegCaseNo,0))),"Case no. not found")')
putH("court", lambda r: f'=IFERROR(IF(${HC["case_no"]}{r}="","",INDEX(RegCourt,MATCH(${HC["case_no"]}{r},RegCaseNo,0))),"")')
putH("handler", lambda r: f'=IFERROR(IF(${HC["case_no"]}{r}="","",INDEX(RegPersonnel,MATCH(${HC["case_no"]}{r},RegCaseNo,0))),"")')
putH("days", lambda r: f'=IF({HC["hdate"]}{r}="","",{HC["hdate"]}{r}-TODAY())')
putH("flag", lambda r: (
    f'=IF({HC["case_no"]}{r}="","",'
    f'IF(AND({HC["result"]}{r}="",{HC["hdate"]}{r}<TODAY()),"Overdue",'
    f'IF({HC["result"]}{r}="","Upcoming",'
    f'IF({HC["result"]}{r}="Held as Scheduled","Held",'
    f'IF({HC["result"]}{r}="Cancelled","Cancelled",'
    f'IF({HC["result"]}{r}="Submitted for Resolution","Cleared","Postponed"))))))'))

# validations
dv_case_h = DataValidation(type="list", formula1="RegCaseNo", allow_blank=True, showErrorMessage=True,
                           errorTitle="Case not in register", error="Encode the case in the Case Register first.")
wsH.add_data_validation(dv_case_h); dv_case_h.add(f"A3:A{H_LAST}")
for key, name, err in [("htype", "HearTypeList", "Pick a hearing type from the Lists sheet."),
                       ("result", "HearResultList", "Pick a result from the Lists sheet.")]:
    dv = DataValidation(type="list", formula1=name, allow_blank=True, showErrorMessage=True,
                        errorTitle="Invalid entry", error=err)
    wsH.add_data_validation(dv); dv.add(f"{HC[key]}3:{HC[key]}{H_LAST}")
dv_d = DataValidation(type="date", operator="between", formula1="DATE(2000,1,1)", formula2="DATE(2100,12,31)",
                      allow_blank=True, showErrorMessage=True, errorTitle="Invalid date", error="Enter a valid date.")
wsH.add_data_validation(dv_d)
dv_d.add(f"{HC['hdate']}3:{HC['hdate']}{H_LAST}")
dv_d.add(f"{HC['nextset']}3:{HC['nextset']}{H_LAST}")

fl = f"K3:K{H_LAST}"
wsH.conditional_formatting.add(fl, FormulaRule(formula=['$K3="Overdue"'], fill=fill(RED),
                                               font=font(9, True, WHITE), stopIfTrue=True))
wsH.conditional_formatting.add(fl, FormulaRule(formula=['$K3="Upcoming"'], fill=fill(GOLD_L),
                                               font=font(9, True, GOLD), stopIfTrue=True))
wsH.conditional_formatting.add(fl, FormulaRule(formula=['$K3="Held"'], fill=fill(GREEN_L),
                                               font=font(9, True, GREEN), stopIfTrue=True))
wsH.conditional_formatting.add(fl, FormulaRule(formula=['$K3="Postponed"'], fill=fill(AMBER_L),
                                               font=font(9, True, AMBER), stopIfTrue=True))
wsH.conditional_formatting.add(fl, FormulaRule(formula=['$K3="Cancelled"'], fill=fill("F3F4F6"),
                                               font=font(9, True, MUTED), stopIfTrue=True))
wsH.conditional_formatting.add(fl, FormulaRule(formula=['$K3="Cleared"'], fill=fill(BLUE_L),
                                               font=font(9, True, BLUE), stopIfTrue=True))

H_SAMPLES = [
 ("CC-2026-0001", D(-26), "09:00 AM", "Trial", "Held as Scheduled", D(4), "Prosecution witness No. 3 testified; exhibits A to D admitted."),
 ("CC-2026-0001", D(4), "09:00 AM", "Trial", "", None, "Continue trial. Subpoena remaining witnesses."),
 ("CC-2026-0002", D(-18), "10:30 AM", "Pre-Trial", "Held as Scheduled", D(12), "Stipulation on chain of custody; pre-trial order due."),
 ("CC-2026-0003", D(-37), "02:00 PM", "Arraignment", "Reset", D(-6), "Counsel moved for reset; court granted with warning."),
 ("CC-2026-0003", D(-6), "02:00 PM", "Arraignment", "", None, "Overdue setting. Confirm new calendar with branch clerk."),
 ("CC-2026-0004", D(-10), "08:30 AM", "Arraignment", "Held as Scheduled", D(20), "Not guilty plea entered; pre-trial set."),
 ("CC-2026-0005", D(-2), "09:30 AM", "Trial", "Postponed - Defense", D(2), "Defense counsel unavailable; motion for continuance granted."),
 ("CC-2026-0007", D(-40), "10:00 AM", "Promulgation", "Held as Scheduled", None, "Judgment of conviction promulgated; records updated."),
 ("CC-2025-0012", D(-15), "01:30 PM", "Arraignment", "Held as Scheduled", D(7), "Accused arraigned with counsel; bail bond approved."),
 ("CC-2025-0019", D(-30), "09:00 AM", "Motion Hearing", "Submitted for Resolution", None, "Appeal briefs submitted; awaiting CA resolution."),
 ("CC-2026-0022", D(5), "03:00 PM", "Motion Hearing", "", None, "Status conference on possible compromise agreement."),
 ("CC-2026-0025", D(-28), "08:30 AM", "Arraignment", "Held as Scheduled", D(11), "Plea of not guilty; pre-trial conference scheduled."),
]
for i, (cn, hd, tm, ht, rs, ns, nt) in enumerate(H_SAMPLES):
    r = 3 + i
    wsH[f"{HC['case_no']}{r}"] = cn
    wsH[f"{HC['hdate']}{r}"] = hd
    wsH[f"{HC['htime']}{r}"] = tm
    wsH[f"{HC['htype']}{r}"] = ht
    if rs:
        wsH[f"{HC['result']}{r}"] = rs
    if ns:
        wsH[f"{HC['nextset']}{r}"] = ns
    wsH[f"{HC['notes']}{r}"] = nt

tblH = Table(displayName="tblHearings", ref=f"A2:{HLAST_COL}{H_LAST}")
tblH.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True,
                                     showColumnStripes=False, showFirstColumn=False, showLastColumn=False)
wsH.add_table(tblH)
wsH.freeze_panes = "C3"
wsH.print_title_rows = "1:2"
wsH.page_setup.orientation = "landscape"
wsH.page_setup.fitToWidth = 1
wsH.page_setup.fitToHeight = 0
wsH.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ DOCUMENT COMPLIANCE
for i in range(1, 9):
    reg_name(f"RegDoc{i}", f"doc{i}")

wsC = wb.create_sheet("Document Compliance")
wsC.sheet_properties.tabColor = "B08F3C"
wsC.sheet_view.showGridLines = False
wsC.sheet_view.zoomScale = 90
C_LAST = 252

CCOLS = [
    ("case_no", "CASE NO.", 15, "text"),
    ("title", "CASE TITLE", 36, "auto"),
    ("onfile", "ON FILE", 9, "autonum"),
    ("missing", "MISSING", 9, "autonum"),
    ("pct", "COMPLETE", 11, "autopct"),
    ("detail", "MISSING ITEMS", 52, "auto"),
    ("action", "REMEDIAL ACTION", 34, "textw"),
    ("owner", "FOLLOW-UP OWNER", 24, "choice"),
    ("target", "TARGET DATE", 13, "date"),
    ("status", "COMPLIANCE STATUS", 18, "autoc"),
]
CC = {k: get_column_letter(i + 1) for i, (k, *_x) in enumerate(CCOLS)}
CLAST_COL = get_column_letter(len(CCOLS))

wsC.merge_cells(f"A1:{CLAST_COL}1")
wsC["A1"] = "DOCUMENT COMPLETENESS MONITOR   |   flags missing records that derail a case"
wsC["A1"].font = font(11.5, True, NAVY)
wsC["A1"].fill = fill(GOLD)
wsC["A1"].alignment = valign("left", "center", False, 1)
for cidx in range(2, len(CCOLS) + 1):
    wsC.cell(row=1, column=cidx).fill = fill(GOLD)
wsC.row_dimensions[1].height = 28

wsC.row_dimensions[2].height = 34
for k, h, w, kind in CCOLS:
    c = wsC[f"{CC[k]}2"]
    c.value = h
    c.font = font(8.5, True, WHITE)
    c.fill = fill(NAVY_3 if kind != "auto" else CALC)
    c.alignment = valign("center", "center", True)
    c.border = box(WHITE)
    wsC.column_dimensions[CC[k]].width = w

for r in range(3, C_LAST + 1):
    wsC.row_dimensions[r].height = 26
    for k, h, w, kind in CCOLS:
        c = wsC[f"{CC[k]}{r}"]
        c.border = Border(bottom=side("thin", "E4EAF1"), right=side("thin", "EDF1F6"))
        c.font = font(9)
        c.alignment = valign("left", "center")
        if kind == "date":
            c.number_format = "dd mmm yyyy"; c.alignment = valign("center", "center")
        elif kind == "textw":
            c.alignment = valign("left", "center", True)
        elif kind in ("auto", "autonum", "autopct", "autoc"):
            c.fill = fill("EFF2F6")
            c.alignment = valign("center", "center") if kind != "auto" else valign("left", "center", True)
            if kind == "autonum":
                c.number_format = "0"
            if kind == "autopct":
                c.number_format = "0%"

def putC(key, builder):
    for r in range(3, C_LAST + 1):
        wsC[f"{CC[key]}{r}"] = builder(r)

putC("title", lambda r: f'=IFERROR(IF($A{r}="","",INDEX(RegTitle,MATCH($A{r},RegCaseNo,0))),"Case no. not found")')

def counted(kind, r):
    parts = "+".join([f'(RegDoc{i}="{kind}")' for i in range(1, 9)])
    return f'IF($A{r}="","",SUMPRODUCT((RegCaseNo=$A{r})*({parts})))'

putC("onfile", lambda r: "=" + counted("Y", r))
putC("missing", lambda r: "=" + counted("N", r))
putC("pct", lambda r: f'=IFERROR(IF($A{r}="","",{CC["onfile"]}{r}/({CC["onfile"]}{r}+{CC["missing"]}{r})),0)')

detail_tpl = (
    '=IF($A{r}="","",TRIM('
    'IF(INDEX(RegDoc1,MATCH($A{r},RegCaseNo,0))="N","Affidavit of Complaint; ","")&'
    'IF(INDEX(RegDoc2,MATCH($A{r},RegCaseNo,0))="N","Sworn Statements; ","")&'
    'IF(INDEX(RegDoc3,MATCH($A{r},RegCaseNo,0))="N","Police / Incident Report; ","")&'
    'IF(INDEX(RegDoc4,MATCH($A{r},RegCaseNo,0))="N","Evidence and Exhibits; ","")&'
    'IF(INDEX(RegDoc5,MATCH($A{r},RegCaseNo,0))="N","Resolution / Information; ","")&'
    'IF(INDEX(RegDoc6,MATCH($A{r},RegCaseNo,0))="N","Court Order / Docket; ","")&'
    'IF(INDEX(RegDoc7,MATCH($A{r},RegCaseNo,0))="N","Bail / Bond Records; ","")&'
    'IF(INDEX(RegDoc8,MATCH($A{r},RegCaseNo,0))="N","Judgment / Order; ","")))')
putC("detail", lambda r: detail_tpl.replace("{r}", str(r)))

putC("status", lambda r: (
    f'=IF($A{r}="","",IF({CC["pct"]}{r}=1,"Complete",'
    f'IF(AND({CC["target"]}{r}<>"",{CC["target"]}{r}<TODAY(),{CC["pct"]}{r}<1),"Follow-up overdue",'
    f'IF({CC["pct"]}{r}>=0.75,"Minor gaps","Action needed"))))'))

dv_case_c = DataValidation(type="list", formula1="RegCaseNo", allow_blank=True, showErrorMessage=True,
                           errorTitle="Case not in register", error="Encode the case in the Case Register first.")
wsC.add_data_validation(dv_case_c); dv_case_c.add(f"A3:A{C_LAST}")
dv_own = DataValidation(type="list", formula1="PersonnelList", allow_blank=True, showErrorMessage=True,
                        errorTitle="Invalid entry", error="Pick a person from the Lists sheet.")
wsC.add_data_validation(dv_own); dv_own.add(f"{CC['owner']}3:{CC['owner']}{C_LAST}")
dv_tgt = DataValidation(type="date", operator="between", formula1="DATE(2000,1,1)", formula2="DATE(2100,12,31)",
                        allow_blank=True, showErrorMessage=True, errorTitle="Invalid date")
wsC.add_data_validation(dv_tgt); dv_tgt.add(f"{CC['target']}3:{CC['target']}{C_LAST}")

cs = f"{CC['status']}3:{CC['status']}{C_LAST}"
wsC.conditional_formatting.add(cs, FormulaRule(formula=[f'${CC["status"]}3="Complete"'], fill=fill(GREEN_L),
                                               font=font(9, True, GREEN), stopIfTrue=True))
wsC.conditional_formatting.add(cs, FormulaRule(formula=[f'${CC["status"]}3="Follow-up overdue"'], fill=fill(RED_L),
                                               font=font(9, True, RED), stopIfTrue=True))
wsC.conditional_formatting.add(cs, FormulaRule(formula=[f'${CC["status"]}3="Action needed"'], fill=fill(AMBER_L),
                                               font=font(9, True, AMBER), stopIfTrue=True))
wsC.conditional_formatting.add(cs, FormulaRule(formula=[f'${CC["status"]}3="Minor gaps"'], fill=fill(BLUE_L),
                                               font=font(9, True, BLUE), stopIfTrue=True))
wsC.conditional_formatting.add(f"{CC['pct']}3:{CC['pct']}{C_LAST}",
    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=GREEN, showValue=True))

C_SAMPLES = [
 ("CC-2026-0001", "Request certified true copy of court orders from RTC records", "Atty. R. Villanueva", D(10)),
 ("CC-2026-0003", "Secure police blotter and sworn statements from complainant", "Investigator P. Reyes", D(-4)),
 ("CC-2026-0006", "Complete documentary stamp requirements before filing", "Investigator P. Reyes", D(7)),
 ("CC-2025-0010", "Request bail records and judgment documents from branch clerk", "Atty. S. Navarro", D(-12)),
 ("CC-2025-0013", "Obtain DSWD report and medical certificate of the child witness", "Atty. M. Fernandez", D(14)),
 ("CC-2025-0015", "Prepare affidavit of complaint and police report for filing", "Investigator L. Gomez", D(5)),
 ("CC-2025-0016", "Certify screenshots and preserve social media evidence", "Atty. S. Navarro", D(9)),
 ("CC-2026-0021", "Attach trafficking referral and victim statement", "Atty. S. Navarro", D(-2)),
 ("CC-2026-0026", "Secure EMB monitoring report; recommend case review", "Investigator L. Gomez", D(20)),
]
for i, (cn, act, own, tgt) in enumerate(C_SAMPLES):
    r = 3 + i
    wsC[f"{CC['case_no']}{r}"] = cn
    wsC[f"{CC['action']}{r}"] = act
    wsC[f"{CC['owner']}{r}"] = own
    wsC[f"{CC['target']}{r}"] = tgt

tblC = Table(displayName="tblDocs", ref=f"A2:{CLAST_COL}{C_LAST}")
tblC.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True,
                                     showColumnStripes=False, showFirstColumn=False, showLastColumn=False)
wsC.add_table(tblC)
wsC.freeze_panes = "C3"
wsC.print_title_rows = "1:2"
wsC.page_setup.orientation = "landscape"
wsC.page_setup.fitToWidth = 1
wsC.page_setup.fitToHeight = 0
wsC.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ WATCHLIST
wsW = wb.create_sheet("Watchlist")
wsW.sheet_properties.tabColor = RED
wsW.sheet_view.showGridLines = False
wsW.sheet_view.zoomScale = 100

set_widths(wsW, {"A": 5, "B": 16, "C": 46, "D": 32, "E": 26, "F": 26, "G": 14, "H": 12, "I": 14, "J": 15})
wsW.merge_cells("A1:J1")
wsW["A1"] = "FOLLOW-UP WATCHLIST   |   the ten longest-running open cases"
wsW["A1"].font = font(13, True, WHITE)
wsW["A1"].fill = fill(RED)
wsW["A1"].alignment = valign("left", "center", False, 1)
for cidx in range(2, 11):
    wsW.cell(row=1, column=cidx).fill = fill(RED)
wsW.row_dimensions[1].height = 30

wsW.merge_cells("A2:J2")
wsW["A2"] = ("Ranked automatically by days pending, counting open cases only. This is the shortlist to bring to the "
             "weekly case conference - no manual sorting needed.")
wsW["A2"].font = font(9, False, MUTED, True)
wsW["A2"].alignment = valign("left", "center", True)
wsW.row_dimensions[2].height = 20

WHDR = ["#", "CASE NO.", "CASE TITLE", "OFFENSE / VIOLATION", "COURT / FORUM",
        "ASSIGNED HANDLER", "STATUS", "DAYS PENDING", "NEXT HEARING", "STAGE AGEING"]
wsW.row_dimensions[4].height = 32
for i, h in enumerate(WHDR, start=1):
    c = wsW.cell(row=4, column=i, value=h)
    c.font = font(8.5, True, WHITE)
    c.fill = fill(NAVY)
    c.alignment = valign("center", "center", True)
    c.border = box(WHITE)

for k in range(10):
    r = 5 + k
    wsW.row_dimensions[r].height = 30
    rank_cell = wsW.cell(row=r, column=1, value=k + 1)
    rank_cell.font = font(11, True, GOLD)
    rank_cell.alignment = valign("center", "center")
    rank_cell.fill = fill(GOLD_L)
    key = f"LARGE(RegRankKey,$A{r})"
    idx = f"MATCH({key},RegRankKey,0)"
    def safe(rng):
        return f'=IFERROR(IF(INDEX({rng},{idx})="","",INDEX({rng},{idx})),"")'
    formulas = [safe(n) for n in ("RegCaseNo", "RegTitle", "RegOffense", "RegCourt",
                                  "RegPersonnel", "RegStatus", "RegDaysPending",
                                  "RegNextHearing", "RegAge")]
    for j, f in enumerate(formulas, start=2):
        c = wsW.cell(row=r, column=j, value=f)
        c.font = font(9.5, j == 2)
        c.alignment = valign("left", "center", j in (3, 4))
        c.border = Border(bottom=side("thin", "E4EAF1"))
        if j == 2:
            c.font = font(9.5, True, NAVY)
        if j == 8:
            c.number_format = "dd mmm yyyy"; c.alignment = valign("center", "center")
        if j == 7:
            c.alignment = valign("center", "center")
            c.font = font(9, True)
        if j == 9:
            c.alignment = valign("center", "center")
        if j == 10:
            c.alignment = valign("center", "center")
        if j == 1:
            pass
    for j in range(1, 11):
        wsW.cell(row=r, column=j).border = Border(bottom=side("thin", "E4EAF1"),
                                                  right=side("thin", "F1F4F8"))

age_rng = "J5:J14"
wsW.conditional_formatting.add(age_rng, FormulaRule(formula=['$J5="Critical"'], fill=fill(RED_L),
                                                    font=font(9, True, RED), stopIfTrue=True))
wsW.conditional_formatting.add(age_rng, FormulaRule(formula=['$J5="For Follow-up"'], fill=fill(AMBER_L),
                                                    font=font(9, True, AMBER), stopIfTrue=True))
wsW.conditional_formatting.add(age_rng, FormulaRule(formula=['$J5="Monitor"'], fill=fill(BLUE_L),
                                                    font=font(9, True, BLUE), stopIfTrue=True))
wsW.conditional_formatting.add(age_rng, FormulaRule(formula=['$J5="Within Standard"'], fill=fill(GREEN_L),
                                                    font=font(9, True, GREEN), stopIfTrue=True))
wsW.conditional_formatting.add("H5:H14", DataBarRule(start_type="num", start_value=0, end_type="max",
                                                     color=RED, showValue=True))

wsW.merge_cells("A16:J16")
wsW["A16"] = ("Escalation guide:  Monitor = review at case conference.   For Follow-up = written follow-up with the "
              "court or prosecutor within 7 days.   Critical = escalate to the Local Chief Executive / City Prosecutor "
              "and document the action taken in the Remarks column of the Case Register.")
wsW["A16"].font = font(8.5, False, "6B7280", True)
wsW["A16"].alignment = valign("left", "top", True)
wsW.row_dimensions[16].height = 34

wsW.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
wsW.page_setup.orientation = "landscape"
wsW.page_setup.fitToWidth = 1
wsW.page_setup.fitToHeight = 1

# ================================================================ USER GUIDE
wsG = wb.create_sheet("User Guide")
wsG.sheet_properties.tabColor = GREEN
wsG.sheet_view.showGridLines = False
set_widths(wsG, {"A": 3, "B": 30, "C": 104})
wsG.merge_cells("A1:C1")
wsG["A1"] = "USER GUIDE  |  LGU CRIMINAL CASE MONITORING SYSTEM"
wsG["A1"].font = font(15, True, WHITE)
wsG["A1"].fill = fill(NAVY)
wsG["A1"].alignment = valign("left", "center", False, 1)
for cidx in (2, 3):
    wsG.cell(row=1, column=cidx).fill = fill(NAVY)
wsG.row_dimensions[1].height = 36

GUIDE = [
 ("1. WHAT THIS WORKBOOK DOES", [
   ("Purpose", "One live record of every criminal case handled or monitored by the LGU: intake, stage, court, "
               "custody, documents, hearings and outcome - with an automatic dashboard, ageing alerts and a "
               "follow-up watchlist."),
   ("Sheets", "Case Dashboard (management view)  |  Case Register (the master list)  |  Hearing Calendar  |  "
              "Document Compliance  |  Watchlist  |  Analysis (chart data)  |  Lists (dropdowns and settings)  |  "
              "User Guide."),
   ("How the parts connect", "You type only into white cells on the Register, Hearing Calendar and Document "
                              "Compliance sheets. Grey cells hold formulas. The dashboard, watchlist and analysis "
                              "tables read from those sheets, so they update the moment you encode or edit a record."),
 ]),
 ("2. FIRST-TIME SETUP (about 10 minutes)", [
   ("Step 1 - Office profile", "Open the Lists sheet and go to the OFFICE PROFILE AND SETTINGS block (column A, "
                               "lower part). Replace the LGU name, province, office, preparer and reporting period. "
                               "These appear on the dashboard header."),
   ("Step 2 - Thresholds", "Still on the Lists sheet, set the ageing thresholds in days (default: Monitor 90, "
                           "For Follow-up 180, Critical 365) and the hearing window (default 7 days). Every colour "
                           "flag in the workbook follows these numbers."),
   ("Step 3 - Barangays", "Replace the placeholder barangay names with the actual barangays of your LGU. Keep at "
                          "least one entry; add as many as you need."),
   ("Step 4 - People", "Replace the placeholder prosecutors and investigators with real names and designations. "
                       "These feed the assignment dropdowns and the caseload chart."),
   ("Step 5 - Offenses", "The OFFENSE MASTER block already lists the common offenses with the correct category and "
                         "legal basis. Add any local offense your office monitors (put a category in column B so the "
                         "dashboard categorises it correctly)."),
   ("Step 6 - Clear the samples", "The workbook ships with 26 illustrative cases so you can see it working. Delete "
                                  "the sample rows (right-click > Delete Table Rows) once you start encoding real "
                                  "data. Never delete rows 1 and 2, and never delete the last helper columns."),
 ]),
 ("3. DAILY OPERATION", [
   ("New case", "Case Register > first empty row. Type the case number (use one consistent format, e.g. CC-2026-0001), "
                "date recorded, case title, parties, offense, stage, court and handler. Grey columns fill themselves."),
   ("Logging a hearing", "Hearing Calendar > new row. Pick the case number, enter the hearing date, type and time. "
                         "After the hearing, pick the result. The STATUS column turns red if a hearing date passes "
                         "with no result, gold if it falls inside the hearing window, green when held."),
   ("Updating a stage", "When a case moves forward, change CASE STAGE and set STAGE SINCE to the date of the change. "
                        "The ageing clock resets automatically."),
   ("Tracking documents", "The Register has a Y / N / N/A checklist for eight core documents. Mark Y when the paper "
                          "is physically on file. The Document Compliance sheet summarises what is missing, "
                          "assigns a follow-up owner and a target date."),
   ("Closing a case", "Set CASE STATUS to Closed, choose the OUTCOME, and enter the DATE CLOSED. The ageing flag "
                      "switches to Closed and the case drops out of the watchlist."),
   ("Escalation", "Work the Watchlist weekly. Monitor = discuss at case conference; For Follow-up = written "
                  "follow-up within 7 days; Critical = escalate to the Local Chief Executive and City Prosecutor "
                  "and record the action in Remarks."),
 ]),
 ("4. HOW EACH CALCULATION WORKS", [
   ("Days pending", "Date closed minus date recorded for closed cases; today minus date recorded for open cases. "
                    "It measures the whole life of the case, not the current stage."),
   ("Days in current stage", "Today minus STAGE SINCE (frozen at the closing date for closed cases). This drives the "
                             "ageing flag because delay usually happens stage by stage."),
   ("Stage ageing", "Compared against the three thresholds: Critical, For Follow-up, Monitor, Within Standard. "
                    "Closed cases are excluded."),
   ("Case progress", "The percent weight assigned to each stage on the Lists sheet (For Filing 5% to Judgment 97%). "
                     "Adjust the weights if your office measures progress differently."),
   ("Document completeness", "Documents marked Y divided by documents marked Y or N (N/A entries are ignored, so a "
                             "court case without bail records is not penalised)."),
   ("Days in custody", "Today (or the closing date) minus DATE OF DETENTION, shown only for accused who are actually "
                       "detained - useful for flagging preventive-detention limits."),
 ]),
 ("5. INTEGRITY RULES - PLEASE FOLLOW", [
   ("Never type in grey cells", "Grey columns are formulas. Overwriting them breaks the dashboard for that row. If a "
                                "formula is accidentally deleted, copy the same cell from the row above or below."),
   ("Grow the table, do not paste outside it", "Each sheet is a real Excel table sized for 600 cases, 300 hearings "
                                                "and 250 document reviews. Add rows inside the table so formulas and "
                                                "dropdowns copy down automatically."),
   ("One row per case, one row per hearing", "Do not merge cells in the data areas and do not leave blank rows in the "
                                             "middle of the table - filters and counts will skip them."),
   ("Keep case numbers consistent", "The Hearing Calendar and Document Compliance sheets look cases up by case number. "
                                    "If a case number is edited in the Register, re-point the related rows."),
   ("Monthly snapshot", "At month end use File > Save As to keep a read-only copy (e.g. CASES-2026-10.xlsx). Keep the "
                        "working file as the live version."),
 ]),
 ("6. DATA PRIVACY AND SECURITY (PHILIPPINES)", [
   ("Sensitive records", "This file will contain names of accused, victims and witnesses. Store it on a restricted "
                         "drive, use a password (File > Info > Protect Workbook) and limit access to authorised "
                         "personnel under RA 10173 (Data Privacy Act)."),
   ("Minors and sensitive cases", "For cases involving children, rape, VAWC or trafficking, use initials or codes "
                                  "instead of full names in the complainant and title columns (RA 7610, RA 9262, "
                                  "RA 9208). Keep the identifying details in a separate secured log."),
   ("Retention", "Follow your records retention schedule and the National Archives rules before destroying or "
                 "archiving case records."),
   ("Sharing", "When emailing a report, send the dashboard printout or a PDF rather than the raw workbook."),
 ]),
 ("7. TROUBLESHOOTING", [
   ("A dropdown is missing", "Check the Lists sheet: if the block header or the range was deleted, the named range "
                             "breaks. Restore the block and its items."),
   ("Dashboard shows zero", "Confirm the case rows are inside the table and that CASE NO. is filled in - every "
                            "formula is gated on the case number."),
   ("#N/A in a hearing row", "The case number typed in the Hearing Calendar does not exist in the Register. Pick it "
                             "from the dropdown."),
   ("Chart does not update", "Charts read the Analysis sheet. Press F9 (or Ctrl+Alt+F9) to force a full recalculation."),
 ]),
]

row = 3
for title, items in GUIDE:
    wsG.merge_cells(f"A{row}:C{row}")
    c = wsG[f"A{row}"]
    c.value = title
    c.font = font(10.5, True, WHITE)
    c.fill = fill(NAVY_2)
    c.alignment = valign("left", "center", False, 1)
    wsG.row_dimensions[row].height = 24
    row += 1
    for i, (label, body) in enumerate(items):
        lc = wsG.cell(row=row, column=2, value=label)
        lc.font = font(9, True, NAVY)
        lc.alignment = valign("left", "top", True, 1)
        bc = wsG.cell(row=row, column=3, value=body)
        bc.font = font(9)
        bc.alignment = valign("left", "top", True)
        fill_c = WHITE if i % 2 == 0 else BG
        for col in (2, 3):
            cc = wsG.cell(row=row, column=col)
            cc.fill = fill(fill_c)
            cc.border = Border(bottom=side("thin", "E9EEF4"))
        lines = max(1, int(len(body) / 105) + 1)
        wsG.row_dimensions[row].height = max(18, 13 * lines + 6)
        row += 1
    row += 1

wsG.merge_cells(f"A{row}:C{row}")
wsG[f"A{row}"] = ("Prepared for local government case monitoring. Every figure in this workbook is a live formula - "
                  "if a number looks wrong, check the source rows in the Case Register, Hearing Calendar or "
                  "Document Compliance sheet.")
wsG[f"A{row}"].font = font(8.5, False, MUTED, True)
wsG[f"A{row}"].alignment = valign("left", "center", True, 1)

# ================================================================ ORDER AND SAVE
order = ["Case Dashboard", "Case Register", "Hearing Calendar", "Document Compliance",
         "Watchlist", "Analysis", "Lists", "User Guide"]
wb._sheets = [wb[name] for name in order]
wb.active = 0

wb.properties.title = "LGU Criminal Case Monitoring System"
wb.properties.subject = "Criminal case monitoring, ageing and compliance tracker for a Philippine local government unit"
wb.properties.creator = "Case Monitoring Office"
wb.properties.description = "Formula-driven case register, hearing calendar, document compliance monitor and dashboard."
wb.properties.keywords = "criminal cases; LGU; Philippines; case monitoring; dashboard"
wb.properties.category = "Case Management"

OUT = "/home/user/My_AI_Projects/LGU_Criminal_Case_Monitoring_System.xlsx"
wb.save(OUT)
print("saved", OUT)
