# LGU Criminal Case Monitoring System

**File:** `LGU_Criminal_Case_Monitoring_System.xlsx`
**Built for:** a Philippine local government unit (LGU) that has to monitor every criminal case
filed, referred, heard and resolved within its jurisdiction.
**Format:** native Excel workbook (.xlsx) — no macros, no add-ins, no external links.
Opens in Excel 2013 and later, Microsoft 365, LibreOffice Calc, and Google Sheets.

---

## What is inside

| Sheet | What it does |
|---|---|
| **Case Dashboard** | Management view: 12 live KPI cards, six charts (stage, status mix, monthly intake, offense category, caseload per person, barangay) and the ageing legend. Prints on one A3 landscape page. |
| **Case Register** | The master list. 600 rows, 42 columns: case identity, parties, offense with auto-category, stage, priority, status, outcome, court and docket, assignment, six key dates, custody/bail, an 8-item document checklist, and six auto-calculated measures plus remarks. |
| **Hearing Calendar** | 300 rows. Log every setting; case title, court and handler are pulled from the register. Each row flags itself as Overdue, Upcoming, Held, Postponed, Cancelled or Cleared. |
| **Document Compliance** | Per-case roll-up of the checklist: on file, missing, % complete, the exact missing items, remedial action, owner and target date. |
| **Watchlist** | Auto-ranked top ten longest-running **open** cases, ready for the weekly case conference. |
| **Analysis** | The formula tables that feed the charts (stage, category, status, monthly intake, court, personnel, barangay, hearing and document totals). |
| **Lists** | 99 pre-loaded Philippine offenses with category and legal basis, plus stages, statuses, outcomes, courts, custody statuses, hearing types/results, personnel, barangays, and the office profile + ageing thresholds. Every dropdown in the workbook reads from here. |
| **User Guide** | Setup steps, daily procedure, how each calculation works, data-integrity rules, RA 10173 / RA 7610 privacy notes, troubleshooting. |

26 illustrative cases, 12 hearings and 9 document follow-ups are pre-loaded so every formula,
chart and colour rule is visibly working. Delete those rows when you start encoding real data.

---

## Calculated measures (all live formulas, no manual math)

- **Days pending** — closed date minus date recorded, or today minus date recorded.
- **Days in current stage** — today minus *Stage since*, frozen at the closing date.
- **Stage ageing** — Critical / For Follow-up / Monitor / Within Standard, driven by three
  editable thresholds on the Lists sheet (defaults 365 / 180 / 90 days).
- **Case progress** — weighted by stage (For Filing 5% → Judgment 97%).
- **Document completeness** — Y ÷ (Y + N), so N/A items never distort the score.
- **Days in custody** — today (or closure) minus date of detention, for preventive-detention review.
- **Hearing status** — overdue / this window / held / postponed, from the hearing date and result.

Conditional formatting then paints the register: critical ageing in red, near-threshold in amber,
closed cases in grey, high priority in red, missing documents in red, overdue hearing dates in solid
red and hearings inside the notice window in gold.

---

## First-run setup (about 10 minutes)

1. **Lists sheet → Office profile and settings** (lower part of column A): LGU name, province,
   office, preparer, reporting period. These feed the dashboard header.
2. **Ageing thresholds** (same block): Monitor / For Follow-up / Critical days, and the hearing
   window. Every colour flag follows these numbers.
3. **Barangays**: replace the placeholders with the actual barangays.
4. **Personnel**: replace the placeholder prosecutors and investigators.
5. **Offenses**: keep the pre-loaded list; add local offenses and give each one a category so the
   dashboard classifies it correctly.
6. Delete the 26 sample case rows (right-click → Delete Table Rows). Never delete rows 1–2.

## Daily use

- New case → **Case Register**, first empty row; the grey columns fill themselves.
- Hearing → **Hearing Calendar**, one row per setting; update the result after the hearing.
- Stage moved → change **Case Stage** and set **Stage since** to that date (resets the ageing clock).
- Documents → mark **Y / N / N/A** in the register; work the **Document Compliance** follow-ups.
- Case closed → set status **Closed**, choose the **Outcome**, enter **Date closed**.
- Weekly → work the **Watchlist**: Monitor = discuss; For Follow-up = written follow-up in 7 days;
  Critical = escalate to the LCE / City Prosecutor and record the action in Remarks.

---

## Design and integrity notes

- Real Excel tables (`tblCases`, `tblHearings`, `tblDocs`), so adding a row copies every formula and
  dropdown automatically; AutoFilter and totals are available on every column.
- 51 defined names keep formulas readable and let you extend ranges in one place.
- Data validation on every list field and on all date fields; invalid entries are rejected with a message.
- Grey cells are formulas. The three helper columns on the far right of the register are hidden and
  excluded from printing — leave them alone.
- Print setup: A3 landscape, repeating headers, fit-to-width, footers with page numbers.
- Privacy: the file will hold names of accused, victims and witnesses. Store it on a restricted drive,
  password-protect it (File → Info → Protect Workbook) and follow RA 10173; use initials or codes for
  minors, rape, VAWC and trafficking cases (RA 7610, RA 9262, RA 9208).

## Rebuilding / customising

`tools/build_case_monitor.py` regenerates the whole workbook from scratch (openpyxl). Edit the colour
palette, column set, sample data or reference lists there, then run:

```bash
pip install openpyxl
python3 tools/build_case_monitor.py
```

For day-to-day customisation you never need the script — change the Lists sheet, the office profile,
or add rows directly in Excel.
