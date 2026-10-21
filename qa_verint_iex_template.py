"""
qa_verint_iex_template.py
=========================
Quality gate for Verint_IEX_Raw_Template.xlsx.

Two layers:
  1. Structural checks  - sheets, schema, defined names, sample rows.
  2. Arithmetic checks  - independent Python totals vs the workbook's own
     formulas, recalculated by the optional `formulas` engine.

The arithmetic layer is skipped automatically when `formulas` is not
installed, so the script still works on a bare Python install:

    pip install openpyxl formulas
    python qa_verint_iex_template.py

Exit code 0 = all checks passed.
"""

from __future__ import annotations

import calendar
import sys
from datetime import date, datetime

import openpyxl

TEMPLATE = "Verint_IEX_Raw_Template.xlsx"

SHEETS = ["README", "Setup", "Lists", "Verint_Raw", "IEX_Raw",
          "Calc_Data", "Dashboard", "Reconciliation", "DQ_Checks"]

RAW_FIELDS = ["DATE", "TIME", "END", "INT", "SITE", "QUEUE", "CHAN", "OFFERED",
              "HANDLED", "ABAND", "SHORT", "ASA", "AHT", "TALK", "HOLD", "ACW",
              "STAFFED", "REQ", "SL", "WAIT", "SRC", "NOTES"]

FAILURES: list[str] = []
CHECKS = 0


def check(condition: bool, message: str) -> None:
    global CHECKS
    CHECKS += 1
    if not condition:
        FAILURES.append(message)
        print(f"  FAIL  {message}")


def load():
    return openpyxl.load_workbook(TEMPLATE, data_only=False)


# --------------------------------------------------------------------------
# Layer 1 - structure
# --------------------------------------------------------------------------
def structural_checks(wb) -> None:
    print("\n[1] Structure")
    check(wb.sheetnames == SHEETS, f"sheet order/name mismatch: {wb.sheetnames}")

    names = set(wb.defined_names)
    for prefix in ("VER", "IEX"):
        for field in RAW_FIELDS:
            check(f"{prefix}_{field}" in names, f"missing defined name {prefix}_{field}")

    for sheet in ("Verint_Raw", "IEX_Raw"):
        ws = wb[sheet]
        headers = [ws.cell(row=5, column=c).value for c in range(1, 23)]
        check(headers[0] == "Date" and headers[7] == "Offered",
              f"{sheet}: header row is not the canonical schema")
        check(ws.freeze_panes == "A6", f"{sheet}: header not frozen at A6")
        check(ws.auto_filter.ref is not None, f"{sheet}: autofilter missing")

    setup = wb["Setup"]
    check(setup["B5"].value == date(2026, 10, 1).isoformat() or hasattr(setup["B5"].value, "year"),
          "Setup!B5 period start is not a date")
    check(str(setup["B6"].value).startswith("=EOMONTH"), "Setup!B6 is not EOMONTH based")
    check(setup["B7"].value == 30, "Setup!B7 interval minutes should default to 30")
    check(setup["B13"].value == "COMBINED", "Setup!B13 source should default to COMBINED")

    dash = wb["Dashboard"]
    check(len(dash._charts) == 4, f"Dashboard has {len(dash._charts)} charts, expected 4")


# --------------------------------------------------------------------------
# Layer 2 - arithmetic
# --------------------------------------------------------------------------
def read_raw(wb, sheet: str):
    ws = wb[sheet]
    rows = []
    for r in range(6, 10_006):
        d = ws.cell(row=r, column=1).value
        if d is None:
            continue
        rows.append({
            "date": d,
            "time": ws.cell(row=r, column=2).value,
            "site": ws.cell(row=r, column=5).value,
            "queue": ws.cell(row=r, column=6).value,
            "chan": ws.cell(row=r, column=7).value,
            "offered": ws.cell(row=r, column=8).value or 0,
            "handled": ws.cell(row=r, column=9).value or 0,
            "abandoned": ws.cell(row=r, column=10).value or 0,
            "asa": ws.cell(row=r, column=12).value or 0,
            "sl": ws.cell(row=r, column=19).value or 0,
            "talk": ws.cell(row=r, column=14).value or 0,
            "hold": ws.cell(row=r, column=15).value or 0,
            "acw": ws.cell(row=r, column=16).value or 0,
            "staffed": ws.cell(row=r, column=17).value or 0,
        })
    return rows


def as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raise TypeError(f"not a date: {value!r}")


def expected(rows, start: date, end: date, interval_minutes: int = 30) -> dict:
    rows = [x for x in rows if start <= as_date(x["date"]) <= end]
    offered = sum(x["offered"] for x in rows)
    handled = sum(x["handled"] for x in rows)
    work = sum(x["talk"] + x["hold"] + x["acw"] for x in rows)
    staffed_hours = sum(x["staffed"] for x in rows) * interval_minutes / 60
    return {
        "offered": offered,
        "handled": handled,
        "abandon_rate": sum(x["abandoned"] for x in rows) / offered if offered else 0,
        "sl": sum(x["sl"] * x["offered"] for x in rows) / offered if offered else 0,
        "aht": work / handled if handled else 0,
        "staffed_hours": staffed_hours,
        "occupancy": (work / 3600) / staffed_hours if staffed_hours else 0,
        "intervals": len(rows),
    }


def recalc_cells(path: str, cells: dict[str, list[str]]) -> dict[str, float]:
    """Recalculate a trimmed copy of the workbook with the `formulas` engine."""
    import warnings

    import formulas

    warnings.filterwarnings("ignore")
    model = formulas.ExcelModel().loads(path).finish()
    solution = model.calculate()
    wanted = {(sheet.upper(), cell.upper()) for sheet, refs in cells.items() for cell in refs}
    found: dict[str, float] = {}
    for key, value in solution.items():
        tail = key.upper().split("]")[-1]
        if "!" not in tail:
            continue
        sheet, cell = tail.replace("'", "").split("!")
        if (sheet, cell) in wanted:
            try:
                found[f"{sheet}!{cell}".upper()] = float(value.value[0, 0])
            except Exception:  # pragma: no cover - engine shape quirk
                pass
    return found


def arithmetic_checks(wb) -> None:
    print("\n[2] Arithmetic (independent totals)")
    verint = read_raw(wb, "Verint_Raw")
    iex = read_raw(wb, "IEX_Raw")
    check(len(verint) == 96, f"Verint_Raw should carry 96 sample rows, found {len(verint)}")
    check(len(iex) == 96, f"IEX_Raw should carry 96 sample rows, found {len(iex)}")

    start = as_date(wb["Setup"]["B5"].value)
    end = date(start.year, start.month,
               calendar.monthrange(start.year, start.month)[1])

    exp_v = expected(verint, start, end)
    exp_i = expected(iex, start, end)
    print(f"  python: Verint offered={exp_v['offered']} handled={exp_v['handled']} "
          f"SL={exp_v['sl']:.4f} AHT={exp_v['aht']:.2f}")
    print(f"  python: IEX    offered={exp_i['offered']} handled={exp_i['handled']} "
          f"SL={exp_i['sl']:.4f} AHT={exp_i['aht']:.2f}")

    try:
        import formulas  # noqa: F401
    except ImportError:
        print("  (skipped) install `formulas` to verify workbook arithmetic: pip install formulas")
        return

    # trim defined ranges to the sample block so the engine stays fast
    from openpyxl.workbook.defined_name import DefinedName

    for name, dn in list(wb.defined_names.items()):
        if name.startswith(("VER_", "IEX_")):
            letter = dn.attr_text.split("$")[1]
            sheet = "Verint_Raw" if name.startswith("VER_") else "IEX_Raw"
            wb.defined_names[name] = DefinedName(
                name, attr_text=f"{sheet}!${letter}$6:${letter}$105")
    trimmed = "/tmp/_verint_iex_qa.xlsx"
    wb.save(trimmed)

    cells = {"Calc_Data": ["C14", "C15", "C17", "C18", "C20", "C21", "C25", "C26",
                           "D14", "D15", "D18", "D20"]}
    got = recalc_cells(trimmed, cells)

    pairs = [
        ("Calc_Data!C14", exp_v["offered"], 0.5),
        ("Calc_Data!C15", exp_v["handled"], 0.5),
        ("Calc_Data!C17", exp_v["abandon_rate"], 1e-6),
        ("Calc_Data!C18", exp_v["sl"], 1e-6),
        ("Calc_Data!C20", exp_v["aht"], 0.05),
        ("Calc_Data!C21", exp_v["staffed_hours"], 0.05),
        ("Calc_Data!C25", exp_v["occupancy"], 1e-6),
        ("Calc_Data!C26", exp_v["intervals"], 0.5),
        ("Calc_Data!D14", exp_i["offered"], 0.5),
        ("Calc_Data!D15", exp_i["handled"], 0.5),
        ("Calc_Data!D18", exp_i["sl"], 1e-6),
        ("Calc_Data!D20", exp_i["aht"], 0.05),
    ]
    for key, want, tol in pairs:
        have = got.get(key.upper())
        status = "n/a" if have is None else f"{have:.6g}"
        check(have is not None and abs(have - want) <= tol,
              f"{key}: workbook={status} python={want:.6g}")


def main() -> int:
    print(f"QA: {TEMPLATE}")
    wb = load()
    structural_checks(wb)
    arithmetic_checks(wb)

    print(f"\n{CHECKS - len(FAILURES)}/{CHECKS} checks passed")
    if FAILURES:
        print("FAILED:")
        for f in FAILURES:
            print(" -", f)
        return 1
    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
