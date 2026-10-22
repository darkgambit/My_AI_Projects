# J&Z Trucking Command Center — build source

The deliverable is **`../JZ_Trucking_Command_Center.xlsx`** (repo root). This folder only holds the
generator and the QA tooling, so the template can be rebuilt or extended later.

```
build_workbook.py        generates the whole workbook (formulas, dropdowns, charts, formatting)
tools/preview.py         rough PNG render of any sheet, used to eyeball layout
JZ_TEST=1 python3 build_workbook.py    builds a small 60-row copy (../JZ_TEST.xlsx) for fast testing
```

## What the workbook contains

| Sheet | Role |
|---|---|
| **Dashboard** | Read-only. Month + truck filters drive 12 KPI cards, a "vs last month" strip, 4 charts, the P&L, per-truck results, the loan book, destination ranking and crew settlement. |
| **Trips** | Input #1 — one row per delivery (date, invoice no., truck, driver, helper, source, destination, tons, allowance, reimbursable, fuel, toll). Fee / tax / crew shares / profit are computed. |
| **Expenses** | Input #2 — one row per receipt. Expense type drives a cost group which drives the P&L. |
| **Payroll** | Automatic. One block per month, one row per crew member: share earned, allowance, reimbursables, cash advances, payouts released, balance to pay. |
| **Setup** | Master data: money rules (2% tax, 10% driver, 5% helper, 18% surcharge, allowance), the J&Z rate table (₱/ton by lane), trucks, crew, expense types, loans, dropdown lists. |
| **Read Me** | In-workbook guide (setup, daily routine, formulas, Google Sheets import notes). |
| **Calc** | Hidden engine. Protected. |

## Money model

```
rate          = Setup rate table (source → destination, ₱/ton)
gross fee     = tons × rate × (1 + fuel surcharge)      surcharge default 18%, editable per trip
2% tax        = gross × 2%
net fee       = gross − tax
driver share  = net × 10%      helper share = net × 5%   (per-person % overridable in Setup §5)
trip profit   = net − driver − helper − allowance − reimbursables − fuel − toll − other trip cost
operating     = Σ trip profit − Vehicle expenses − Overhead expenses
cash profit   = operating − loan amortisation paid
```
Cash advances, payroll/reimbursement payouts, owner's draw and asset buys are classified as
**Transfer / Other** and never touch profit (they are already counted when the trip or receipt was logged).

### Fuel rule (anti double-count)

* Per-trip diesel → Trips column **FUEL COST** → P&L line **Fuel**.
* Bulk diesel delivered to your own tank → Expenses → **Fuel – Bulk Purchase** → P&L line
  **Vehicle costs booked in Expenses**.
* There is deliberately **no "Fuel – Trip" expense type**. If one is typed anyway, the Dashboard raises
  a red warning (`Calc!Y5:Y10` guard cells also warn when bulk fuel and trip fuel are both present).

### Free typing (no dropdown needed)

`Expenses!B` accepts anything. `Expenses!I` resolves the cost group in two steps:

1. **Exact match** against Setup §6 (so your own wording wins).
2. **Keyword map** if there is no exact match — e.g. `diesel/krudo/fuel/toll/tire/repair/insurance/lto → Vehicle`,
   `loan/amortization → Financing`, `advance/payout/release/reimbursement → Transfer`,
   `rent/upa/sweldo/kuryente/office/bank → Overhead`, `draw/owner/asset → Other`.
   Unrecognised wording falls back to **Overhead** (never dropped) and is flagged amber (`Expenses!K` = 1),
   counted in `Calc!Y9` and surfaced as an ℹ note above the KPI cards.

Hidden helper columns `Expenses!J:N` hold month key, unclassified flag, settlement kind (ADV/PAY) and the
fuel-duplicate flags. Payroll and the crew settlement read the settlement flags, so free-typed
"cash advance to <name>" / "payout to <name>" still settle correctly.

## QA performed

* Every formula in the shipped 1,000-row workbook was executed with HyperFormula:
  **17,899 cells, 0 errors**, and reconciled by hand against three sample months
  (gross fee, tax, driver/helper shares, allowances, trip profit, crew balances, loan balances).
* Edge cases exercised: unknown destination (→ `NO RATE`, no `#VALUE!`), blank source (destination-only
  lookup), reversed lane, helper blank, driver blank, allowance blank (→ default), no trips in a month,
  truck filter vs `All`, cash advance vs payout settlement.
* Layout checked with `tools/preview.py` (dashboard, trips, expenses, payroll, setup, read-me).
