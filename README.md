# Payment Cycle Tracker

A lightweight Python + Excel tool that turns a raw invoice ledger into the
metrics a finance/ops team actually watches every week: **DSO, aging, and
collection efficiency.**

## Problem

Most early-stage finance teams track receivables in a spreadsheet that grows
messier every month — no consistent aging buckets, no trend line, and no fast
way to answer "how much of our cash is stuck past 90 days?" That question
usually takes someone 30+ minutes to reconstruct by hand before it reaches
the CEO.

## What I built

A small pipeline that:
1. Takes an invoice ledger (invoice ID, customer, issue date, due date, amount, paid date)
2. Computes, per invoice: days to pay, days past due, and an aging bucket (Current / 1-30 / 31-60 / 61-90 / 90+)
3. Rolls those up into DSO (Days Sales Outstanding), a 9-month DSO trend, and an aging breakdown
4. Outputs both a **matplotlib report** (`payment_cycle_report.png`) for a quick visual, and a **formula-driven Excel dashboard** (`payment_cycle_model.xlsx`) that recalculates live if the underlying data changes

`invoices_sample.csv` contains 420 synthetic but realistic invoices (12 months,
60 customers, mixed payment terms) so the tool runs end-to-end out of the box.
Swap in a real export from your AR system and it works the same way.

## How it works

- **`payment_cycle_tracker.py`** — generates/loads invoice data, computes
  aging + DSO metrics with pandas, and renders the two-panel chart
  (aging bar chart + DSO trend line vs. a 45-day target).
- **`build_excel_model.py`** — builds `payment_cycle_model.xlsx` with three
  tabs:
  - **Dashboard** — headline metrics (open AR, avg days to pay, % paid on
    time, % of AR over 90 days), all live formulas
  - **Invoice Ledger** — the full invoice-level data with formula-derived
    status, days past due, and aging bucket per row
  - **Aging Analysis** — bucketed open-AR totals with a chart, feeding the
    Dashboard
- All Excel calculations use native formulas (`SUMIFS`, `COUNTIFS`,
  `AVERAGEIFS`) rather than pre-computed Python values, so the workbook stays
  live if you paste in new invoices.

## Key metrics produced

| Metric | What it tells you |
|---|---|
| DSO (Days Sales Outstanding) | Average time cash is tied up in receivables |
| % Paid On Time | Collection discipline across customers |
| Aging buckets (Current → 90+ days) | Where risk is concentrated |
| % of Open AR 90+ days | Bad-debt exposure worth flagging early |

## Run it yourself

```bash
pip install pandas numpy matplotlib openpyxl
python payment_cycle_tracker.py     # generates data + chart
python build_excel_model.py         # builds the Excel dashboard
```

## Stack

Python (pandas, numpy, matplotlib), openpyxl for the Excel model.
