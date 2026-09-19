# 90-Taylor

Hub for tracking the Hodlin rental of 90 Taylor Street.

## Rental income & expense tracker

A Google Sheets–ready workbook for tracking rental income, expenses, monthly
cash flow, and Schedule E (Form 1040) tax figures.

| File | What it is |
|---|---|
| `90-Taylor-Rental-Tracker.xlsx` | The tracker. Upload to Google Drive, open, then **File → Save as Google Sheets**. |
| `build_tracker.py` | Script that generates the workbook (`python3 build_tracker.py`, requires `openpyxl`). Edit and re-run to change the template. |

### What's inside the workbook

Tabs are ordered and color-coded by function: **blue = reporting**,
**amber = inputs**, **violet = calculations**.

**Instructions** (gray)

- Setup steps, monthly routine, year-end tax checklist, and a Massachusetts
  landlord checklist (deposit law, allowed fees, lead paint, insurance).

**Reporting** (blue — read-only, build themselves)

- **Cash Flow** — auto-built monthly dashboard with net/cumulative cash flow,
  YTD summary, plan-vs-actual, and two charts.
- **Tax Summary** — mirrors Schedule E Part I line by line; a single year-end
  input for Form 1098 mortgage interest.

**Inputs** (amber — where you type)

- **Settings** — rent, mortgage, HOA, rental start date, tax year, category lists.
- **Income / Expenses** — transaction logs with dropdown categories mapped to
  Schedule E lines.
- **Tenants** — lease details, lease-end countdown, and a monthly rent roll
  with Paid/Partial/Unpaid status.
- **Deposit** — Massachusetts security-deposit compliance (G.L. c. 186 § 15B):
  deadlines, interest, move-out deduction ledger.
- **Cap Improvements** — improvements log with per-item current-year depreciation.
- **Mileage** — trip log at the IRS standard mileage rate.

**Calculations** (violet — one-time setup, then automatic)

- **Depreciation** — 27.5-year residential straight-line with the IRS mid-month
  convention (Pub 946 Table A-6), building + improvements.

All formulas were verified end-to-end by seeding sample transactions and
recalculating with LibreOffice.

*Bookkeeping support only — not tax or legal advice.*
