#!/usr/bin/env python3
"""Build the 90 Taylor St rental income/expense tracker workbook.

Produces an .xlsx designed to be imported into Google Sheets
(File > Import in Drive). All formulas are Sheets-compatible.

Usage: python3 build_tracker.py
Output: 90-Taylor-Rental-Tracker.xlsx
"""

import datetime as dt

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- constants

OUT = "90-Taylor-Rental-Tracker.xlsx"

TAX_YEAR = 2026
RENT_START = dt.date(2026, 9, 1)
MONTHLY_RENT = 3200.00
MORTGAGE_PMT = 2236.32
HOA_FEE = 290.00
MILEAGE_RATE = 0.70  # update each January to the new IRS rate

# Lease signed 09/09/2026 (GBREB fixed-term lease + addendum)
TENANT_1 = "Scott Reynolds"
TENANT_2 = "Sara Daly"
LEASE_START = dt.date(2026, 10, 1)
LEASE_END = dt.date(2027, 9, 30)
RENT_DUE_DAY = 1
KEYPAD_CODE = "839214"

INCOME_TYPES = [
    "Rent",
    "Late fee",
    "Utility reimbursement",
    "Parking",
    "Deposit forfeited (kept)",
    "Other income",
]

PAY_METHODS = ["Check", "ACH / bank transfer", "Zelle", "Venmo", "Cash",
               "Credit card", "Other"]

# (category, Schedule E line, tax treatment)
CATEGORIES = [
    ("Advertising", "5", "Deductible"),
    ("Auto and travel", "6", "Deductible"),
    ("Cleaning and maintenance", "7", "Deductible"),
    ("Commissions", "8", "Deductible"),
    ("Insurance", "9", "Deductible"),
    ("Legal and professional fees", "10", "Deductible"),
    ("Management fees", "11", "Deductible"),
    ("Other interest (not mortgage)", "13", "Deductible"),
    ("Repairs", "14", "Deductible"),
    ("Supplies", "15", "Deductible"),
    ("Taxes - real estate", "16", "Deductible"),
    ("Utilities", "17", "Deductible"),
    ("HOA / condo fees", "19 (Other)", "Deductible"),
    ("Bank & misc fees", "19 (Other)", "Deductible"),
    ("Other (describe in notes)", "19 (Other)", "Deductible"),
    ("Mortgage payment (P&I)", "-", "Cash flow only - interest deducted via Form 1098 at year-end"),
]

# IRS Pub 946 Table A-6 - 27.5-year residential rental, mid-month convention,
# year-1 rate by month placed in service.
FIRST_YEAR_RATES = [0.03485, 0.03182, 0.02879, 0.02576, 0.02273, 0.01970,
                    0.01667, 0.01364, 0.01061, 0.00758, 0.00455, 0.00152]

INCOME_ROWS = 400   # data rows in the Income log
EXP_ROWS = 500      # data rows in the Expenses log
MILE_ROWS = 300
CAPEX_ROWS = 100

# ------------------------------------------------------------------ styles

NAVY = "104281"       # sheet title bars / table headers (palette blue-650)
GRAY_HDR = "E1E0D9"   # secondary header fill
INPUT_YEL = "FFF3C4"  # editable input cells
BLUE = "2A78D6"       # chart series 1 (categorical slot 1)
AQUA = "1BAF7A"       # chart series 2 (categorical slot 2)
GREEN_FILL = "D7EED7"
YELLOW_FILL = "FDEBB2"
RED_FILL = "F6D3D3"

F_TITLE = Font(bold=True, size=14, color="FFFFFF")
F_HDR = Font(bold=True, color="FFFFFF")
F_SUBHDR = Font(bold=True, color="0B0B0B")
F_BOLD = Font(bold=True)
F_NOTE = Font(italic=True, color="52514E", size=10)

FILL_TITLE = PatternFill("solid", fgColor=NAVY)
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SUB = PatternFill("solid", fgColor=GRAY_HDR)
FILL_IN = PatternFill("solid", fgColor=INPUT_YEL)

THIN = Side(style="thin", color="C3C2B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '"$"#,##0.00'
MONEY_RED = '"$"#,##0.00;[Red]-"$"#,##0.00'
DATE_FMT = "mm/dd/yyyy"
MONTH_FMT = "mmm yyyy"
PCT3 = "0.000%"


def title_bar(ws, text, last_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c = ws.cell(row=1, column=1, value=text)
    c.font = F_TITLE
    c.fill = FILL_TITLE
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 24


def header_row(ws, row, headers, start_col=1, fill=FILL_HDR, font=F_HDR):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = font
        c.fill = fill
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)


def label_input(ws, row, label, value, fmt=None, note=None, label_col=1):
    ws.cell(row=row, column=label_col, value=label).font = F_BOLD
    c = ws.cell(row=row, column=label_col + 1, value=value)
    c.fill = FILL_IN
    c.border = BORDER
    if fmt:
        c.number_format = fmt
    if note:
        ws.cell(row=row, column=label_col + 2, value=note).font = F_NOTE
    return c


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


wb = Workbook()

# =================================================================
# Instructions
# =================================================================
ws = wb.active
ws.title = "Instructions"
title_bar(ws, "90 TAYLOR ST RENTAL TRACKER - HOW TO USE", 8)
set_widths(ws, {"A": 4, "B": 110})

INSTRUCTIONS = [
    ("", ""),
    ("H", "1. GET THIS INTO GOOGLE SHEETS"),
    ("", "Upload this .xlsx to Google Drive, open it, then File > Save as Google Sheets. "
         "Everything (formulas, dropdowns, named ranges, conditional formatting) converts. "
         "Verify the two charts on the Cash Flow tab came through; recreate via Insert > Chart if not."),
    ("", ""),
    ("H", "2. HOW THE TABS ARE ORGANIZED"),
    ("", "Tabs are ordered and color-coded by what they do:"),
    ("", "- BLUE tabs = REPORTING (Cash Flow, Tax Summary). Read-only - they build themselves from your entries."),
    ("", "- AMBER tabs = INPUTS (Settings, Income, Expenses, Tenants, Deposit, Cap Improvements, Mileage). "
         "This is where you type."),
    ("", "- VIOLET tab = CALCULATIONS (Depreciation). Set its one-time basis inputs, then it runs itself."),
    ("", "Within every tab: YELLOW cells = you type here. Everything else calculates automatically - "
         "don't overwrite it."),
    ("", ""),
    ("H", "3. ONE-TIME SETUP (do now, ~10 minutes)"),
    ("", "- Settings tab: confirm rent, mortgage, HOA, rental start date, and tax year."),
    ("", "- Depreciation tab: enter purchase price, closing costs added to basis, and land value. "
         "The Waltham assessor's property card (assessedvalues2.com or city site) shows the land/building split - "
         "using the assessor's ratio applied to your purchase price is the standard method."),
    ("", "- Tenants tab: fill in tenant and lease details once you sign."),
    ("", "- Deposit tab: fill in as soon as you collect a deposit. Massachusetts is strict - see that tab."),
    ("", ""),
    ("H", "4. MONTHLY ROUTINE (~5 minutes)"),
    ("", "- Income tab: one row per payment received (rent, late fees, etc.). Set 'For month' to the month a rent "
         "payment covers so the rent roll credits the right month. Do NOT log security deposits here - "
         "deposits are not income; they live on the Deposit tab."),
    ("", "- Expenses tab: one row per expense, including the mortgage payment (category 'Mortgage payment (P&I)') "
         "and the HOA fee (category 'HOA / condo fees'). Keep every receipt - photo into a Drive folder per year."),
    ("", "- Capital improvements (new floors, new HVAC, remodel) go ONLY on the Cap Improvements tab, never Expenses. "
         "Repairs (fix a leak, patch, repaint a room) are Expenses > 'Repairs'. Rule of thumb: a repair keeps the "
         "property in operating condition; an improvement betters it, restores it, or adapts it to a new use. "
         "Items under $2,500 per invoice can usually be expensed under the de minimis safe harbor - ask your CPA."),
    ("", "- Log trips to the property on the Mileage tab."),
    ("", "- Glance at the Cash Flow tab - it updates itself."),
    ("", ""),
    ("H", "5. YEAR-END / TAX TIME"),
    ("", "- Enter the mortgage interest from your lender's Form 1098 in the yellow cell on the Tax Summary tab."),
    ("", "- Confirm real-estate tax paid (check your escrow analysis or city bills) is fully logged in Expenses."),
    ("", "- The Tax Summary tab then mirrors Schedule E Part I (attached to your Form 1040) line by line - "
         "hand it to your CPA or copy it into tax software."),
    ("", "- Days rented / personal use days on the Tax Summary feed Schedule E lines - update if the unit sat vacant."),
    ("", ""),
    ("H", "6. BEFORE THE FIRST TENANT MOVES IN (Massachusetts checklist)"),
    ("", "- Tell your insurer: you need a landlord/dwelling-fire policy, not a homeowner's policy. Require renters "
         "insurance in the lease."),
    ("", "- Check your mortgage's occupancy clause (most require ~12 months owner-occupancy before converting to a rental)."),
    ("", "- Lead paint: if the building is pre-1978 and a child under 6 will live there, MA requires deletion/abatement. "
         "Give the tenant the Tenant Lead Law Notification regardless."),
    ("", "- Working smoke + CO detectors; keep the unit up to the State Sanitary Code."),
    ("", "- Check with the City of Waltham and your condo association: rental registration/inspection rules and "
         "any HOA rental cap, lease minimums, or move-in fees."),
    ("", "- MA landlords may only charge: first month, last month, security deposit (each max one month's rent), "
         "and a lock/key fee. No application, pet, or move-in fees."),
    ("", "- Screen consistently (fair-housing safe): same criteria for every applicant, written."),
    ("", ""),
    ("H", "7. GOOD-TO-KNOW TAX NOTES (talk to your CPA)"),
    ("", "- Rental results go on Schedule E; a net loss may be limited by passive-activity rules. If you 'actively "
         "participate' you can generally deduct up to $25,000 of losses, phasing out between $100k-$150k MAGI."),
    ("", "- Depreciation is not optional - the IRS assumes you took it when you sell (recapture). This sheet computes it."),
    ("", "- The QBI (Section 199A) deduction may apply to rentals meeting the 250-hour safe harbor - worth asking about."),
    ("", "- Keep records at least 3 years after filing; keep basis records (purchase docs, improvements) until "
         "3 years after you sell."),
    ("", ""),
    ("", "This workbook is bookkeeping help, not tax or legal advice."),
]
r = 2
for kind, text in INSTRUCTIONS:
    r += 1
    c = ws.cell(row=r, column=2, value=text)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "H":
        c.font = F_SUBHDR
        c.fill = FILL_SUB

# =================================================================
# Settings
# =================================================================
ws = wb.create_sheet("Settings")
title_bar(ws, "SETTINGS - edit the yellow cells", 11)
set_widths(ws, {"A": 34, "B": 22, "C": 46, "D": 2, "E": 26, "F": 2, "G": 22,
                "H": 2, "I": 30, "J": 14, "K": 52})

label_input(ws, 4, "Property address", "90 Taylor Street, Waltham, MA")
label_input(ws, 5, "Tax / reporting year", TAX_YEAR, "0",
            "Update each January.")
label_input(ws, 6, "Rental start (placed in service)", RENT_START, DATE_FMT,
            "First day the unit is ready and available to rent.")
label_input(ws, 7, "Monthly rent", MONTHLY_RENT, MONEY,
            "Per lease of 10/01/2026 - 09/30/2027.")
label_input(ws, 8, "Monthly mortgage payment (P&I)", MORTGAGE_PMT, MONEY)
label_input(ws, 9, "Monthly HOA fee", HOA_FEE, MONEY)
label_input(ws, 10, "IRS standard mileage rate", MILEAGE_RATE, '"$"0.000',
            "Update each January (2025 rate was $0.70/mile).")

ws.cell(row=3, column=5, value="Income types").font = F_SUBHDR
ws.cell(row=3, column=5).fill = FILL_SUB
for i, t in enumerate(INCOME_TYPES):
    ws.cell(row=4 + i, column=5, value=t).border = BORDER

ws.cell(row=3, column=7, value="Payment methods").font = F_SUBHDR
ws.cell(row=3, column=7).fill = FILL_SUB
for i, t in enumerate(PAY_METHODS):
    ws.cell(row=4 + i, column=7, value=t).border = BORDER

header_row(ws, 3, ["Expense category", "Schedule E line", "Tax treatment"],
           start_col=9, fill=FILL_SUB, font=F_SUBHDR)
for i, (cat, line, treat) in enumerate(CATEGORIES):
    for j, v in enumerate((cat, line, treat)):
        c = ws.cell(row=4 + i, column=9 + j, value=v)
        c.border = BORDER

n_cat = len(CATEGORIES)
defs = {
    "TaxYear": "Settings!$B$5",
    "RentStart": "Settings!$B$6",
    "MonthlyRent": "Settings!$B$7",
    "MortgagePmt": "Settings!$B$8",
    "HOAFee": "Settings!$B$9",
    "MileageRate": "Settings!$B$10",
    "IncomeTypes": f"Settings!$E$4:$E${3 + len(INCOME_TYPES)}",
    "PayMethods": f"Settings!$G$4:$G${3 + len(PAY_METHODS)}",
    "ExpenseCats": f"Settings!$I$4:$I${3 + n_cat}",
    "CatTable": f"Settings!$I$4:$K${3 + n_cat}",
    "IncomeDates": f"Income!$A$2:$A${1 + INCOME_ROWS}",
    "IncomeType": f"Income!$C$2:$C${1 + INCOME_ROWS}",
    "IncomeFor": f"Income!$D$2:$D${1 + INCOME_ROWS}",
    "IncomeAmts": f"Income!$E$2:$E${1 + INCOME_ROWS}",
    "ExpDates": f"Expenses!$A$2:$A${1 + EXP_ROWS}",
    "ExpCats": f"Expenses!$C$2:$C${1 + EXP_ROWS}",
    "ExpAmts": f"Expenses!$D$2:$D${1 + EXP_ROWS}",
    "MileDates": f"Mileage!$A$2:$A${1 + MILE_ROWS}",
    "MileDed": f"Mileage!$F$2:$F${1 + MILE_ROWS}",
    "CapexDates": f"'Cap Improvements'!$A$6:$A${5 + CAPEX_ROWS}",
    "CapexCosts": f"'Cap Improvements'!$C$6:$C${5 + CAPEX_ROWS}",
    "CapexDep": f"'Cap Improvements'!$F$6:$F${5 + CAPEX_ROWS}",
    "FirstYearRates": "Depreciation!$E$4:$F$15",
    "CurrentYearDep": "Depreciation!$B$14",
    "Form1098Interest": "'Tax Summary'!$C$16",
    "LeaseStart": "Tenants!$B$9",
}
for name, ref in defs.items():
    wb.defined_names.add(DefinedName(name, attr_text=ref))

# =================================================================
# Income log
# =================================================================
ws = wb.create_sheet("Income")
header_row(ws, 1, ["Date", "Payer", "Type", "For month", "Amount", "Method",
                   "Notes"])
set_widths(ws, {"A": 12, "B": 26, "C": 22, "D": 12, "E": 13, "F": 18, "G": 46})
ws.freeze_panes = "A2"
for row in range(2, 2 + INCOME_ROWS):
    ws.cell(row=row, column=1).number_format = DATE_FMT
    ws.cell(row=row, column=4).number_format = MONTH_FMT
    ws.cell(row=row, column=5).number_format = MONEY
INCOME_PREFILL = [
    (dt.date(2026, 9, 18), f"{TENANT_1} & {TENANT_2}", "Rent",
     dt.date(2026, 10, 1), 3200, None, "First month's rent"),
    (dt.date(2026, 9, 19), f"{TENANT_1} & {TENANT_2}", "Rent",
     dt.date(2027, 9, 1), 3200, None,
     "Last month's rent (prepaid) - accrues interest to tenant, see Deposit tab"),
]
for i, r in enumerate(INCOME_PREFILL):
    for j, v in enumerate(r):
        if v is not None:
            ws.cell(row=2 + i, column=1 + j, value=v)
dv_type = DataValidation(type="list", formula1="=IncomeTypes", allow_blank=True)
dv_pay = DataValidation(type="list", formula1="=PayMethods", allow_blank=True)
ws.add_data_validation(dv_type)
ws.add_data_validation(dv_pay)
dv_type.add(f"C2:C{1 + INCOME_ROWS}")
dv_pay.add(f"F2:F{1 + INCOME_ROWS}")

# =================================================================
# Expenses log
# =================================================================
ws = wb.create_sheet("Expenses")
header_row(ws, 1, ["Date", "Vendor / payee", "Category", "Amount", "Method",
                   "Receipt?", "Notes", "Tax treatment (auto)"])
set_widths(ws, {"A": 12, "B": 24, "C": 28, "D": 13, "E": 18, "F": 10,
                "G": 36, "H": 42})
ws.freeze_panes = "A2"
EXP_PREFILL = [
    (dt.date(2026, 9, 18), "Listing agent", "Commissions", 3200, None, None,
     "Broker fee for tenant placement - adjust date to the day actually paid"),
]
for i, r in enumerate(EXP_PREFILL):
    for j, v in enumerate(r):
        if v is not None:
            ws.cell(row=2 + i, column=1 + j, value=v)
for row in range(2, 2 + EXP_ROWS):
    ws.cell(row=row, column=1).number_format = DATE_FMT
    ws.cell(row=row, column=4).number_format = MONEY
    ws.cell(row=row, column=8,
            value=f'=IF($C{row}="","",IFERROR(VLOOKUP($C{row},CatTable,3,FALSE),"?"))').font = F_NOTE
dv_cat = DataValidation(type="list", formula1="=ExpenseCats", allow_blank=True)
dv_pay2 = DataValidation(type="list", formula1="=PayMethods", allow_blank=True)
dv_yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
for dv in (dv_cat, dv_pay2, dv_yn):
    ws.add_data_validation(dv)
dv_cat.add(f"C2:C{1 + EXP_ROWS}")
dv_pay2.add(f"E2:E{1 + EXP_ROWS}")
dv_yn.add(f"F2:F{1 + EXP_ROWS}")

# =================================================================
# Cash Flow dashboard
# =================================================================
ws = wb.create_sheet("Cash Flow")
title_bar(ws, "MONTHLY CASH FLOW", 10)
ws.cell(row=2, column=1, value="Year:").font = F_BOLD
c = ws.cell(row=2, column=2, value="=TaxYear")
c.font = F_BOLD

header_row(ws, 4, ["Month", "Rent income", "Other income", "Total income",
                   "Operating expenses", "Mortgage (P&I)",
                   "Capital improvements", "Total cash out",
                   "Net cash flow", "Cumulative net"])
set_widths(ws, {"A": 12, "B": 13, "C": 13, "D": 13, "E": 16, "F": 14,
                "G": 16, "H": 14, "I": 14, "J": 14, "K": 8})

for i in range(12):
    row = 5 + i
    a = f"$A{row}"
    ws.cell(row=row, column=1,
            value="=DATE(TaxYear,1,1)" if i == 0 else f"=EDATE(A{row - 1},1)")
    ws.cell(row=row, column=1).number_format = MONTH_FMT
    ws.cell(row=row, column=2,
            value=f'=SUMIFS(IncomeAmts,IncomeDates,">="&{a},IncomeDates,"<="&EOMONTH({a},0),IncomeType,"Rent")')
    ws.cell(row=row, column=3,
            value=f'=SUMIFS(IncomeAmts,IncomeDates,">="&{a},IncomeDates,"<="&EOMONTH({a},0))-B{row}')
    ws.cell(row=row, column=4, value=f"=B{row}+C{row}")
    ws.cell(row=row, column=5,
            value=f'=SUMIFS(ExpAmts,ExpDates,">="&{a},ExpDates,"<="&EOMONTH({a},0))-F{row}')
    ws.cell(row=row, column=6,
            value=f'=SUMIFS(ExpAmts,ExpDates,">="&{a},ExpDates,"<="&EOMONTH({a},0),ExpCats,"Mortgage payment (P&I)")')
    ws.cell(row=row, column=7,
            value=f'=SUMIFS(CapexCosts,CapexDates,">="&{a},CapexDates,"<="&EOMONTH({a},0))')
    ws.cell(row=row, column=8, value=f"=E{row}+F{row}+G{row}")
    ws.cell(row=row, column=9, value=f"=D{row}-H{row}")
    ws.cell(row=row, column=10, value=f"=SUM($I$5:I{row})")
    ws.cell(row=row, column=11, value=f'=TEXT($A{row},"mmm")')
    for col in range(2, 11):
        ws.cell(row=row, column=col).number_format = MONEY_RED
        ws.cell(row=row, column=col).border = BORDER
    ws.cell(row=row, column=1).border = BORDER

ws.cell(row=17, column=1, value="TOTAL").font = F_BOLD
for col in range(2, 10):
    letter = get_column_letter(col)
    c = ws.cell(row=17, column=col, value=f"=SUM({letter}5:{letter}16)")
    c.font = F_BOLD
    c.number_format = MONEY_RED
    c.fill = FILL_SUB
    c.border = BORDER
ws.cell(row=17, column=1).fill = FILL_SUB
ws.cell(row=17, column=1).border = BORDER
ws.column_dimensions["K"].hidden = False  # tiny chart-label helper column

# values live in column C so the labels in A can spill across empty B
ws.merge_cells("A19:C19")
ws.cell(row=19, column=1, value="YEAR-TO-DATE SUMMARY").font = F_SUBHDR
ws.cell(row=19, column=1).fill = FILL_SUB
kpis = [
    ("Total income", "=D17"),
    ("Total cash out", "=H17"),
    ("Net cash flow", "=I17"),
    ("Months with rent received", '=COUNTIF(B5:B16,">0")'),
    ("Average net in months with rent", '=IF(C23=0,0,I17/C23)'),
]
for i, (lbl, f) in enumerate(kpis):
    row = 20 + i
    ws.cell(row=row, column=1, value=lbl)
    c = ws.cell(row=row, column=3, value=f)
    c.font = F_BOLD
    c.number_format = MONEY_RED if i not in (3,) else "0"

ws.merge_cells("A26:C26")
ws.cell(row=26, column=1, value="MONTHLY PLAN").font = F_SUBHDR
ws.cell(row=26, column=1).fill = FILL_SUB
plan = [
    ("Target rent", "=MonthlyRent"),
    ("Mortgage (P&I)", "=-MortgagePmt"),
    ("HOA fee", "=-HOAFee"),
    ("Planned net before other costs", "=MonthlyRent-MortgagePmt-HOAFee"),
]
for i, (lbl, f) in enumerate(plan):
    row = 27 + i
    ws.cell(row=row, column=1, value=lbl)
    c = ws.cell(row=row, column=3, value=f)
    c.number_format = MONEY_RED
    if i == 3:
        c.font = F_BOLD
ws.cell(row=31, column=1,
        value="Plan excludes insurance, repairs, vacancy, taxes not in escrow - budget ~5-10% of rent for upkeep.").font = F_NOTE

# --- charts -------------------------------------------------------
chart1 = BarChart()
chart1.type = "col"
chart1.title = "Net cash flow by month"
data = Reference(ws, min_col=9, min_row=4, max_row=16)
chart1.add_data(data, titles_from_data=True)
cats = Reference(ws, min_col=11, min_row=5, max_row=16)
chart1.set_categories(cats)
chart1.legend = None
s = chart1.series[0]
s.graphicalProperties.solidFill = BLUE
chart1.y_axis.numFmt = '"$"#,##0'
chart1.y_axis.delete = False
chart1.x_axis.delete = False
chart1.width = 18
chart1.height = 9
ws.add_chart(chart1, "L4")

chart2 = BarChart()
chart2.type = "col"
chart2.grouping = "clustered"
chart2.title = "Income vs. total cash out"
data = Reference(ws, min_col=4, min_row=4, max_row=16)
chart2.add_data(data, titles_from_data=True)
data2 = Reference(ws, min_col=8, min_row=4, max_row=16)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats)
chart2.series[0].graphicalProperties.solidFill = BLUE
chart2.series[1].graphicalProperties.solidFill = AQUA
chart2.y_axis.numFmt = '"$"#,##0'
chart2.y_axis.delete = False
chart2.x_axis.delete = False
chart2.width = 18
chart2.height = 9
ws.add_chart(chart2, "L23")

# =================================================================
# Tax Summary (Schedule E)
# =================================================================
ws = wb.create_sheet("Tax Summary")
title_bar(ws, "SCHEDULE E (FORM 1040) - PART I SUMMARY", 5)
set_widths(ws, {"A": 8, "B": 46, "C": 16, "D": 46})

ws.cell(row=2, column=2, value="Tax year:").font = F_BOLD
ws.cell(row=2, column=3, value="=TaxYear").font = F_BOLD
ws.cell(row=3, column=2, value="Property (Schedule E line 1a):")
ws.cell(row=3, column=3, value="=Settings!B4")

ws.cell(row=4, column=2, value="Fair rental days (line 2)")
c = ws.cell(row=4, column=3,
            value='=MAX(0,DATE(TaxYear,12,31)-MAX(IF(LeaseStart="",RentStart,LeaseStart),DATE(TaxYear,1,1))+1)')
c.fill = FILL_IN
ws.cell(row=4, column=4,
        value="Overwrite with the actual days rented if the unit sat vacant.").font = F_NOTE
ws.cell(row=5, column=2, value="Personal use days (line 2)")
c = ws.cell(row=5, column=3, value=0)
c.fill = FILL_IN

YR = '">="&DATE(TaxYear,1,1),ExpDates,"<="&DATE(TaxYear,12,31)'


def sumifs_cat(cat):
    return f'=SUMIFS(ExpAmts,ExpDates,{YR},ExpCats,"{cat}")'


header_row(ws, 7, ["Line", "Description", "Amount"])
ws.cell(row=8, column=1, value="3").border = BORDER
ws.cell(row=8, column=2, value="Rents received").border = BORDER
c = ws.cell(row=8, column=3,
            value='=SUMIFS(IncomeAmts,IncomeDates,">="&DATE(TaxYear,1,1),IncomeDates,"<="&DATE(TaxYear,12,31))')
c.number_format = MONEY
c.border = BORDER

tax_lines = [
    ("5", "Advertising", sumifs_cat("Advertising")),
    ("6", "Auto and travel",
     f'=SUMIFS(ExpAmts,ExpDates,{YR},ExpCats,"Auto and travel")'
     f'+SUMIFS(MileDed,MileDates,">="&DATE(TaxYear,1,1),MileDates,"<="&DATE(TaxYear,12,31))'),
    ("7", "Cleaning and maintenance", sumifs_cat("Cleaning and maintenance")),
    ("8", "Commissions", sumifs_cat("Commissions")),
    ("9", "Insurance", sumifs_cat("Insurance")),
    ("10", "Legal and other professional fees", sumifs_cat("Legal and professional fees")),
    ("11", "Management fees", sumifs_cat("Management fees")),
    ("12", "Mortgage interest - enter from Form 1098 box 1", None),
    ("13", "Other interest", sumifs_cat("Other interest (not mortgage)")),
    ("14", "Repairs", sumifs_cat("Repairs")),
    ("15", "Supplies", sumifs_cat("Supplies")),
    ("16", "Taxes (real estate)", sumifs_cat("Taxes - real estate")),
    ("17", "Utilities", sumifs_cat("Utilities")),
    ("18", "Depreciation", "=CurrentYearDep"),
    ("19", "Other: HOA / condo fees", sumifs_cat("HOA / condo fees")),
    ("19", "Other: bank & misc fees", sumifs_cat("Bank & misc fees")),
    ("19", "Other: everything else", sumifs_cat("Other (describe in notes)")),
]
row = 8
for line, desc, formula in tax_lines:
    row += 1
    ws.cell(row=row, column=1, value=line).border = BORDER
    ws.cell(row=row, column=2, value=desc).border = BORDER
    c = ws.cell(row=row, column=3)
    if formula is None:
        c.value = 0
        c.fill = FILL_IN
    else:
        c.value = formula
    c.number_format = MONEY
    c.border = BORDER
# the Form 1098 input (line 12) lands at C16 (named Form1098Interest)

row += 1
ws.cell(row=row, column=1, value="20").border = BORDER
ws.cell(row=row, column=2, value="Total expenses").font = F_BOLD
ws.cell(row=row, column=2).border = BORDER
c = ws.cell(row=row, column=3, value=f"=SUM(C9:C{row - 1})")
c.font = F_BOLD
c.number_format = MONEY
c.fill = FILL_SUB
c.border = BORDER
total_row = row

row += 1
ws.cell(row=row, column=1, value="21").border = BORDER
ws.cell(row=row, column=2, value="Income or (loss)").font = F_BOLD
ws.cell(row=row, column=2).border = BORDER
c = ws.cell(row=row, column=3, value=f"=C8-C{total_row}")
c.font = F_BOLD
c.number_format = MONEY_RED
c.fill = FILL_SUB
c.border = BORDER

row += 2
for note in [
    "Security deposits you hold are not income, and returning them is not an expense.",
    "Only the interest portion of the mortgage payment (Form 1098) is deductible - never the full payment.",
    "A net loss may be limited by passive-activity rules ($25k active-participation allowance, phases out $100k-$150k MAGI).",
    "Review with your CPA - this is bookkeeping support, not tax advice.",
]:
    ws.cell(row=row, column=2, value="- " + note).font = F_NOTE
    row += 1

# =================================================================
# Depreciation
# =================================================================
ws = wb.create_sheet("Depreciation")
title_bar(ws, "DEPRECIATION - 27.5-YEAR RESIDENTIAL (MID-MONTH CONVENTION)", 8)
set_widths(ws, {"A": 38, "B": 18, "C": 12, "D": 16, "E": 16, "F": 14, "G": 4})

ws.cell(row=3, column=1, value="INPUTS").font = F_SUBHDR
ws.cell(row=3, column=1).fill = FILL_SUB
label_input(ws, 4, "Purchase price", 0, MONEY)
label_input(ws, 5, "Closing costs added to basis", 0, MONEY)
label_input(ws, 6, "Land value", 0, MONEY)
ws.cell(row=7, column=1, value="Building (depreciable) basis").font = F_BOLD
c = ws.cell(row=7, column=2, value="=B4+B5-B6")
c.number_format = MONEY
c.font = F_BOLD
c.border = BORDER
label_input(ws, 8, "Placed-in-service date", "=RentStart", DATE_FMT)
ws.cell(row=9, column=1,
        value="Land split tip: apply the Waltham assessor's land/building ratio to your purchase price.").font = F_NOTE

ws.cell(row=11, column=1, value="RESULTS").font = F_SUBHDR
ws.cell(row=11, column=1).fill = FILL_SUB
ws.cell(row=12, column=1, value="Building depreciation for the tax year")
c = ws.cell(row=12, column=2,
            value="=IFERROR(INDEX($D$17:$D$45,MATCH(TaxYear,$B$17:$B$45,0)),0)")
c.number_format = MONEY
ws.cell(row=13, column=1, value="Improvements depreciation")
c = ws.cell(row=13, column=2, value="=SUM(CapexDep)")
c.number_format = MONEY
ws.cell(row=14, column=1, value="Total depreciation for the tax year").font = F_BOLD
c = ws.cell(row=14, column=2, value="=B12+B13")
c.number_format = MONEY
c.font = F_BOLD
c.fill = FILL_SUB
c.border = BORDER

header_row(ws, 3, ["Month placed in service", "Year-1 rate"], start_col=5,
           fill=FILL_SUB, font=F_SUBHDR)
for i, rate in enumerate(FIRST_YEAR_RATES):
    ws.cell(row=4 + i, column=5, value=i + 1).border = BORDER
    c = ws.cell(row=4 + i, column=6, value=rate)
    c.number_format = PCT3
    c.border = BORDER

header_row(ws, 16, ["Year #", "Calendar year", "Rate", "Depreciation", "Cumulative"])
for i in range(29):
    row = 17 + i
    ws.cell(row=row, column=1, value=i + 1)
    ws.cell(row=row, column=2, value=f"=YEAR($B$8)+A{row}-1")
    ws.cell(row=row, column=3,
            value="=VLOOKUP(MONTH($B$8),FirstYearRates,2,FALSE)" if i == 0 else 0.03636)
    if i == 0:
        ws.cell(row=row, column=4, value="=$B$7*C17")
        ws.cell(row=row, column=5, value="=D17")
    else:
        ws.cell(row=row, column=4, value=f"=MAX(0,MIN($B$7*C{row},$B$7-E{row - 1}))")
        ws.cell(row=row, column=5, value=f"=E{row - 1}+D{row}")
    ws.cell(row=row, column=3).number_format = PCT3
    ws.cell(row=row, column=4).number_format = MONEY
    ws.cell(row=row, column=5).number_format = MONEY
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = BORDER
ws.cell(row=47, column=1,
        value="IRS Pub 946 Table A-6. Year-1 rate depends on the month placed in service; later years 3.636%.").font = F_NOTE

# =================================================================
# Cap Improvements
# =================================================================
ws = wb.create_sheet("Cap Improvements")
title_bar(ws, "CAPITAL IMPROVEMENTS - depreciate, don't deduct", 6)
set_widths(ws, {"A": 16, "B": 36, "C": 14, "D": 34, "E": 12, "F": 18})
ws.cell(row=2, column=1,
        value="Improvements (new roof share, HVAC, remodel, new flooring) depreciate over 27.5 years "
              "and add to your basis when you sell.").font = F_NOTE
ws.cell(row=3, column=1,
        value="Items under $2,500/invoice can often be expensed instead "
              "(de minimis safe harbor) - ask your CPA.").font = F_NOTE
ws.cell(row=4, column=2, value="Totals:").font = F_BOLD
c = ws.cell(row=4, column=3, value=f"=SUM(C6:C{5 + CAPEX_ROWS})")
c.font = F_BOLD
c.number_format = MONEY
c = ws.cell(row=4, column=6, value=f"=SUM(F6:F{5 + CAPEX_ROWS})")
c.font = F_BOLD
c.number_format = MONEY

header_row(ws, 5, ["Date placed in service", "Description", "Cost", "Notes",
                   "Recovery yr # (auto)", "Depreciation this tax year (auto)"])
ws.freeze_panes = "A6"
for i in range(CAPEX_ROWS):
    row = 6 + i
    ws.cell(row=row, column=1).number_format = DATE_FMT
    ws.cell(row=row, column=3).number_format = MONEY
    ws.cell(row=row, column=5, value=f'=IF($A{row}="","",TaxYear-YEAR($A{row})+1)')
    ws.cell(row=row, column=6, value=(
        f'=IF($A{row}="",0,IF(E{row}<1,0,IF(E{row}=1,'
        f'$C{row}*VLOOKUP(MONTH($A{row}),FirstYearRates,2,FALSE),'
        f'IF(E{row}>29,0,MAX(0,MIN($C{row}*0.03636,'
        f'$C{row}-$C{row}*VLOOKUP(MONTH($A{row}),FirstYearRates,2,FALSE)'
        f'-$C{row}*0.03636*(E{row}-2)))))))'))
    ws.cell(row=row, column=6).number_format = MONEY

# =================================================================
# Tenants & rent roll
# =================================================================
ws = wb.create_sheet("Tenants")
title_bar(ws, "TENANTS & LEASE", 5)
set_widths(ws, {"A": 30, "B": 24, "C": 24, "D": 14, "E": 12})

ws.cell(row=3, column=1, value="LEASE DETAILS").font = F_SUBHDR
ws.cell(row=3, column=1).fill = FILL_SUB
header_row(ws, 4, ["", "Tenant 1", "Tenant 2 / co-tenant"], fill=FILL_SUB, font=F_SUBHDR)
fields = ["Name", "Phone", "Email", "Move-in date", "Lease start", "Lease end",
          "Rent due day of month", "Monthly rent", "Security deposit held"]
for i, f in enumerate(fields):
    row = 5 + i
    ws.cell(row=row, column=1, value=f).font = F_BOLD
    for col in (2, 3):
        c = ws.cell(row=row, column=col)
        c.fill = FILL_IN
        c.border = BORDER
        if "date" in f.lower() or "start" in f.lower() or "end" in f.lower():
            c.number_format = DATE_FMT
        if "rent" in f.lower() or "deposit" in f.lower():
            c.number_format = MONEY
# lease of 09/09/2026: Scott Reynolds & Sara Daly, $3,200/mo, 10/01/2026-09/30/2027
ws["B5"] = TENANT_1
ws["C5"] = TENANT_2
for col in ("B", "C"):
    ws[f"{col}8"] = LEASE_START
    ws[f"{col}9"] = LEASE_START
    ws[f"{col}10"] = LEASE_END
    ws[f"{col}11"] = RENT_DUE_DAY
ws["B12"] = MONTHLY_RENT
ws["B11"].number_format = "0"
ws["C11"].number_format = "0"
ws.cell(row=14, column=1, value="Door keypad code").font = F_BOLD
c = ws.cell(row=14, column=2, value=KEYPAD_CODE)
c.number_format = "@"
c.fill = FILL_IN
c.border = BORDER
ws.cell(row=15, column=1, value="Days until lease end").font = F_BOLD
c = ws.cell(row=15, column=2, value='=IF(B10="","",B10-TODAY())')
c.border = BORDER
ws.conditional_formatting.add(
    "B15",
    CellIsRule(operator="lessThan", formula=["60"],
               fill=PatternFill("solid", fgColor=RED_FILL)))
ws.conditional_formatting.add(
    "B15",
    CellIsRule(operator="between", formula=["60", "90"],
               fill=PatternFill("solid", fgColor=YELLOW_FILL)))
ws.cell(row=16, column=1,
        value="Send a renewal or non-renewal notice 60-90 days before lease end.").font = F_NOTE

ws.cell(row=18, column=1, value="RENT ROLL").font = F_SUBHDR
ws.cell(row=18, column=1).fill = FILL_SUB
header_row(ws, 19, ["Month", "Expected", "Received", "Variance", "Status"])
for i in range(12):
    row = 20 + i
    ws.cell(row=row, column=1,
            value="=DATE(TaxYear,1,1)" if i == 0 else f"=EDATE(A{row - 1},1)")
    ws.cell(row=row, column=1).number_format = MONTH_FMT
    ws.cell(row=row, column=2,
            value=f'=IF(AND($B$9<>"",EOMONTH(A{row},0)>=$B$9,A{row}<=$B$10),$B$12,0)')
    ws.cell(row=row, column=3,
            value=f'=SUMIFS(IncomeAmts,IncomeFor,">="&A{row},IncomeFor,"<="&EOMONTH(A{row},0),IncomeType,"Rent")')
    ws.cell(row=row, column=4, value=f"=C{row}-B{row}")
    ws.cell(row=row, column=5,
            value=f'=IF(B{row}=0,"-",IF(C{row}>=B{row},"Paid",IF(C{row}>0,"Partial","Unpaid")))')
    for col in range(2, 5):
        ws.cell(row=row, column=col).number_format = MONEY_RED
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = BORDER
ws.cell(row=32, column=1, value="TOTAL").font = F_BOLD
for col in (2, 3, 4):
    letter = get_column_letter(col)
    c = ws.cell(row=32, column=col, value=f"=SUM({letter}20:{letter}31)")
    c.font = F_BOLD
    c.number_format = MONEY_RED
    c.fill = FILL_SUB
for state, color in (("Paid", GREEN_FILL), ("Partial", YELLOW_FILL),
                     ("Unpaid", RED_FILL)):
    ws.conditional_formatting.add(
        "E20:E31",
        CellIsRule(operator="equal", formula=[f'"{state}"'],
                   fill=PatternFill("solid", fgColor=color)))

# =================================================================
# Security Deposit (Massachusetts)
# =================================================================
ws = wb.create_sheet("Deposit")
title_bar(ws, "SECURITY DEPOSIT - MASSACHUSETTS G.L. c.186 s.15B", 6)
set_widths(ws, {"A": 40, "B": 20, "C": 44, "D": 14, "E": 16, "F": 4})
ws.cell(row=2, column=1,
        value="MA deposit law carries treble damages + attorney's fees for violations.").font = F_NOTE

ws.cell(row=4, column=1, value="DEPOSIT DETAILS").font = F_SUBHDR
ws.cell(row=4, column=1).fill = FILL_SUB
label_input(ws, 5, "Deposit amount", 0, MONEY)
c = ws.cell(row=5, column=3,
            value='=IF(B5>MonthlyRent,"EXCEEDS ONE MONTH\'S RENT - NOT ALLOWED IN MA","OK")')
c.font = F_BOLD
label_input(ws, 6, "Date received", None, DATE_FMT)
label_input(ws, 7, "Bank name (must be a Massachusetts bank)", None)
label_input(ws, 8, "Account number (separate, interest-bearing)", None)
label_input(ws, 9, "Bank interest rate (annual)", 0.0, "0.00%",
            "Tenant is owed 5%/yr OR the actual bank rate if properly deposited.")
label_input(ws, 10, "Tenancy start date", None, DATE_FMT)
label_input(ws, 11, "Last month's rent collected (if any)", 3200, MONEY,
            "Received 09/19/2026. Also earns interest for the tenant annually.")

ws.cell(row=13, column=1, value="COMPLIANCE CHECKLIST").font = F_SUBHDR
ws.cell(row=13, column=1).fill = FILL_SUB
header_row(ws, 14, ["Requirement", "Deadline", "", "Done? (Y/N)"])
reqs = [
    ("Give receipt (amount, date, your name) when deposit is paid", "=B6"),
    ("Deposit into separate interest-bearing MA bank account; give tenant bank name, address, account #", '=IF(B6="","",B6+30)'),
    ("Give Statement of Condition to tenant", '=IF(B10="","",B10+10)'),
    ("Pay or credit deposit interest to tenant (every anniversary)", '=IF(B10="","",EDATE(B10,12))'),
    ("Pay interest on last month's rent (every anniversary)", '=IF(B10="","",EDATE(B10,12))'),
]
for i, (req, ded) in enumerate(reqs):
    row = 15 + i
    ws.cell(row=row, column=1, value=req).alignment = Alignment(wrap_text=True)
    c = ws.cell(row=row, column=2, value=ded)
    c.number_format = DATE_FMT
    c.border = BORDER
    c = ws.cell(row=row, column=4)
    c.fill = FILL_IN
    c.border = BORDER
dv_yn2 = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
ws.add_data_validation(dv_yn2)
dv_yn2.add("D15:D19")

ws.cell(row=21, column=1, value="MOVE-OUT").font = F_SUBHDR
ws.cell(row=21, column=1).fill = FILL_SUB
label_input(ws, 22, "Move-out date", None, DATE_FMT)
ws.cell(row=23, column=1, value="Deposit + interest must be returned by").font = F_BOLD
c = ws.cell(row=23, column=2, value='=IF(B22="","",B22+30)')
c.number_format = DATE_FMT
c.font = F_BOLD
ws.cell(row=24, column=1, value="Interest accrued (estimate)")
c = ws.cell(row=24, column=2,
            value='=IF(OR(B6="",B5=0),"",ROUND(B5*B9*YEARFRAC(B6,IF(B22="",TODAY(),B22)),2))')
c.number_format = MONEY

ws.cell(row=26, column=1, value="Itemized deductions at move-out (damage beyond wear & tear, unpaid rent):").font = F_BOLD
header_row(ws, 27, ["Description", "Amount"], fill=FILL_SUB, font=F_SUBHDR)
for i in range(5):
    row = 28 + i
    ws.cell(row=row, column=1).fill = FILL_IN
    ws.cell(row=row, column=1).border = BORDER
    c = ws.cell(row=row, column=2, value=None)
    c.fill = FILL_IN
    c.border = BORDER
    c.number_format = MONEY
ws.cell(row=33, column=1, value="AMOUNT TO RETURN").font = F_BOLD
c = ws.cell(row=33, column=2, value='=B5+IF(B24="",0,B24)-SUM(B28:B32)')
c.number_format = MONEY
c.font = F_BOLD
c.fill = FILL_SUB
c.border = BORDER
ws.cell(row=34, column=1,
        value="Deductions require a sworn, itemized statement with receipts/estimates within 30 days.").font = F_NOTE

# =================================================================
# Mileage
# =================================================================
ws = wb.create_sheet("Mileage")
header_row(ws, 1, ["Date", "From", "To", "Purpose", "Miles",
                   "Deduction (auto)"])
set_widths(ws, {"A": 12, "B": 22, "C": 22, "D": 32, "E": 10, "F": 16,
                "G": 4, "H": 26, "I": 14})
ws.freeze_panes = "A2"
for i in range(MILE_ROWS):
    row = 2 + i
    ws.cell(row=row, column=1).number_format = DATE_FMT
    ws.cell(row=row, column=6, value=f'=IF($E{row}="","",$E{row}*MileageRate)')
    ws.cell(row=row, column=6).number_format = MONEY
ws.cell(row=2, column=8, value="Deduction for tax year:").font = F_BOLD
c = ws.cell(row=2, column=9,
            value='=SUMIFS(MileDed,MileDates,">="&DATE(TaxYear,1,1),MileDates,"<="&DATE(TaxYear,12,31))')
c.font = F_BOLD
c.number_format = MONEY

# ---------------------------------------------------------------- save
# Tabs grouped by function: blue = reporting, amber = inputs,
# violet = calculations; ordered reporting -> inputs -> calculations.
TAB_GROUPS = {
    "898781": ["Instructions"],
    BLUE: ["Cash Flow", "Tax Summary"],
    "EDA100": ["Settings", "Income", "Expenses", "Tenants", "Deposit",
               "Cap Improvements", "Mileage"],
    "4A3AA7": ["Depreciation"],
}
for color, names in TAB_GROUPS.items():
    for name in names:
        wb[name].sheet_properties.tabColor = color

order = ["Instructions", "Cash Flow", "Tax Summary", "Settings", "Income",
         "Expenses", "Tenants", "Deposit", "Cap Improvements", "Mileage",
         "Depreciation"]
wb._sheets = [wb[name] for name in order]
wb.save(OUT)
print(f"wrote {OUT}")
