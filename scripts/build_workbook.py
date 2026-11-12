"""
build_workbook.py
-----------------
Builds the Excel workbook from the CSV files produced by generate_data.py.

    python scripts/generate_data.py
    python scripts/build_workbook.py

Output: output/DC_Asset_Lifecycle_Financial_Dashboard.xlsx

Design rules
* Raw records are written as values (blue text); every metric is an Excel formula,
  so the workbook recalculates when records or settings change.
* Lookups use INDEX/MATCH (works in every Excel version and in LibreOffice).
* Key columns are exposed as workbook names (e.g. AR_Status) so formulas read like sentences.
* Native PivotTables cannot be created with openpyxl, so the Pivot_Analysis sheet gives exact
  build steps plus clearly labelled formula cross-check grids (they are NOT PivotTables).
"""

import json
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "output" / "DC_Asset_Lifecycle_Financial_Dashboard.xlsx"
META = json.loads((DATA / "meta.json").read_text())

HDR_ROW, FIRST, LAST = 4, 5, 1000          # table header row, first data row, formula range end
SITES = ["HEL-01", "AMS-01", "FRA-01", "STO-01"]
SITE_NAMES = {"HEL-01": "Helsinki (FI)", "AMS-01": "Amsterdam (NL)", "FRA-01": "Frankfurt (DE)",
              "STO-01": "Stockholm (SE)"}
CATEGORIES = ["GPU Server", "GPU Module", "Storage Server", "Network Switch", "Optical Transceiver",
              "Motherboard", "Power Supply Unit", "Other Infrastructure"]
STATUSES = ["Deployed", "Available", "Under Maintenance", "Under RMA", "Retired"]
TEAMS = ["GPU Compute Ops", "Storage Ops", "Network Engineering", "DC Facilities", "Hardware Spares Pool",
         "Asset Disposal"]
COMPONENTS = ["GPU", "Motherboard", "Power Supply Unit", "Optical Transceiver", "Network Switch",
              "Storage Controller", "Memory DIMM", "Fan Module", "NVMe Drive"]
VENDOR_STATUS = ["Submitted", "Approved - Advance Replacement", "Replacement Shipped", "Replacement Received",
                 "Repaired", "Rejected - Out of Warranty"]
CASE_STATUS = ["Open", "Awaiting Replacement", "Closed"]
PROC_CATS = ["Equipment Purchase", "Spare Parts Replenishment", "Maintenance Services"]
STAGES = ["Pending Approval", "Approved - No PO", "PO Issued", "Received", "Invoiced", "Cancelled"]
INVOICE_STATUS = ["Not Invoiced", "Invoiced", "N/A"]
ACTION_STATUS = ["Not Required", "Open", "In Progress", "Resolved"]

# ---------------------------------------------------------------- styles
FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_CALC = Font(name=FONT, size=10, color="000000")
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=16, bold=True, color="1F3864")
F_SUB = Font(name=FONT, size=10, italic=True, color="595959")
F_H2 = Font(name=FONT, size=12, bold=True, color="1F3864")
F_BOLD = Font(name=FONT, size=10, bold=True)
FILL_IN_HDR = PatternFill("solid", fgColor="1F3864")     # input column header (navy)
FILL_CALC_HDR = PatternFill("solid", fgColor="2E6B5E")   # calculated column header (green-teal)
FILL_LINK_HDR = PatternFill("solid", fgColor="5B6B2E")   # lookup column header (olive)
FILL_SETTING = PatternFill("solid", fgColor="FFFF00")
FILL_LIGHT = PatternFill("solid", fgColor="F2F2F2")
FILL_CARD = PatternFill("solid", fgColor="EEF3FA")
FILL_RED = PatternFill("solid", fgColor="F8CBAD")
FILL_AMBER = PatternFill("solid", fgColor="FFE699")
FILL_GREEN = PatternFill("solid", fgColor="C6EFCE")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

FMT_DATE = "dd-mmm-yyyy"
FMT_EUR = '"€"#,##0;[Red]("€"#,##0);"-"'
FMT_INT = '#,##0;[Red](#,##0);"0"'
FMT_PCT = "0.0%"
FMT_DEC = "0.0"

wb = Workbook()
wb.calculation.fullCalcOnLoad = True    # Excel recalculates everything when the file opens


def name(nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def col_range(sheet, col):
    return f"'{sheet}'!${col}${FIRST}:${col}${LAST}" if " " in sheet else f"{sheet}!${col}${FIRST}:${col}${LAST}"


def load(fname, dates=()):
    df = pd.read_csv(DATA / fname, parse_dates=list(dates))
    return df.astype(object).where(pd.notna(df), None)


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    ws["A2"] = sub
    ws["A2"].font = F_SUB
    ws.sheet_view.showGridLines = False


def build_table(ws, tname, cols, df):
    """cols: list of (header, kind, width, number_format, formula_template or None).
    kind: 'in' input value from df, 'calc' formula, 'link' formula looking up another sheet."""
    for j, (hdr, kind, width, fmt, f) in enumerate(cols, start=1):
        c = ws.cell(row=HDR_ROW, column=j, value=hdr)
        c.font = F_HDR
        c.fill = {"in": FILL_IN_HDR, "calc": FILL_CALC_HDR, "link": FILL_LINK_HDR}[kind]
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BOX
        ws.column_dimensions[get_column_letter(j)].width = width
    ws.row_dimensions[HDR_ROW].height = 42
    for i, rec in enumerate(df.to_dict("records")):
        r = FIRST + i
        for j, (hdr, kind, width, fmt, f) in enumerate(cols, start=1):
            if kind == "in":
                v = rec.get(hdr)
                if isinstance(v, pd.Timestamp):
                    v = v.date()
                c = ws.cell(row=r, column=j, value=v)
                c.font = F_INPUT
            else:
                c = ws.cell(row=r, column=j, value=f.format(r=r))
                c.font = F_LINK if kind == "link" else F_CALC
            if fmt:
                c.number_format = fmt
            c.border = BOX
    last_row = FIRST + len(df) - 1
    ref = f"A{HDR_ROW}:{get_column_letter(len(cols))}{last_row}"
    t = Table(displayName=tname, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    return last_row


def dv_list(ws, items, col, last_row, allow_blank=True):
    dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=allow_blank,
                        showErrorMessage=True, errorTitle="Invalid value",
                        error="Choose a value from the drop-down list (see Data_Dictionary).")
    ws.add_data_validation(dv)
    dv.add(f"{col}{FIRST}:{col}{LAST}")


def cf_text(ws, rng, text, fill):
    first = rng.split(":")[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{text}"'], fill=fill))


def summary_block(ws, col, row, heading, items, label_w=34, val_w=16):
    """Write a small label/value block. items: list of (label, formula, fmt)."""
    lc, vc = col, get_column_letter(ws[col + "1"].column + 1)
    ws.column_dimensions[lc].width = label_w
    ws.column_dimensions[vc].width = val_w
    ws[f"{lc}{row}"] = heading
    ws[f"{lc}{row}"].font = F_H2
    cells = {}
    for k, (label, formula, fmt) in enumerate(items, start=1):
        r = row + k
        ws[f"{lc}{r}"] = label
        ws[f"{lc}{r}"].font = F_BASE
        ws[f"{lc}{r}"].border = BOX
        ws[f"{lc}{r}"].fill = FILL_LIGHT
        ws[f"{vc}{r}"] = formula
        ws[f"{vc}{r}"].font = F_BOLD
        ws[f"{vc}{r}"].border = BOX
        ws[f"{vc}{r}"].number_format = fmt or "General"
        cells[label] = f"{vc}{r}"
    return cells


# =====================================================================================
# Sheet 1: README (settings live here so they are explained next to where they are set)
# =====================================================================================
ws_readme = wb.active
ws_readme.title = "README"

# =====================================================================================
# Sheet 2: Asset_Register
# =====================================================================================
ws = wb.create_sheet("Asset_Register")
title(ws, "Asset Register: serialized IT assets (simulated)",
      "One row per tagged asset (unit cost at or above the tagging threshold). Blue = input, black = formula, "
      "green = lookup. Settings used by formulas are on the README sheet.")
assets = load("asset_register.csv", ["Purchase Date", "Warranty Expiration", "Last Audit Date", "Disposal Date"])
AR = "Asset_Register"
ar_cols = [
    ("Asset ID", "in", 11, None, None),
    ("Serial Number", "in", 14, None, None),
    ("Category", "in", 19, None, None),
    ("Manufacturer", "in", 13, None, None),
    ("Model", "in", 36, None, None),
    ("Location", "in", 10, None, None),
    ("Rack Position", "in", 18, None, None),
    ("Purchase Date", "in", 12, FMT_DATE, None),
    ("Purchase Order ID", "in", 14, None, None),
    ("Acquisition Cost (EUR)", "in", 13, FMT_EUR, None),
    ("Warranty Expiration", "in", 12, FMT_DATE, None),
    ("Status", "in", 17, None, None),
    ("Assigned Team", "in", 20, None, None),
    ("Last Audit Date", "in", 12, FMT_DATE, None),
    ("Disposal Date", "in", 12, FMT_DATE, None),
    # calculated columns P..Z
    ("Asset Age (Years)", "calc", 9, FMT_DEC, '=IF(H{r}="","",ROUND((RefDate-H{r})/365.25,1))'),
    ("Days to Warranty Expiry", "calc", 10, FMT_INT, '=IF(K{r}="","",K{r}-RefDate)'),
    ("Warranty Flag", "calc", 12, None,
     '=IF(L{r}="Retired","N/A",IF(K{r}="","Missing",IF(K{r}<RefDate,"Expired",'
     'IF(K{r}-RefDate<=Warranty_Window_Days,"Expiring","In Warranty"))))'),
    ("Lifecycle Stage", "calc", 15, None,
     '=IF(L{r}="Retired","Retired",IF(AND(P{r}<>"",P{r}>=Refresh_Age_Years),"Refresh Candidate",'
     'IF(R{r}="Expired","Post-Warranty","In Warranty")))'),
    ("Missing Mandatory Fields", "calc", 10, "0",
     '=COUNTBLANK(A{r}:F{r})+COUNTBLANK(H{r}:N{r})+IF(AND(L{r}="Deployed",G{r}=""),1,0)'
     '+IF(AND(L{r}="Retired",O{r}=""),1,0)'),
    ("Serial Check", "calc", 11, None, f'=IF(B{{r}}="","",IF(COUNTIF($B${FIRST}:$B${LAST},B{{r}})>1,"Duplicate","Unique"))'),
    ("Days Since Audit", "calc", 9, FMT_INT, '=IF(N{r}="","",RefDate-N{r})'),
    ("Inspection Required", "calc", 11, None,
     '=IF(L{r}="Retired","No",IF(N{r}="","Yes",IF(OR(V{r}>Audit_Interval_Days,L{r}="Under Maintenance",'
     'U{r}="Duplicate"),"Yes","No")))'),
    ("Operational Availability", "calc", 13, None,
     '=IF(L{r}="Deployed","In Service",IF(L{r}="Available","Ready Spare",IF(L{r}="Retired","Out of Fleet","Unavailable")))'),
    ("Open RMA Linked", "link", 9, None,
     '=IF(COUNTIFS(RMA_Asset,A{r},RMA_Status,"<>Closed")>0,"Yes","No")'),
    ("Record Check", "calc", 18, None,
     '=IF(T{r}>0,"Missing data",IF(U{r}="Duplicate","Duplicate serial",'
     'IF(AND(L{r}="Under RMA",Y{r}="No"),"Under RMA, no open case","OK")))'),
]
ar_last = build_table(ws, "tblAssets", ar_cols, assets)
for col, items in [("C", CATEGORIES), ("F", SITES), ("L", STATUSES), ("M", TEAMS)]:
    dv_list(ws, items, col, ar_last)
ws.conditional_formatting.add(f"R{FIRST}:R{LAST}", FormulaRule(formula=[f'R{FIRST}="Expiring"'], fill=FILL_AMBER))
ws.conditional_formatting.add(f"R{FIRST}:R{LAST}", FormulaRule(formula=[f'R{FIRST}="Expired"'], fill=FILL_RED))
ws.conditional_formatting.add(f"T{FIRST}:T{LAST}", CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_RED))
ws.conditional_formatting.add(f"A{FIRST}:O{LAST}",                       # blank mandatory cells
                              FormulaRule(formula=[f'AND($A{FIRST}<>"",A{FIRST}="",COLUMN(A{FIRST})<>7,COLUMN(A{FIRST})<>15)'],
                                          fill=FILL_RED))
ws.conditional_formatting.add(f"G{FIRST}:G{LAST}", FormulaRule(formula=[f'AND($L{FIRST}="Deployed",G{FIRST}="")'], fill=FILL_RED))
ws.conditional_formatting.add(f"O{FIRST}:O{LAST}", FormulaRule(formula=[f'AND($L{FIRST}="Retired",O{FIRST}="")'], fill=FILL_RED))
ws.conditional_formatting.add(f"B{FIRST}:B{LAST}", FormulaRule(formula=[f'$U{FIRST}="Duplicate"'], fill=FILL_RED))
cf_text(ws, f"U{FIRST}:U{LAST}", "Duplicate", FILL_RED)
cf_text(ws, f"W{FIRST}:W{LAST}", "Yes", FILL_AMBER)
cf_text(ws, f"X{FIRST}:X{LAST}", "Unavailable", FILL_AMBER)
cf_text(ws, f"X{FIRST}:X{LAST}", "In Service", FILL_GREEN)
ws.conditional_formatting.add(f"Z{FIRST}:Z{LAST}", FormulaRule(formula=[f'AND(Z{FIRST}<>"OK",Z{FIRST}<>"")'], fill=FILL_RED))

for nm, col in [("AR_ID", "A"), ("AR_Serial", "B"), ("AR_Category", "C"), ("AR_Manufacturer", "D"),
                ("AR_Location", "F"), ("AR_Cost", "J"), ("AR_Warranty", "K"), ("AR_Status", "L"),
                ("AR_WarrantyFlag", "R"), ("AR_Lifecycle", "S"), ("AR_MissingFields", "T"), ("AR_SerialCheck", "U"),
                ("AR_Inspect", "W"), ("AR_Check", "Z")]:
    name(nm, col_range(AR, col))

# =====================================================================================
# Sheet 3: Spare_Stock
# =====================================================================================
ws = wb.create_sheet("Spare_Stock")
title(ws, "Spare-Parts Stock: bulk, non-serialized spares by site (simulated)",
      "Only parts below the tagging threshold live here; tagged spares (e.g. GPU modules) are tracked one by one "
      "in Asset_Register, so nothing is counted twice. Stock On Order is pulled from Procurement (stage = PO Issued).")
spares = load("spare_stock.csv")
SS = "Spare_Stock"
ss_cols = [
    ("SKU", "in", 13, None, None),
    ("Site", "in", 9, None, None),
    ("Equipment Category", "in", 17, None, None),
    ("Description", "in", 34, None, None),
    ("Current Stock (On Hand)", "in", 10, FMT_INT, None),
    ("Reserved Stock", "in", 10, FMT_INT, None),
    ("Stock On Order", "link", 10, FMT_INT,
     '=SUMIFS(PR_Qty,PR_SKU,A{r},PR_Site,B{r},PR_Stage,"PO Issued")'),
    ("Avg Monthly Consumption", "in", 11, FMT_DEC, None),
    ("Supplier Lead Time (Days)", "in", 10, FMT_INT, None),
    ("Lead Time (Months)", "calc", 10, "0.00", '=I{r}/Days_Per_Month'),
    ("Safety Stock", "in", 9, FMT_INT, None),
    ("Reorder Point", "calc", 10, FMT_DEC, '=H{r}*J{r}+K{r}'),
    ("Target Stock", "in", 9, FMT_INT, None),
    ("Inventory Position", "calc", 10, FMT_INT, '=E{r}+G{r}-F{r}'),
    ("Months of Cover", "calc", 9, FMT_DEC, '=IF(H{r}=0,"",N{r}/H{r})'),
    ("Stock Status", "calc", 10, None, '=IF(N{r}<=K{r},"Critical",IF(N{r}<=L{r},"Reorder","OK"))'),
    ("Reorder Flag", "calc", 9, None, '=IF(N{r}<=L{r},"YES","NO")'),
    ("Suggested Reorder Qty", "calc", 10, FMT_INT, '=IF(Q{r}="YES",MAX(0,ROUNDUP(M{r}-N{r},0)),0)'),
    ("Unit Cost (EUR)", "in", 10, FMT_EUR, None),
    ("Est. Replenishment Cost (EUR)", "calc", 13, FMT_EUR, '=R{r}*S{r}'),
    ("Data Check", "calc", 14, None,
     '=IF(OR(E{r}<0,F{r}<0,G{r}<0),"Negative qty",IF(F{r}>E{r},"Reserved > on hand",'
     'IF(S{r}>=Tagging_Threshold_EUR,"Should be tagged","OK")))'),
]
ss_last = build_table(ws, "tblSpares", ss_cols, spares)
dv_list(ws, SITES, "B", ss_last)
cf_text(ws, f"P{FIRST}:P{LAST}", "Critical", FILL_RED)
cf_text(ws, f"P{FIRST}:P{LAST}", "Reorder", FILL_AMBER)
cf_text(ws, f"Q{FIRST}:Q{LAST}", "YES", FILL_AMBER)
ws.conditional_formatting.add(f"U{FIRST}:U{LAST}", FormulaRule(formula=[f'AND(U{FIRST}<>"OK",U{FIRST}<>"")'], fill=FILL_RED))
for nm, col in [("SS_SKU", "A"), ("SS_Site", "B"), ("SS_Category", "C"), ("SS_Status", "P"), ("SS_Flag", "Q"),
                ("SS_Qty", "R"), ("SS_Cost", "T"), ("SS_Check", "U")]:
    name(nm, col_range(SS, col))

# summary to the right
sc = summary_block(ws, "W", 4, "Replenishment summary", [
    ("SKU-site lines tracked", '=COUNTA(SS_SKU)', "0"),
    ("Lines at/below reorder point", '=COUNTIF(SS_Flag,"YES")', "0"),
    ("  of which Critical (at/below safety stock)", '=COUNTIF(SS_Status,"Critical")', "0"),
    ("Part types (SKUs) needing replenishment", '=COUNTIF($X$13:$X$26,">0")', "0"),
    ("Suggested order quantity (units)", '=SUM(SS_Qty)', FMT_INT),
    ("Estimated replenishment cost", '=SUM(SS_Cost)', FMT_EUR),
])
ws["W11"] = "By part type"
ws["W11"].font = F_H2
for j, h in enumerate(["SKU", "Lines flagged", "Order qty", "Est. cost (EUR)"]):
    c = ws.cell(row=12, column=23 + j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_CALC_HDR, BOX
skus = list(dict.fromkeys(spares["SKU"]))
for i, sku in enumerate(skus):
    r = 13 + i
    ws.cell(row=r, column=23, value=sku).font = F_BASE
    ws.cell(row=r, column=24, value=f'=COUNTIFS(SS_SKU,W{r},SS_Flag,"YES")')
    ws.cell(row=r, column=25, value=f'=SUMIFS(SS_Qty,SS_SKU,W{r})').number_format = FMT_INT
    ws.cell(row=r, column=26, value=f'=SUMIFS(SS_Cost,SS_SKU,W{r})').number_format = FMT_EUR
    for cc in range(23, 27):
        ws.cell(row=r, column=cc).border = BOX
ws["W30"] = "By site"
ws["W30"].font = F_H2
for j, h in enumerate(["Site", "Lines flagged", "Critical lines", "Est. cost (EUR)"]):
    c = ws.cell(row=31, column=23 + j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_CALC_HDR, BOX
for i, s in enumerate(SITES):
    r = 32 + i
    ws.cell(row=r, column=23, value=s)
    ws.cell(row=r, column=24, value=f'=COUNTIFS(SS_Site,W{r},SS_Flag,"YES")')
    ws.cell(row=r, column=25, value=f'=COUNTIFS(SS_Site,W{r},SS_Status,"Critical")')
    ws.cell(row=r, column=26, value=f'=SUMIFS(SS_Cost,SS_Site,W{r})').number_format = FMT_EUR
    for cc in range(23, 27):
        ws.cell(row=r, column=cc).border = BOX
ws.column_dimensions["X"].width = 13
ws.column_dimensions["Y"].width = 13
ws.column_dimensions["Z"].width = 15
ws["W38"] = ("Formulas: Reorder Point = Avg Monthly Consumption x Lead Time (Months) + Safety Stock. "
             "Inventory Position = On Hand + On Order - Reserved. Suggested Qty = MAX(0, ROUNDUP(Target - Position)) "
             "only when Position <= Reorder Point.")
ws["W38"].alignment = WRAP
ws.merge_cells("W38:Z42")
name("SS_FlaggedSKUs", "Spare_Stock!$X$8")

# =====================================================================================
# Sheet 4: RMA_Tracker
# =====================================================================================
ws = wb.create_sheet("RMA_Tracker")
title(ws, "RMA Tracker: return merchandise authorization cases (simulated)",
      "Each case links to an Asset ID in Asset_Register. Asset Category, Manufacturer, Site and warranty status are "
      "looked up, not typed. Target resolution period is the RMA_Target_Days setting on README.")
rma = load("rma_cases.csv", ["Date Reported", "Case Created Date", "Closure Date"])
lk = '=IFERROR(INDEX(AR_{f},MATCH($B{{r}},AR_ID,0)),"Not found")'
rma_cols = [
    ("RMA ID", "in", 11, None, None),
    ("Asset ID", "in", 10, None, None),
    ("Asset Category", "link", 17, None, lk.format(f="Category")),
    ("Component Type", "in", 17, None, None),
    ("Failed Part Serial", "in", 12, None, None),
    ("Manufacturer", "link", 12, None, lk.format(f="Manufacturer")),
    ("Site", "link", 9, None, lk.format(f="Location")),
    ("Failure Description", "in", 40, None, None),
    ("Date Reported", "in", 12, FMT_DATE, None),
    ("Case Created Date", "in", 12, FMT_DATE, None),
    ("Warranty Status", "link", 13, None,
     '=IFERROR(IF(INDEX(AR_Warranty,MATCH($B{r},AR_ID,0))>=I{r},"In Warranty","Out of Warranty"),"Not found")'),
    ("Vendor Status", "in", 22, None, None),
    ("Case Status", "in", 14, None, None),
    ("Spare Swap Asset ID", "in", 11, None, None),
    ("Replacement Tracking No.", "in", 15, None, None),
    ("Closure Date", "in", 12, FMT_DATE, None),
    ("Target Resolution (Days)", "calc", 10, "0", "=RMA_Target_Days"),
    ("Days Open", "calc", 8, FMT_INT, '=IF(P{r}="",RefDate-J{r},P{r}-J{r})'),
    ("SLA Flag", "calc", 12, None, '=IF(M{r}="Closed","Closed",IF(R{r}>Q{r},"Overdue","Within Target"))'),
    ("Age Bucket", "calc", 11, None,
     '=IF(M{r}="Closed","Closed",IF(R{r}<=15,"0-15 days",IF(R{r}<=30,"16-30 days",IF(R{r}<=60,"31-60 days","61+ days"))))'),
    ("Data Check", "calc", 18, None,
     '=IF(ISNA(MATCH(B{r},AR_ID,0)),"Invalid Asset ID",IF(AND(M{r}="Closed",P{r}=""),"Closed, no closure date",'
     'IF(AND(M{r}<>"Closed",P{r}<>""),"Open with closure date",IF(J{r}<I{r},"Created before reported",'
     'IF(AND(N{r}<>"",ISNA(MATCH(N{r},AR_ID,0))),"Invalid swap Asset ID","OK")))))'),
]
rma_last = build_table(ws, "tblRMA", rma_cols, rma)
dv_list(ws, COMPONENTS, "D", rma_last)
dv_list(ws, VENDOR_STATUS, "L", rma_last)
dv_list(ws, CASE_STATUS, "M", rma_last)
cf_text(ws, f"S{FIRST}:S{LAST}", "Overdue", FILL_RED)
cf_text(ws, f"S{FIRST}:S{LAST}", "Within Target", FILL_AMBER)
cf_text(ws, f"K{FIRST}:K{LAST}", "Out of Warranty", FILL_AMBER)
ws.conditional_formatting.add(f"U{FIRST}:U{LAST}", FormulaRule(formula=[f'AND(U{FIRST}<>"OK",U{FIRST}<>"")'], fill=FILL_RED))
RM = "RMA_Tracker"
for nm, col in [("RMA_ID", "A"), ("RMA_Asset", "B"), ("RMA_Category", "C"), ("RMA_Component", "D"),
                ("RMA_Site", "G"), ("RMA_Warranty", "K"), ("RMA_Status", "M"), ("RMA_Closure", "P"),
                ("RMA_DaysOpen", "R"), ("RMA_SLA", "S"), ("RMA_Bucket", "T"), ("RMA_Check", "U")]:
    name(nm, col_range(RM, col))

rc = summary_block(ws, "W", 4, "RMA KPIs", [
    ("Total RMA cases", '=COUNTA(RMA_ID)', "0"),
    ("Open cases (not Closed)", '=COUNTIFS(RMA_ID,"<>",RMA_Status,"<>Closed")', "0"),
    ("Overdue open cases (> target days)", '=COUNTIF(RMA_SLA,"Overdue")', "0"),
    ("Closed cases", '=COUNTIF(RMA_Status,"Closed")', "0"),
    ("% of cases closed", '=IFERROR(X8/X5,0)', FMT_PCT),
    ("Average resolution time, closed (days)", '=IFERROR(AVERAGEIFS(RMA_DaysOpen,RMA_Status,"Closed"),0)', FMT_DEC),
    ("Average age of open cases (days)", '=IFERROR(AVERAGEIFS(RMA_DaysOpen,RMA_Status,"<>Closed",RMA_ID,"<>"),0)', FMT_DEC),
    ("Open cases out of warranty", '=COUNTIFS(RMA_Status,"<>Closed",RMA_Warranty,"Out of Warranty")', "0"),
])
name("KPI_RMA_Total", "RMA_Tracker!$X$5")
name("KPI_RMA_Open", "RMA_Tracker!$X$6")
name("KPI_RMA_Overdue", "RMA_Tracker!$X$7")
ws["W15"] = "Cases by asset category"
ws["W15"].font = F_H2
for j, h in enumerate(["Asset Category", "Open", "Overdue", "All cases", "Assets in register", "RMA rate"]):
    c = ws.cell(row=16, column=23 + j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_CALC_HDR, BOX
    c.alignment = Alignment(wrap_text=True, horizontal="center")
for i, cat in enumerate(CATEGORIES):
    r = 17 + i
    ws.cell(row=r, column=23, value=cat)
    ws.cell(row=r, column=24, value=f'=COUNTIFS(RMA_Category,W{r},RMA_Status,"<>Closed")')
    ws.cell(row=r, column=25, value=f'=COUNTIFS(RMA_Category,W{r},RMA_SLA,"Overdue")')
    ws.cell(row=r, column=26, value=f'=COUNTIF(RMA_Category,W{r})')
    ws.cell(row=r, column=27, value=f'=COUNTIF(AR_Category,W{r})')
    ws.cell(row=r, column=28, value=f'=IFERROR(Z{r}/AA{r},0)').number_format = FMT_PCT
    for cc in range(23, 29):
        ws.cell(row=r, column=cc).border = BOX
ws["W26"] = ("RMA rate = cases in FY2026 / assets currently in the register for that category (a simple exposure ratio, "
             "not an annualised failure rate). Server cases often concern a component inside the server: see Component Type.")
ws["W26"].alignment = WRAP
ws.merge_cells("W26:AB28")
ws["W30"] = "Cases by component type"
ws["W30"].font = F_H2
for j, h in enumerate(["Component Type", "Open", "Closed", "All cases"]):
    c = ws.cell(row=31, column=23 + j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_CALC_HDR, BOX
for i, comp in enumerate(COMPONENTS):
    r = 32 + i
    ws.cell(row=r, column=23, value=comp)
    ws.cell(row=r, column=24, value=f'=COUNTIFS(RMA_Component,W{r},RMA_Status,"<>Closed")')
    ws.cell(row=r, column=25, value=f'=COUNTIFS(RMA_Component,W{r},RMA_Status,"Closed")')
    ws.cell(row=r, column=26, value=f'=COUNTIF(RMA_Component,W{r})')
    for cc in range(23, 27):
        ws.cell(row=r, column=cc).border = BOX
for col, w in [("X", 10), ("Y", 10), ("Z", 10), ("AA", 11), ("AB", 10)]:
    ws.column_dimensions[col].width = w

# =====================================================================================
# Sheet 5: Procurement
# =====================================================================================
ws = wb.create_sheet("Procurement")
title(ws, "Procurement Requests FY2026 (simulated)",
      "Budget treatment: Invoiced = Actual; PO Issued or Received but not invoiced = Commitment; "
      "Pending Approval / Approved - No PO = Pipeline; Cancelled = Excluded. Each request is counted in exactly one bucket.")
proc = load("procurement.csv", ["Request Date", "Expected Delivery Date", "Actual Delivery Date", "Invoice Date"])
pr_cols = [
    ("Request ID", "in", 11, None, None),
    ("Request Date", "in", 12, FMT_DATE, None),
    ("Requesting Site", "in", 10, None, None),
    ("Procurement Category", "in", 22, None, None),
    ("SKU", "in", 12, None, None),
    ("Item Description", "in", 36, None, None),
    ("Quantity", "in", 8, FMT_INT, None),
    ("Est. Unit Price (EUR)", "in", 12, FMT_EUR, None),
    ("Total Request Value (EUR)", "calc", 14, FMT_EUR, '=G{r}*H{r}'),
    ("Approval Stage", "in", 17, None, None),
    ("PO Reference", "in", 12, None, None),
    ("Supplier", "in", 28, None, None),
    ("Expected Delivery Date", "in", 12, FMT_DATE, None),
    ("Actual Delivery Date", "in", 12, FMT_DATE, None),
    ("Invoice Status", "in", 11, None, None),
    ("Invoice Date", "in", 12, FMT_DATE, None),
    ("Budget Treatment", "calc", 12, None,
     '=IF(O{r}="Invoiced","Actual",IF(OR(J{r}="PO Issued",J{r}="Received"),"Commitment",'
     'IF(J{r}="Cancelled","Excluded","Pipeline")))'),
    ("Invoiced Amount (EUR)", "calc", 13, FMT_EUR, '=IF(Q{r}="Actual",I{r},0)'),
    ("Committed Amount (EUR)", "calc", 13, FMT_EUR, '=IF(Q{r}="Commitment",I{r},0)'),
    ("Pipeline Amount (EUR)", "calc", 13, FMT_EUR, '=IF(Q{r}="Pipeline",I{r},0)'),
    ("Indicative Accounting Class", "calc", 16, None,
     '=IF(D{r}="Equipment Purchase","CAPEX (typical)",IF(D{r}="Maintenance Services","OPEX (typical)","Policy-dependent"))'),
    ("Delivery Status", "calc", 11, None,
     '=IF(J{r}="PO Issued",IF(M{r}<RefDate,"Late","On Track"),IF(OR(J{r}="Received",J{r}="Invoiced"),"Delivered","n/a"))'),
    ("Data Check", "calc", 18, None,
     '=IF(AND(OR(J{r}="PO Issued",J{r}="Received",J{r}="Invoiced"),K{r}=""),"PO reference missing",'
     'IF(AND(J{r}="Invoiced",O{r}<>"Invoiced"),"Stage/invoice mismatch",IF(AND(O{r}="Invoiced",P{r}=""),'
     '"Invoice date missing",IF(AND(D{r}="Spare Parts Replenishment",E{r}=""),"SKU missing","OK"))))'),
]
pr_last = build_table(ws, "tblProcurement", pr_cols, proc)
dv_list(ws, SITES, "C", pr_last)
dv_list(ws, PROC_CATS, "D", pr_last)
dv_list(ws, STAGES, "J", pr_last)
dv_list(ws, INVOICE_STATUS, "O", pr_last)
cf_text(ws, f"V{FIRST}:V{LAST}", "Late", FILL_RED)
cf_text(ws, f"Q{FIRST}:Q{LAST}", "Commitment", FILL_AMBER)
ws.conditional_formatting.add(f"W{FIRST}:W{LAST}", FormulaRule(formula=[f'AND(W{FIRST}<>"OK",W{FIRST}<>"")'], fill=FILL_RED))
PR = "Procurement"
for nm, col in [("PR_ID", "A"), ("PR_Site", "C"), ("PR_Category", "D"), ("PR_SKU", "E"), ("PR_Qty", "G"),
                ("PR_Value", "I"), ("PR_Stage", "J"), ("PR_InvoiceDate", "P"), ("PR_Treatment", "Q"),
                ("PR_Invoiced", "R"), ("PR_Committed", "S"), ("PR_Pipeline", "T"), ("PR_Delivery", "V"),
                ("PR_Check", "W")]:
    name(nm, col_range(PR, col))

ws["Y4"] = "Pipeline by approval stage"
ws["Y4"].font = F_H2
for j, h in enumerate(["Approval Stage", "Requests", "Value (EUR)", "Budget treatment"]):
    c = ws.cell(row=5, column=25 + j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_CALC_HDR, BOX
treat = {"Pending Approval": "Pipeline (not approved)", "Approved - No PO": "Pipeline (approved, no PO)",
         "PO Issued": "Commitment", "Received": "Commitment (until invoiced)", "Invoiced": "Actual",
         "Cancelled": "Excluded"}
for i, st in enumerate(STAGES):
    r = 6 + i
    ws.cell(row=r, column=25, value=st)
    ws.cell(row=r, column=26, value=f'=COUNTIF(PR_Stage,Y{r})')
    ws.cell(row=r, column=27, value=f'=SUMIFS(PR_Value,PR_Stage,Y{r})').number_format = FMT_EUR
    ws.cell(row=r, column=28, value=treat[st])
    for cc in range(25, 29):
        ws.cell(row=r, column=cc).border = BOX
ws.cell(row=12, column=25, value="Total (all requests)").font = F_BOLD
ws.cell(row=12, column=26, value="=SUM(Z6:Z11)").font = F_BOLD
ws.cell(row=12, column=27, value="=SUM(AA6:AA11)").number_format = FMT_EUR
for cc in range(25, 29):
    ws.cell(row=12, column=cc).border = BOX
pc = summary_block(ws, "Y", 14, "Budget buckets (reconciliation)", [
    ("Actual expenditure (invoiced)", "=SUM(PR_Invoiced)", FMT_EUR),
    ("Outstanding commitments", "=SUM(PR_Committed)", FMT_EUR),
    ("Pipeline (not yet committed)", "=SUM(PR_Pipeline)", FMT_EUR),
    ("Cancelled (excluded)", '=SUMIFS(PR_Value,PR_Treatment,"Excluded")', FMT_EUR),
    ("Sum of buckets", "=SUM(Z15:Z18)", FMT_EUR),
    ("Difference vs total of all requests", "=Z19-AA12", FMT_EUR),
    ("Total procurement value excl. cancelled", '=SUMIFS(PR_Value,PR_Stage,"<>Cancelled")', FMT_EUR),
    ("Late POs (expected date passed)", '=COUNTIF(PR_Delivery,"Late")', "0"),
], label_w=34)
ws.column_dimensions["Z"].width = 15
ws.column_dimensions["AA"].width = 15
ws.column_dimensions["AB"].width = 26
ws.conditional_formatting.add("Z20", CellIsRule(operator="notEqual", formula=["0"], fill=FILL_RED))
name("KPI_LatePOs", "Procurement!$Z$22")

# =====================================================================================
# Sheet 6: Budget_Forecast
# =====================================================================================
ws = wb.create_sheet("Budget_Forecast")
title(ws, "Budget and Year-End Forecast FY2026 (simulated)",
      "Forecast = Actual + Outstanding Commitments + Additional Expected. Variance = Approved Budget - Forecast "
      "(positive = projected underspend). Additional Expected never includes committed or invoiced amounts.")
budget = load("budget_inputs.csv")
bud_hdr = ["Site", "Budget Category", "Indicative Accounting Class", "Approved Annual Budget (EUR)",
           "Actual Expenditure (EUR)", "Outstanding Commitments (EUR)", "Approved Pipeline, no PO (EUR)",
           "Pending Pipeline x Approval Rate (EUR)", "Other Planned Spend (EUR)", "Additional Expected (EUR)",
           "Forecast Year-End (EUR)", "Projected Variance (EUR)", "Spent to Date %", "Forecast Utilisation %",
           "Risk Flag"]
kinds = ["in", "in", "calc", "in", "link", "link", "link", "link", "in", "calc", "calc", "calc", "calc", "calc", "calc"]
widths = [26, 24, 22, 15, 15, 15, 15, 15, 13, 15, 15, 15, 10, 11, 13]
BH = 5
for j, (h, k, w) in enumerate(zip(bud_hdr, kinds, widths), start=1):
    c = ws.cell(row=BH, column=j, value=h)
    c.font = F_HDR
    c.fill = {"in": FILL_IN_HDR, "calc": FILL_CALC_HDR, "link": FILL_LINK_HDR}[k]
    c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    c.border = BOX
    ws.column_dimensions[get_column_letter(j)].width = w
ws.row_dimensions[BH].height = 54
B0 = BH + 1
for i, rec in enumerate(budget.to_dict("records")):
    r = B0 + i
    vals = [
        (rec["Site"], F_INPUT, None), (rec["Budget Category"], F_INPUT, None),
        (f'=IF(B{r}="Equipment Purchase","CAPEX (typical)",IF(B{r}="Maintenance Services","OPEX (typical)","Policy-dependent"))', F_CALC, None),
        (rec["Approved Annual Budget (EUR)"], F_INPUT, FMT_EUR),
        (f'=SUMIFS(PR_Invoiced,PR_Site,A{r},PR_Category,B{r},PR_InvoiceDate,">="&FY_Start,PR_InvoiceDate,"<="&FY_End)', F_LINK, FMT_EUR),
        (f'=SUMIFS(PR_Committed,PR_Site,A{r},PR_Category,B{r})', F_LINK, FMT_EUR),
        (f'=SUMIFS(PR_Pipeline,PR_Site,A{r},PR_Category,B{r},PR_Stage,"Approved - No PO")', F_LINK, FMT_EUR),
        (f'=SUMIFS(PR_Pipeline,PR_Site,A{r},PR_Category,B{r},PR_Stage,"Pending Approval")*Pending_Approval_Rate', F_LINK, FMT_EUR),
        (rec["Other Planned Spend (EUR)"], F_INPUT, FMT_EUR),
        (f'=G{r}+H{r}+I{r}', F_CALC, FMT_EUR),
        (f'=E{r}+F{r}+J{r}', F_CALC, FMT_EUR),
        (f'=D{r}-K{r}', F_CALC, FMT_EUR),
        (f'=IFERROR(E{r}/D{r},0)', F_CALC, FMT_PCT),
        (f'=IFERROR(K{r}/D{r},0)', F_CALC, FMT_PCT),
        (f'=IF(L{r}<0,"Overrun risk",IF(N{r}>=0.95,"Watch","On track"))', F_CALC, None),
    ]
    for j, (v, f, fmt) in enumerate(vals, start=1):
        c = ws.cell(row=r, column=j, value=v)
        c.font, c.border = f, BOX
        if fmt:
            c.number_format = fmt
BL = B0 + len(budget) - 1            # last budget line row
tbl = Table(displayName="tblBudget", ref=f"A{BH}:O{BL}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
ws.add_table(tbl)
TR = BL + 1
ws.cell(row=TR, column=1, value="Region total").font = F_BOLD
for j in [4, 5, 6, 7, 8, 9, 10, 11, 12]:
    cl = get_column_letter(j)
    c = ws.cell(row=TR, column=j, value=f"=SUM({cl}{B0}:{cl}{BL})")
    c.font, c.number_format, c.border = F_BOLD, FMT_EUR, BOX
ws.cell(row=TR, column=13, value=f"=IFERROR(E{TR}/D{TR},0)").number_format = FMT_PCT
ws.cell(row=TR, column=14, value=f"=IFERROR(K{TR}/D{TR},0)").number_format = FMT_PCT
ws.cell(row=TR, column=15, value=f'=IF(L{TR}<0,"Overrun risk",IF(N{TR}>=0.95,"Watch","On track"))')
for j in range(1, 16):
    ws.cell(row=TR, column=j).fill = FILL_LIGHT
    ws.cell(row=TR, column=j).border = BOX
cf_text(ws, f"O{B0}:O{TR}", "Overrun risk", FILL_RED)
cf_text(ws, f"O{B0}:O{TR}", "Watch", FILL_AMBER)
cf_text(ws, f"O{B0}:O{TR}", "On track", FILL_GREEN)
ws.conditional_formatting.add(f"L{B0}:L{TR}", CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED))
for nm, col in [("BU_Site", "A"), ("BU_Category", "B"), ("BU_Budget", "D"), ("BU_Actual", "E"), ("BU_Commit", "F"),
                ("BU_Additional", "J"), ("BU_Forecast", "K"), ("BU_Variance", "L"), ("BU_Flag", "O")]:
    name(nm, f"Budget_Forecast!${col}${B0}:${col}${BL}")
name("KPI_Budget", f"Budget_Forecast!$D${TR}")
name("KPI_Forecast", f"Budget_Forecast!$K${TR}")
name("KPI_Variance", f"Budget_Forecast!$L${TR}")
name("KPI_Actual", f"Budget_Forecast!$E${TR}")

# Regional summaries
r0 = TR + 3
ws.cell(row=r0, column=1, value="Consolidated regional summary by site").font = F_H2
sum_hdr = ["Site", "Site name", "Approved Budget", "Actual", "Commitments", "Additional Expected", "Forecast Year-End",
           "Projected Variance", "Forecast Utilisation %"]
for j, h in enumerate(sum_hdr, start=1):
    c = ws.cell(row=r0 + 1, column=j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_CALC_HDR, BOX
    c.alignment = Alignment(wrap_text=True, horizontal="center")
SITE_SUM_ROW = r0 + 2
for i, s in enumerate(SITES):
    r = r0 + 2 + i
    ws.cell(row=r, column=1, value=s)
    ws.cell(row=r, column=2, value=SITE_NAMES[s])
    for j, src in zip(range(3, 9), ["BU_Budget", "BU_Actual", "BU_Commit", "BU_Additional", "BU_Forecast", "BU_Variance"]):
        ws.cell(row=r, column=j, value=f"=SUMIFS({src},BU_Site,$A{r})").number_format = FMT_EUR
    ws.cell(row=r, column=9, value=f"=IFERROR(G{r}/C{r},0)").number_format = FMT_PCT
    for j in range(1, 10):
        ws.cell(row=r, column=j).border = BOX
rt = r0 + 6
ws.cell(row=rt, column=1, value="Total").font = F_BOLD
for j in range(3, 9):
    cl = get_column_letter(j)
    ws.cell(row=rt, column=j, value=f"=SUM({cl}{r0 + 2}:{cl}{r0 + 5})").number_format = FMT_EUR
ws.cell(row=rt, column=9, value=f"=IFERROR(G{rt}/C{rt},0)").number_format = FMT_PCT
ws.cell(row=rt + 1, column=1, value="Check: total equals budget table (should be 0)")
ws.cell(row=rt + 1, column=7, value=f"=G{rt}-K{TR}").number_format = FMT_EUR
for j in range(1, 10):
    ws.cell(row=rt, column=j).border = BOX
    ws.cell(row=rt, column=j).fill = FILL_LIGHT
ws.conditional_formatting.add(f"H{r0 + 2}:H{rt}", CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED))
BSUM = (r0 + 2, r0 + 5)

c0 = rt + 3
ws.cell(row=c0, column=1, value="Summary by budget category").font = F_H2
for j, h in enumerate(["Category", "Accounting class", "Approved Budget", "Actual", "Commitments",
                       "Additional Expected", "Forecast Year-End", "Projected Variance", "Forecast Utilisation %"], start=1):
    c = ws.cell(row=c0 + 1, column=j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_CALC_HDR, BOX
    c.alignment = Alignment(wrap_text=True, horizontal="center")
for i, pc_ in enumerate(PROC_CATS):
    r = c0 + 2 + i
    ws.cell(row=r, column=1, value=pc_)
    ws.cell(row=r, column=2, value=["CAPEX (typical)", "Policy-dependent", "OPEX (typical)"][i])
    for j, src in zip(range(3, 9), ["BU_Budget", "BU_Actual", "BU_Commit", "BU_Additional", "BU_Forecast", "BU_Variance"]):
        ws.cell(row=r, column=j, value=f"=SUMIFS({src},BU_Category,$A{r})").number_format = FMT_EUR
    ws.cell(row=r, column=9, value=f"=IFERROR(G{r}/C{r},0)").number_format = FMT_PCT
    for j in range(1, 10):
        ws.cell(row=r, column=j).border = BOX
ws.conditional_formatting.add(f"H{c0 + 2}:H{c0 + 4}", CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED))

# Monthly reporting
m0 = c0 + 7
ws.cell(row=m0, column=1, value="Monthly reporting: invoiced actuals by invoice month vs straight-line phased budget").font = F_H2
mh = ["Month", "Month End", "Period Status"] + SITES + ["Total Actual", "Phased Budget", "Cumulative Actual",
                                                        "Cumulative Phased Budget", "Cumulative Variance"]
for j, h in enumerate(mh, start=1):
    c = ws.cell(row=m0 + 1, column=j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_CALC_HDR, BOX
    c.alignment = Alignment(wrap_text=True, horizontal="center")
for k in range(12):
    r = m0 + 2 + k
    ws.cell(row=r, column=1, value=date(2026, k + 1, 1)).number_format = "mmm-yyyy"
    ws.cell(row=r, column=2, value=f"=EOMONTH(A{r},0)").number_format = FMT_DATE
    ws.cell(row=r, column=3, value=f'=IF(A{r}>RefDate,"Future",IF(B{r}>=RefDate,"Current (partial)","Closed"))')
    for j in range(4, 8):
        cl = get_column_letter(j)
        ws.cell(row=r, column=j, value=(f'=SUMIFS(PR_Invoiced,PR_Site,{cl}${m0 + 1},PR_InvoiceDate,">="&$A{r},'
                                        f'PR_InvoiceDate,"<="&$B{r})')).number_format = FMT_EUR
    ws.cell(row=r, column=8, value=f"=SUM(D{r}:G{r})").number_format = FMT_EUR
    ws.cell(row=r, column=9, value="=KPI_Budget/12").number_format = FMT_EUR
    ws.cell(row=r, column=10, value=f"=H{r}" if k == 0 else f"=J{r - 1}+H{r}").number_format = FMT_EUR
    ws.cell(row=r, column=11, value=f"=I{r}" if k == 0 else f"=K{r - 1}+I{r}").number_format = FMT_EUR
    ws.cell(row=r, column=12, value=f"=K{r}-J{r}").number_format = FMT_EUR
    for j in range(1, 13):
        ws.cell(row=r, column=j).border = BOX
mt = m0 + 14
ws.cell(row=mt, column=1, value="FY total").font = F_BOLD
for j in range(4, 10):
    cl = get_column_letter(j)
    ws.cell(row=mt, column=j, value=f"=SUM({cl}{m0 + 2}:{cl}{m0 + 13})").number_format = FMT_EUR
ws.cell(row=mt + 1, column=1, value="Check: monthly actuals equal budget-table actuals (should be 0)")
ws.cell(row=mt + 1, column=8, value=f"=H{mt}-KPI_Actual").number_format = FMT_EUR
for j in range(1, 13):
    ws.cell(row=mt, column=j).border = BOX
    ws.cell(row=mt, column=j).fill = FILL_LIGHT
ws.conditional_formatting.add(f"C{m0 + 2}:C{m0 + 13}", FormulaRule(formula=[f'C{m0 + 2}="Future"'], font=Font(color="A6A6A6")))
MONTH_ROWS = (m0 + 2, m0 + 13)

n0 = mt + 4
notes = [
    ("CAPEX and OPEX in this model", True),
    ("Capital expenditure (CAPEX) buys or extends an asset that gives economic benefit for more than one year; it is "
     "capitalised on the balance sheet and depreciated over the asset's useful life (for example GPU servers, switches).", False),
    ("Operating expenditure (OPEX) is consumed in the period and expensed immediately in the income statement "
     "(for example maintenance services, repairs, consumables).", False),
    ("Spare parts are deliberately labelled 'Policy-dependent'. Depending on the company's accounting policy and the "
     "applicable standard, spares can be capitalised (e.g. major spares expected to be used over more than one period, "
     "or above a capitalisation threshold) or held as inventory and expensed when consumed. Finance, not the asset "
     "manager, makes this classification.", False),
    ("The 'Indicative Accounting Class' column is a planning aid only; it does not change any calculation.", False),
    ("Simplifying assumptions", True),
    ("Invoiced amounts are treated as actual expenditure; issued but uninvoiced POs (stages PO Issued and Received) are "
     "commitments. A request moves from Commitment to Actual when its Invoice Status becomes 'Invoiced', so it can never "
     "be counted twice.", False),
    ("Additional Expected = approved requests without a PO (100%) + pending requests x Pending_Approval_Rate + other "
     "planned spend (manual estimate). It excludes anything committed or invoiced. All commitments and pipeline are assumed "
     "to be invoiced within FY2026; in practice some would slip into the next year.", False),
    ("Phased budget is straight-line (annual / 12). Real budgets are often phased to the purchase plan.", False),
    ("Budget amounts are fictional and were sized so that some lines show overrun risk and others underspend.", False),
]
for k, (txt, hdr) in enumerate(notes):
    c = ws.cell(row=n0 + k, column=1, value=txt)
    c.font = F_H2 if hdr else F_BASE
    if not hdr:
        c.alignment = WRAP
        ws.merge_cells(start_row=n0 + k, start_column=1, end_row=n0 + k, end_column=12)
        ws.row_dimensions[n0 + k].height = 30
ws.freeze_panes = "C6"

# =====================================================================================
# Sheet 7: Asset_Audits
# =====================================================================================
ws = wb.create_sheet("Asset_Audits")
title(ws, "Physical Asset Audits Aug-Nov 2026 (simulated sample)",
      "Recorded Site/Status are a snapshot of the register at audit time. Audit accuracy describes this sample only, "
      "not the whole fleet. 'Register Now Matches Observed' shows whether the register was corrected afterwards.")
aud = load("asset_audits.csv", ["Audit Date"])
au_cols = [
    ("Audit ID", "in", 11, None, None),
    ("Audit Date", "in", 12, FMT_DATE, None),
    ("Asset ID", "in", 10, None, None),
    ("Category", "link", 17, None, '=IFERROR(INDEX(AR_Category,MATCH($C{r},AR_ID,0)),"Not found")'),
    ("Recorded Site", "in", 10, None, None),
    ("Observed Site", "in", 11, None, None),
    ("Recorded Status", "in", 14, None, None),
    ("Observed Status", "in", 15, None, None),
    ("Auditor", "in", 20, None, None),
    ("Audit Result", "calc", 11, None, '=IF(AND(E{r}=F{r},G{r}=H{r}),"Match","Discrepancy")'),
    ("Discrepancy Type", "calc", 15, None,
     '=IF(J{r}="Match","None",IF(F{r}="Not Located","Not Located",IF(AND(E{r}<>F{r},G{r}<>H{r}),"Location & Status",'
     'IF(E{r}<>F{r},"Location","Status"))))'),
    ("Corrective Action", "in", 52, None, None),
    ("Corrective Action Status", "in", 12, None, None),
    ("Register Now Matches Observed", "link", 12, None,
     '=IFERROR(IF(AND(INDEX(AR_Location,MATCH($C{r},AR_ID,0))=F{r},INDEX(AR_Status,MATCH($C{r},AR_ID,0))=H{r}),"Yes","No"),"Not found")'),
    ("Unresolved Discrepancy", "calc", 11, None, '=IF(AND(J{r}="Discrepancy",M{r}<>"Resolved"),"Yes","No")'),
    ("Asset Value (EUR)", "link", 12, FMT_EUR, '=IFERROR(INDEX(AR_Cost,MATCH($C{r},AR_ID,0)),0)'),
    ("Data Check", "calc", 18, None,
     '=IF(ISNA(MATCH(C{r},AR_ID,0)),"Invalid Asset ID",IF(AND(J{r}="Match",M{r}<>"Not Required"),"Action on a match",'
     'IF(AND(J{r}="Discrepancy",M{r}="Not Required"),"Discrepancy without action",'
     'IF(AND(M{r}="Resolved",N{r}="No"),"Resolved but register differs","OK"))))'),
]
au_last = build_table(ws, "tblAudits", au_cols, aud)
dv_list(ws, SITES + ["Not Located"], "F", au_last)
dv_list(ws, STATUSES + ["Not Located"], "H", au_last)
dv_list(ws, ACTION_STATUS, "M", au_last)
cf_text(ws, f"J{FIRST}:J{LAST}", "Discrepancy", FILL_AMBER)
cf_text(ws, f"O{FIRST}:O{LAST}", "Yes", FILL_RED)
ws.conditional_formatting.add(f"Q{FIRST}:Q{LAST}", FormulaRule(formula=[f'AND(Q{FIRST}<>"OK",Q{FIRST}<>"")'], fill=FILL_RED))
AU = "Asset_Audits"
for nm, col in [("AU_ID", "A"), ("AU_Asset", "C"), ("AU_Result", "J"), ("AU_Type", "K"), ("AU_Action", "M"),
                ("AU_Unresolved", "O"), ("AU_Value", "P"), ("AU_Check", "Q")]:
    name(nm, col_range(AU, col))
ac = summary_block(ws, "S", 4, "Audit KPIs (sample only)", [
    ("Audit records", "=COUNTA(AU_ID)", "0"),
    ("Matching records", '=COUNTIF(AU_Result,"Match")', "0"),
    ("Audit accuracy (matches / records)", "=IFERROR(T6/T5,0)", FMT_PCT),
    ("Discrepancies", '=COUNTIF(AU_Result,"Discrepancy")', "0"),
    ("Unresolved discrepancies", '=COUNTIF(AU_Unresolved,"Yes")', "0"),
    ("Value of assets with unresolved discrepancy", '=SUMIFS(AU_Value,AU_Unresolved,"Yes")', FMT_EUR),
    ("Distinct assets audited", f'=SUMPRODUCT((C{FIRST}:C{LAST}<>"")/COUNTIF(C{FIRST}:C{LAST},C{FIRST}:C{LAST}&""))', "0"),
    ("Active (non-retired) assets in register", '=COUNTIFS(AR_ID,"<>",AR_Status,"<>Retired")', "0"),
    ("Audit coverage of active fleet", "=IFERROR(T11/T12,0)", FMT_PCT),
], label_w=38)
name("KPI_AuditAccuracy", "Asset_Audits!$T$7")
name("KPI_AuditRecords", "Asset_Audits!$T$5")
name("KPI_AuditUnresolved", "Asset_Audits!$T$9")
name("KPI_AuditCoverage", "Asset_Audits!$T$13")
ws["S16"] = "Discrepancies by type"
ws["S16"].font = F_H2
for j, h in enumerate(["Discrepancy Type", "Count", "Unresolved"]):
    c = ws.cell(row=17, column=19 + j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_CALC_HDR, BOX
for i, t_ in enumerate(["Location", "Status", "Location & Status", "Not Located"]):
    r = 18 + i
    ws.cell(row=r, column=19, value=t_)
    ws.cell(row=r, column=20, value=f"=COUNTIF(AU_Type,S{r})")
    ws.cell(row=r, column=21, value=f'=COUNTIFS(AU_Type,S{r},AU_Unresolved,"Yes")')
    for cc in range(19, 22):
        ws.cell(row=r, column=cc).border = BOX
ws.column_dimensions["U"].width = 12
recs = [
    "Recommended corrective actions",
    "1. 'Not Located' assets: search spare cages, RMA shipments and inter-site transfer logs; if not found within 10 "
    "working days raise a loss report with Finance and Security.",
    "2. Location mismatches: require a signed transfer record for every inter-site move and update the register on "
    "dispatch, not on arrival.",
    "3. Status mismatches: tie status changes to the change/maintenance ticket so a ticket cannot close until the "
    "register is updated.",
    "4. Re-audit every asset with an unresolved discrepancy within 30 days, then extend the sample to assets flagged "
    "'Inspection Required' in Asset_Register.",
]
for k, t_ in enumerate(recs):
    c = ws.cell(row=24 + k, column=19, value=t_)
    c.font = F_H2 if k == 0 else F_BASE
    if k:
        c.alignment = WRAP
        ws.merge_cells(start_row=24 + k, start_column=19, end_row=24 + k, end_column=21)
        ws.row_dimensions[24 + k].height = 40

# =====================================================================================
# Sheet 8: Management_Dashboard
# =====================================================================================
ws = wb.create_sheet("Management_Dashboard")
ws.sheet_view.showGridLines = False
ws["B1"] = "Regional Data Centre IT Asset Management Dashboard"
ws["B1"].font = Font(name=FONT, size=18, bold=True, color="1F3864")
ws["B2"] = "SIMULATED CASE STUDY: all sites, assets, prices and events are fictional."
ws["B2"].font = Font(name=FONT, size=10, italic=True, color="C00000")
ws["B3"] = '="Reporting date: "&TEXT(RefDate,"dd mmm yyyy")&"   |   RMA target: "&RMA_Target_Days&" days   |   Warranty window: "&Warranty_Window_Days&" days"'
ws["B3"].font = F_SUB
for col in "BCDEF":
    ws.column_dimensions[col].width = 25
ws.column_dimensions["A"].width = 2
ws.column_dimensions["G"].width = 3

kpis = [
    ("Total assets registered", "=COUNTA(AR_ID)", "0",
     '=COUNTIFS(AR_ID,"<>",AR_Status,"<>Retired")&" active, "&COUNTIF(AR_Status,"Retired")&" retired"'),
    ("Deployed assets", '=COUNTIF(AR_Status,"Deployed")', "0",
     '=COUNTIF(AR_Status,"Available")&" ready spares available"'),
    ("Operational utilisation", '=IFERROR(COUNTIF(AR_Status,"Deployed")/COUNTIFS(AR_ID,"<>",AR_Status,"<>Retired"),0)', FMT_PCT,
     '="Servers & switches only: "&TEXT(IFERROR((COUNTIFS(AR_Category,"GPU Server",AR_Status,"Deployed")+COUNTIFS(AR_Category,"Storage Server",AR_Status,"Deployed")+COUNTIFS(AR_Category,"Network Switch",AR_Status,"Deployed"))/(COUNTIFS(AR_Category,"GPU Server",AR_Status,"<>Retired")+COUNTIFS(AR_Category,"Storage Server",AR_Status,"<>Retired")+COUNTIFS(AR_Category,"Network Switch",AR_Status,"<>Retired")),0),"0.0%")'),
    ("Audit accuracy (sample)", "=KPI_AuditAccuracy", FMT_PCT,
     '=KPI_AuditRecords&" records, "&TEXT(KPI_AuditCoverage,"0%")&" of active fleet; "&KPI_AuditUnresolved&" unresolved"'),
    ("Open RMA cases", "=KPI_RMA_Open", "0", '=KPI_RMA_Overdue&" overdue (> "&RMA_Target_Days&" days)"'),
    ("Spare part types to reorder", "=SS_FlaggedSKUs", "0",
     '=COUNTIF(SS_Flag,"YES")&" SKU-site lines, est. "&TEXT(SUM(SS_Cost),"€#,##0")'),
    ("Warranties expiring <= window", '=COUNTIF(AR_WarrantyFlag,"Expiring")', "0",
     '=COUNTIF(AR_WarrantyFlag,"Expired")&" active assets already out of warranty"'),
    ("Approved annual budget", "=KPI_Budget", FMT_EUR, '="Actual to date: "&TEXT(KPI_Actual,"€#,##0")'),
    ("Forecast year-end spend", "=KPI_Forecast", FMT_EUR,
     '="Forecast utilisation: "&TEXT(IFERROR(KPI_Forecast/KPI_Budget,0),"0.0%")'),
    ("Projected budget variance", "=KPI_Variance", FMT_EUR,
     '=IF(KPI_Variance>=0,"Projected underspend","Projected OVERSPEND")&"; "&COUNTIF(BU_Flag,"Overrun risk")&" lines at overrun risk"'),
]
card_cells = {}
for k, (label, f, fmt, note) in enumerate(kpis):
    row = 5 if k < 5 else 10
    col = get_column_letter(2 + (k % 5))
    ws[f"{col}{row}"] = label
    ws[f"{col}{row}"].font = Font(name=FONT, size=9, bold=True, color="595959")
    ws[f"{col}{row + 1}"] = f
    ws[f"{col}{row + 1}"].font = Font(name=FONT, size=18, bold=True, color="1F3864")
    ws[f"{col}{row + 1}"].number_format = fmt
    ws[f"{col}{row + 2}"] = note
    ws[f"{col}{row + 2}"].font = Font(name=FONT, size=8, color="595959")
    ws[f"{col}{row + 2}"].alignment = Alignment(wrap_text=True, vertical="top")
    for rr in range(row, row + 3):
        ws[f"{col}{rr}"].fill = FILL_CARD
        ws[f"{col}{rr}"].border = BOX
        if rr < row + 2:
            ws[f"{col}{rr}"].alignment = Alignment(horizontal="left", vertical="center")
    card_cells[label] = f"{col}{row + 1}"
    ws.row_dimensions[row + 2].height = 26
ws.row_dimensions[6].height = 28
ws.row_dimensions[11].height = 28
# KPI risk colouring
ws.conditional_formatting.add(card_cells["Operational utilisation"], CellIsRule(operator="lessThan", formula=["0.75"], fill=FILL_AMBER))
ws.conditional_formatting.add(card_cells["Audit accuracy (sample)"], CellIsRule(operator="lessThan", formula=["0.95"], fill=FILL_RED))
ws.conditional_formatting.add("F6", FormulaRule(formula=["KPI_RMA_Overdue>0"], fill=FILL_RED))
ws.conditional_formatting.add(card_cells["Spare part types to reorder"], CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_AMBER))
ws.conditional_formatting.add(card_cells["Warranties expiring <= window"], CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_AMBER))
ws.conditional_formatting.add(card_cells["Projected budget variance"], CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED))
ws.conditional_formatting.add(card_cells["Projected budget variance"], CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=FILL_GREEN))

# Risk watchlist (right side)
ws["H4"] = "Operational risk watchlist"
ws["H4"].font = F_H2
ws.column_dimensions["H"].width = 46
ws.column_dimensions["I"].width = 10
watch = [
    ("Overdue open RMA cases", "=KPI_RMA_Overdue"),
    ("Critical spare lines (at/below safety stock)", '=COUNTIF(SS_Status,"Critical")'),
    ("Unresolved audit discrepancies", "=KPI_AuditUnresolved"),
    ("Budget lines at overrun risk", '=COUNTIF(BU_Flag,"Overrun risk")'),
    ("Late purchase orders", "=KPI_LatePOs"),
    ("Assets with missing mandatory data", '=COUNTIF(AR_MissingFields,">0")'),
    ("Duplicate serial numbers (records)", '=COUNTIF(AR_SerialCheck,"Duplicate")'),
    ("Active assets requiring inspection", '=COUNTIFS(AR_Inspect,"Yes",AR_Status,"<>Retired")'),
    ("Assets under RMA without an open case", '=COUNTIF(AR_Check,"Under RMA, no open case")'),
    ("Record-level data checks failing (all sheets)",
     '=COUNTIFS(AR_Check,"<>OK",AR_ID,"<>")+COUNTIFS(RMA_Check,"<>OK",RMA_ID,"<>")+COUNTIFS(PR_Check,"<>OK",PR_ID,"<>")'
     '+COUNTIFS(AU_Check,"<>OK",AU_ID,"<>")+COUNTIFS(SS_Check,"<>OK",SS_SKU,"<>")'),
]
for k, (label, f) in enumerate(watch):
    r = 5 + k
    ws[f"H{r}"] = label
    ws[f"I{r}"] = f
    ws[f"H{r}"].border = ws[f"I{r}"].border = BOX
    ws[f"H{r}"].font = F_BASE
    ws[f"I{r}"].font = F_BOLD
    ws[f"I{r}"].alignment = Alignment(horizontal="center")
ws.conditional_formatting.add("I5:I14", CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_RED))
ws.conditional_formatting.add("I5:I14", CellIsRule(operator="equal", formula=["0"], fill=FILL_GREEN))
ws["H15"] = ("Red = needs action. The data-check line counts deliberate exceptions planted in the simulated data "
             "(duplicate serial, missing fields) plus any errors introduced when editing.")
ws["H15"].font = Font(name=FONT, size=8, italic=True, color="595959")
ws["H15"].alignment = WRAP
ws.row_dimensions[15].height = 36

# Chart source data (formula driven) from row 50
D0 = 50
ws[f"B{D0 - 1}"] = "Chart source data (formula-driven; do not type over)"
ws[f"B{D0 - 1}"].font = F_H2


def mini_table(r, c, headers, rows):
    for j, h in enumerate(headers):
        cell = ws.cell(row=r, column=c + j, value=h)
        cell.font, cell.fill, cell.border = F_HDR, FILL_CALC_HDR, BOX
    for i, rowvals in enumerate(rows, start=1):
        for j, v in enumerate(rowvals):
            cell = ws.cell(row=r + i, column=c + j, value=v.format(r=r + i) if isinstance(v, str) else v)
            cell.border = BOX
            cell.font = F_BASE
    return r + 1, r + len(rows)


s1 = mini_table(D0, 2, ["Site", "Active assets", "Deployed", "Retired"],
                [[s, '=COUNTIFS(AR_Location,B{r},AR_Status,"<>Retired")', '=COUNTIFS(AR_Location,B{r},AR_Status,"Deployed")',
                  '=COUNTIFS(AR_Location,B{r},AR_Status,"Retired")'] for s in SITES])
s2 = mini_table(D0 + 7, 2, ["Status", "Assets"], [[s, '=COUNTIF(AR_Status,B{r})'] for s in STATUSES])
s3 = mini_table(D0 + 15, 2, ["Open case age", "Open RMAs"],
                [[b, '=COUNTIF(RMA_Bucket,B{r})'] for b in ["0-15 days", "16-30 days", "31-60 days", "61+ days"]])
s4 = mini_table(D0 + 22, 2, ["Site", "Approved budget", "Forecast year-end"],
                [[s, f"=Budget_Forecast!C{BSUM[0] + i}", f"=Budget_Forecast!G{BSUM[0] + i}"] for i, s in enumerate(SITES)])
for r in range(s4[0], s4[1] + 1):
    ws[f"C{r}"].number_format = ws[f"D{r}"].number_format = FMT_EUR
# Site x category grid (supports chart 1 reading)
mini_table(D0, 7 + 1, ["Site"] + CATEGORIES,
           [[s] + [f'=COUNTIFS(AR_Location,$H{{r}},AR_Category,"{c}",AR_Status,"<>Retired")' for c in CATEGORIES] for s in SITES])
ws.cell(row=D0 - 1, column=8, value="Active assets by site and category").font = F_BOLD


def bar(title_, cats, series, anchor, horizontal=False, colors=("2A78D6", "EB6834"), y_title=None, fmt=None):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.title = title_
    ch.height, ch.width = 7.5, 14
    ch.legend = None if len(series) == 1 else ch.legend
    if ch.legend is not None:
        ch.legend.position = "b"
    for k, (ref, ttl) in enumerate(series):
        ch.add_data(ref, titles_from_data=ttl)
        s = ch.series[-1]
        s.graphicalProperties.solidFill = colors[k]
        s.graphicalProperties.line.solidFill = colors[k]
    ch.set_categories(cats)
    ch.gapWidth = 60
    if y_title:
        ch.y_axis.title = y_title
    if fmt:
        ch.y_axis.numFmt = fmt

    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ws.add_chart(ch, anchor)


bar("Active assets by site", Reference(ws, min_col=2, min_row=s1[0], max_row=s1[1]),
    [(Reference(ws, min_col=3, min_row=s1[0], max_row=s1[1]), False)], "B17")
bar("Asset status breakdown (all registered assets)", Reference(ws, min_col=2, min_row=s2[0], max_row=s2[1]),
    [(Reference(ws, min_col=3, min_row=s2[0], max_row=s2[1]), False)], "E17", horizontal=True)
bar("Open RMA cases by age", Reference(ws, min_col=2, min_row=s3[0], max_row=s3[1]),
    [(Reference(ws, min_col=3, min_row=s3[0], max_row=s3[1]), False)], "B33")
bar("Approved budget vs forecast year-end by site (EUR)", Reference(ws, min_col=2, min_row=s4[0], max_row=s4[1]),
    [(Reference(ws, min_col=3, min_row=s4[0] - 1, max_row=s4[1]), True),
     (Reference(ws, min_col=4, min_row=s4[0] - 1, max_row=s4[1]), True)], "E33", fmt='"€"#,##0,"k"')

# =====================================================================================
# Sheet 9: Pivot_Analysis
# =====================================================================================
ws = wb.create_sheet("Pivot_Analysis")
title(ws, "Pivot Analysis: build three native PivotTables here",
      "The Python library used (openpyxl) cannot create native PivotTables. The grids below are FORMULA CROSS-CHECKS, "
      "not PivotTables. Build the real PivotTables with the steps below, at the cells given, then compare totals.")
ws.column_dimensions["A"].width = 24
for col in "BCDEFGHIJ":
    ws.column_dimensions[col].width = 13
steps = [
    ("General steps (Excel for Windows / Mac, Microsoft 365)", True),
    ("1. Click any cell inside the source table (e.g. a cell in tblAssets on Asset_Register).", False),
    ("2. Insert > PivotTable > From Table/Range. Excel fills in the table name (tblAssets). Choose 'Existing Worksheet' "
     "and type the location given below (e.g. Pivot_Analysis!$L$6). Click OK.", False),
    ("3. In the PivotTable Fields pane drag fields into Rows, Columns and Values as listed. If Values shows 'Sum of' "
     "for a text field change it via Value Field Settings > Count.", False),
    ("4. Refresh after editing data: right-click the PivotTable > Refresh, or Data > Refresh All (Ctrl+Alt+F5). "
     "Because the source is an Excel Table, new rows added to the table are included automatically on refresh.", False),
    ("5. Optional: PivotTable Analyze > Options > Data > 'Refresh data when opening the file'.", False),
    ("", False),
    ("PivotTable A: asset count by location and category", True),
    ("Source: tblAssets (Asset_Register). Location: Pivot_Analysis!$L$6. Rows = Location. Columns = Category. "
     "Values = Asset ID (Count). Optional filter: Status (exclude Retired to match the dashboard 'active' view).", False),
    ("PivotTable B: invoiced expenditure by location and procurement category", True),
    ("Source: tblProcurement (Procurement). Location: Pivot_Analysis!$L$20. Rows = Requesting Site. "
     "Columns = Procurement Category. Values = Invoiced Amount (EUR) (Sum). Format values as currency with "
     "Value Field Settings > Number Format.", False),
    ("PivotTable C: RMA cases by component type and case status", True),
    ("Source: tblRMA (RMA_Tracker). Location: Pivot_Analysis!$L$34. Rows = Component Type. Columns = Case Status. "
     "Values = RMA ID (Count).", False),
    ("Check: each PivotTable's grand total must equal the cross-check grid total on the left. "
     "If not, refresh the PivotTable and confirm the source is the whole table.", True),
]
for k, (t_, bold) in enumerate(steps):
    c = ws.cell(row=4 + k, column=1, value=t_)
    c.font = F_H2 if bold and t_ else F_BASE
    if t_ and not bold:
        c.alignment = WRAP
        ws.merge_cells(start_row=4 + k, start_column=1, end_row=4 + k, end_column=10)
        ws.row_dimensions[4 + k].height = 30


def grid(r, heading, row_vals, col_vals, f_tpl, fmt="0"):
    ws.cell(row=r, column=1, value=heading).font = F_H2
    ws.cell(row=r + 1, column=1, value="Formula cross-check, not a PivotTable").font = F_SUB
    hdr = r + 2
    ws.cell(row=hdr, column=1, value="Row \\ Column").font = F_HDR
    ws.cell(row=hdr, column=1).fill = FILL_CALC_HDR
    for j, cv in enumerate(col_vals + ["Total"], start=2):
        c = ws.cell(row=hdr, column=j, value=cv)
        c.font, c.fill = F_HDR, FILL_CALC_HDR
        c.alignment = Alignment(wrap_text=True, horizontal="center")
    ws.row_dimensions[hdr].height = 40
    for i, rv in enumerate(row_vals + ["Total"], start=1):
        rr = hdr + i
        ws.cell(row=rr, column=1, value=rv).font = F_BOLD if rv == "Total" else F_BASE
        for j in range(2, len(col_vals) + 3):
            cl = get_column_letter(j)
            if rv == "Total":
                v = f"=SUM({cl}{hdr + 1}:{cl}{rr - 1})"
            elif j == len(col_vals) + 2:
                v = f"=SUM(B{rr}:{get_column_letter(j - 1)}{rr})"
            else:
                v = f_tpl.format(row=f"$A{rr}", col=f"{cl}${hdr}")
            c = ws.cell(row=rr, column=j, value=v)
            c.number_format = fmt
            c.border = BOX
        ws.cell(row=rr, column=1).border = BOX
    return hdr + len(row_vals) + 2


g = grid(20, "A. Asset count by location and category (all statuses)", SITES, CATEGORIES,
         "=COUNTIFS(AR_Location,{row},AR_Category,{col})")
g = grid(g + 2, "B. Invoiced expenditure by site and procurement category (EUR)", SITES, PROC_CATS,
         "=SUMIFS(PR_Invoiced,PR_Site,{row},PR_Category,{col})", FMT_EUR)
grid(g + 2, "C. RMA cases by component type and case status", COMPONENTS, CASE_STATUS,
     "=COUNTIFS(RMA_Component,{row},RMA_Status,{col})")
ws["L4"] = "Build PivotTable A at L6, B at L20 and C at L34 (columns L onwards are intentionally empty)."
ws["L4"].font = Font(name=FONT, size=10, bold=True, color="C00000")

# =====================================================================================
# Sheet 10: Data_Dictionary
# =====================================================================================
ws = wb.create_sheet("Data_Dictionary")
title(ws, "Data Dictionary, KPI Definitions and Allowed Values",
      "I = input typed by a user; C = calculated by formula; L = looked up from another sheet.")
dd = [
    # sheet, field, type, unit, I/C/L, definition
    ("Asset_Register", "Asset ID", "Text", "AST-nnnn", "I", "Unique asset tag. Primary key referenced by RMA_Tracker and Asset_Audits."),
    ("Asset_Register", "Serial Number", "Text", "", "I", "Manufacturer serial (fictional). Must be unique; Serial Check flags duplicates."),
    ("Asset_Register", "Category", "List", "", "I", "Equipment category; see allowed values."),
    ("Asset_Register", "Location", "List", "Site code", "I", "Site where the register says the asset is."),
    ("Asset_Register", "Rack Position", "Text", "Hall-Row-U", "I", "Mandatory when Status = Deployed; spare cage/workbench/RMA staging otherwise."),
    ("Asset_Register", "Purchase Date", "Date", "", "I", "Date the asset was received/capitalised."),
    ("Asset_Register", "Purchase Order ID", "Text", "", "I", "Historical PO that bought the asset (pre-FY2026 POs are not in the Procurement extract)."),
    ("Asset_Register", "Acquisition Cost (EUR)", "Number", "EUR", "I", "Illustrative purchase cost, not a verified market quote."),
    ("Asset_Register", "Warranty Expiration", "Date", "", "I", "End of OEM warranty (purchase date + warranty term)."),
    ("Asset_Register", "Status", "List", "", "I", "Deployed, Available, Under Maintenance, Under RMA, Retired."),
    ("Asset_Register", "Last Audit Date", "Date", "", "I", "Most recent physical verification."),
    ("Asset_Register", "Disposal Date", "Date", "", "I", "Mandatory only when Status = Retired."),
    ("Asset_Register", "Asset Age (Years)", "Number", "Years", "C", "(RefDate - Purchase Date) / 365.25."),
    ("Asset_Register", "Days to Warranty Expiry", "Number", "Days", "C", "Warranty Expiration - RefDate (negative = expired)."),
    ("Asset_Register", "Warranty Flag", "Text", "", "C", "Expired / Expiring (within Warranty_Window_Days) / In Warranty / N/A for retired."),
    ("Asset_Register", "Lifecycle Stage", "Text", "", "C", "Retired, Refresh Candidate (age >= Refresh_Age_Years), Post-Warranty, In Warranty."),
    ("Asset_Register", "Missing Mandatory Fields", "Number", "Count", "C", "Blank mandatory cells (A-F, H-N) + missing rack for deployed + missing disposal date for retired."),
    ("Asset_Register", "Serial Check", "Text", "", "C", "Duplicate if the serial appears more than once (COUNTIF > 1)."),
    ("Asset_Register", "Inspection Required", "Text", "", "C", "Yes if never audited, last audit older than Audit_Interval_Days, under maintenance, or duplicate serial."),
    ("Asset_Register", "Operational Availability", "Text", "", "C", "In Service (Deployed), Ready Spare (Available), Unavailable (Maintenance/RMA), Out of Fleet (Retired)."),
    ("Asset_Register", "Open RMA Linked", "Text", "", "L", "Yes if RMA_Tracker has a non-closed case for this Asset ID."),
    ("Spare_Stock", "Current Stock (On Hand)", "Number", "Units", "I", "Physical units in the site spares store."),
    ("Spare_Stock", "Reserved Stock", "Number", "Units", "I", "On-hand units already allocated to planned work; cannot exceed on hand."),
    ("Spare_Stock", "Stock On Order", "Number", "Units", "L", "SUMIFS of Procurement quantity for this SKU and site where Approval Stage = PO Issued."),
    ("Spare_Stock", "Avg Monthly Consumption", "Number", "Units/month", "I", "Average units issued per month (planning estimate)."),
    ("Spare_Stock", "Lead Time (Months)", "Number", "Months", "C", "Supplier lead time in days / Days_Per_Month."),
    ("Spare_Stock", "Reorder Point", "Number", "Units", "C", "Avg Monthly Consumption x Lead Time (Months) + Safety Stock."),
    ("Spare_Stock", "Inventory Position", "Number", "Units", "C", "On Hand + On Order - Reserved."),
    ("Spare_Stock", "Suggested Reorder Qty", "Number", "Units", "C", "If Position <= Reorder Point: MAX(0, ROUNDUP(Target Stock - Position)); else 0. Never negative."),
    ("RMA_Tracker", "Asset ID", "Text", "", "I", "Asset the case is raised against. For a component inside a server this is the server."),
    ("RMA_Tracker", "Component Type", "List", "", "I", "The failed part (GPU, PSU, DIMM, ...)."),
    ("RMA_Tracker", "Warranty Status", "Text", "", "L", "In Warranty if the asset's warranty expiration >= Date Reported."),
    ("RMA_Tracker", "Case Status", "List", "", "I", "Open, Awaiting Replacement, Closed. Closure Date is filled only when Closed."),
    ("RMA_Tracker", "Spare Swap Asset ID", "Text", "", "I", "Tagged spare installed to restore service before the vendor replacement arrives."),
    ("RMA_Tracker", "Days Open", "Number", "Days", "C", "Closed: Closure Date - Case Created Date. Open: RefDate - Case Created Date."),
    ("RMA_Tracker", "SLA Flag", "Text", "", "C", "Overdue when an open case's Days Open > RMA_Target_Days."),
    ("Procurement", "Total Request Value (EUR)", "Number", "EUR", "C", "Quantity x Est. Unit Price."),
    ("Procurement", "Approval Stage", "List", "", "I", "Pending Approval > Approved - No PO > PO Issued > Received > Invoiced, or Cancelled."),
    ("Procurement", "Budget Treatment", "Text", "", "C", "Actual (invoiced), Commitment (PO Issued/Received, not invoiced), Pipeline (not yet PO), Excluded (cancelled)."),
    ("Procurement", "Invoiced / Committed / Pipeline Amount", "Number", "EUR", "C", "Request value placed in exactly one bucket according to Budget Treatment."),
    ("Procurement", "Delivery Status", "Text", "", "C", "Late if PO Issued and Expected Delivery Date < RefDate."),
    ("Budget_Forecast", "Actual Expenditure", "Number", "EUR", "L", "Invoiced amounts for the site and category with invoice date inside FY_Start..FY_End."),
    ("Budget_Forecast", "Outstanding Commitments", "Number", "EUR", "L", "Committed Amount for the site and category."),
    ("Budget_Forecast", "Additional Expected", "Number", "EUR", "C", "Approved pipeline (no PO) + pending pipeline x Pending_Approval_Rate + Other Planned Spend."),
    ("Budget_Forecast", "Forecast Year-End", "Number", "EUR", "C", "Actual + Commitments + Additional Expected."),
    ("Budget_Forecast", "Projected Variance", "Number", "EUR", "C", "Approved Budget - Forecast. Positive = underspend; negative = overrun risk."),
    ("Asset_Audits", "Recorded Site / Status", "Text", "", "I", "Register values at the time of the audit (snapshot)."),
    ("Asset_Audits", "Observed Site / Status", "List", "", "I", "What the auditor physically found; 'Not Located' if the asset could not be found."),
    ("Asset_Audits", "Audit Result", "Text", "", "C", "Match if both site and status agree; otherwise Discrepancy."),
    ("Asset_Audits", "Unresolved Discrepancy", "Text", "", "C", "Discrepancy whose corrective action is not Resolved."),
]
hdrs = ["Sheet", "Field", "Data Type", "Unit / Format", "I/C/L", "Definition / Formula"]
widths = [16, 30, 10, 13, 6, 100]
for j, (h, w) in enumerate(zip(hdrs, widths), start=1):
    c = ws.cell(row=4, column=j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_IN_HDR, BOX
    ws.column_dimensions[get_column_letter(j)].width = w
for i, rowv in enumerate(dd, start=5):
    for j, v in enumerate(rowv, start=1):
        c = ws.cell(row=i, column=j, value=v)
        c.font, c.border = F_BASE, BOX
        c.alignment = Alignment(wrap_text=True, vertical="top")
k0 = 5 + len(dd) + 2
ws.cell(row=k0, column=1, value="KPI definitions (fixed numerators and denominators)").font = F_H2
kpi_defs = [
    ("Total assets registered", "Count of Asset IDs in Asset_Register", "-", "All statuses including Retired"),
    ("Deployed assets", "Assets with Status = Deployed", "-", ""),
    ("Operational utilisation", "Deployed assets", "Active assets (Status <> Retired)", "Spares, maintenance and RMA units count as idle capacity"),
    ("Audit accuracy", "Audit records with Audit Result = Match", "All audit records in Asset_Audits", "Describes the audited sample only, not the whole fleet"),
    ("Audit coverage", "Distinct assets audited", "Active assets in register", "Shows how representative the accuracy figure is"),
    ("Open RMA cases", "Cases with Case Status <> Closed", "-", ""),
    ("Overdue RMA cases", "Open cases with Days Open > RMA_Target_Days", "-", ""),
    ("% of RMA cases closed", "Closed cases", "All cases", ""),
    ("Average RMA resolution time", "Sum of Days Open for closed cases", "Number of closed cases", "Open cases excluded"),
    ("Spare part types to reorder", "Distinct SKUs with at least one site line at/below reorder point", "-", "Lines = SKU x site rows"),
    ("Warranties expiring", "Non-retired assets with 0 <= days to expiry <= Warranty_Window_Days", "-", ""),
    ("Forecast utilisation %", "Forecast year-end", "Approved budget", ""),
    ("Spent to date %", "Actual (invoiced) expenditure", "Approved budget", ""),
]
for j, h in enumerate(["KPI", "Numerator", "Denominator", "Note"], start=1):
    c = ws.cell(row=k0 + 1, column=[1, 2, 4, 6][j - 1], value=h)
    c.font, c.fill = F_HDR, FILL_CALC_HDR
for i, (a, b, c_, d) in enumerate(kpi_defs, start=k0 + 2):
    for col, v in zip([1, 2, 4, 6], [a, b, c_, d]):
        cell = ws.cell(row=i, column=col, value=v)
        cell.font = F_BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3)
    ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=5)
l0 = k0 + len(kpi_defs) + 4
ws.cell(row=l0, column=1, value="Allowed values (drop-down lists)").font = F_H2
lists = [("Site", SITES), ("Asset category", CATEGORIES), ("Asset status", STATUSES), ("Assigned team", TEAMS),
         ("RMA component type", COMPONENTS), ("RMA vendor status", VENDOR_STATUS), ("RMA case status", CASE_STATUS),
         ("Procurement category", PROC_CATS), ("Approval stage", STAGES), ("Invoice status", INVOICE_STATUS),
         ("Audit observed site", SITES + ["Not Located"]), ("Audit action status", ACTION_STATUS)]
for i, (lbl, vals) in enumerate(lists, start=l0 + 1):
    ws.cell(row=i, column=1, value=lbl).font = F_BOLD
    c = ws.cell(row=i, column=2, value=", ".join(vals))
    c.font = F_BASE
    ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=6)
ws.freeze_panes = "A5"

# =====================================================================================
# README contents (written last so it can describe everything)
# =====================================================================================
ws = ws_readme
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 16
ws.column_dimensions["D"].width = 95
ws["B1"] = "Regional Data Centre IT Asset Management and Financial Dashboard: Simulated Case Study"
ws["B1"].font = F_TITLE
ws["B2"] = ("Portfolio project by Syed Hamza Shah. All data are synthetic and fictional (generated by "
            "scripts/generate_data.py, seed 42). Not developed for, and contains no information from, any real company.")
ws["B2"].font = Font(name=FONT, size=10, italic=True, color="C00000")
ws["B4"] = "Control settings (yellow cells are editable; every formula reads them by name)"
ws["B4"].font = F_H2
settings = [
    ("RefDate", date(2026, 11, 12), FMT_DATE, "Reporting reference date. All ages, overdue checks and warranty windows use it."),
    ("FY_Start", date(2026, 1, 1), FMT_DATE, "Start of the budget year used for actual expenditure."),
    ("FY_End", date(2026, 12, 31), FMT_DATE, "End of the budget year."),
    ("RMA_Target_Days", 30, "0", "Target resolution period for an RMA case (days). Open cases older than this are Overdue."),
    ("Warranty_Window_Days", 90, "0", "Warranties expiring within this many days are flagged 'Expiring'."),
    ("Audit_Interval_Days", 180, "0", "An asset not physically audited within this many days needs inspection."),
    ("Refresh_Age_Years", 5, "0", "Age at which an asset becomes a refresh (replacement) candidate."),
    ("Days_Per_Month", 30, "0", "Converts supplier lead time in days to months for the reorder point."),
    ("Pending_Approval_Rate", 0.6, FMT_PCT, "Share of pending purchase requests expected to be approved and invoiced this FY (assumption)."),
    ("Tagging_Threshold_EUR", 1000, FMT_EUR, "Unit cost at/above which an item is tagged and tracked in Asset_Register; cheaper parts are bulk stock in Spare_Stock."),
]
for j, h in enumerate(["Setting (workbook name)", "Value", "Meaning"], start=2):
    c = ws.cell(row=5, column=j, value=h)
    c.font, c.fill, c.border = F_HDR, FILL_IN_HDR, BOX
for i, (nm, val, fmt, desc) in enumerate(settings, start=6):
    ws.cell(row=i, column=2, value=nm).font = F_BOLD
    c = ws.cell(row=i, column=3, value=val)
    c.font, c.fill, c.number_format = F_INPUT, FILL_SETTING, fmt
    ws.cell(row=i, column=4, value=desc).font = F_BASE
    for j in (2, 3, 4):
        ws.cell(row=i, column=j).border = BOX
    name(nm, f"README!$C${i}")

ex = META["exceptions"]
SHOW_SITE = assets.loc[assets["Asset ID"] == META["showcase_gpu_server"], "Location"].iloc[0]
text = [
    ("What this workbook demonstrates", "h"),
    ("Serialized IT asset inventory and lifecycle tracking, physical audits, bulk spare-parts replenishment, procurement "
     "tracking, RMA (return merchandise authorization) workflow, FY budget vs forecast, and an executive dashboard. "
     "Every metric is a live Excel formula driven by the records and the settings above.", "p"),
    ("Sheet map", "h"),
    ("Asset_Register: 200 tagged assets across 4 fictional sites (HEL-01 Helsinki, AMS-01 Amsterdam, FRA-01 Frankfurt, "
     "STO-01 Stockholm), with warranty, lifecycle, data-quality and inspection flags.", "p"),
    ("Spare_Stock: 14 bulk part types x 4 sites = 56 lines, with reorder point, inventory position and suggested order quantity.", "p"),
    ("RMA_Tracker: 36 RMA cases linked to Asset IDs, with ageing, SLA flag and KPI tables.", "p"),
    ("Procurement: 55 FY2026 requests; each is classified as Actual, Commitment, Pipeline or Excluded.", "p"),
    ("Budget_Forecast: approved budget vs actual + commitments + additional expected, by site and category, with "
     "regional summary, monthly reporting and a CAPEX/OPEX note.", "p"),
    ("Asset_Audits: 55 audit records (one per sampled asset) with discrepancy analysis and corrective actions.", "p"),
    ("Management_Dashboard: 10 KPI cards, risk watchlist and 4 native charts.", "p"),
    ("Pivot_Analysis: exact steps to build 3 native PivotTables plus formula cross-check grids (not PivotTables).", "p"),
    ("Data_Dictionary: fields, formulas, KPI numerators/denominators and allowed values.", "p"),
    ("Colour conventions", "h"),
    ("Column headers: navy = input data typed by a user; teal-green = calculated by formula; olive = looked up from "
     "another sheet. Cell text: blue = input, black = formula, green = cross-sheet lookup. Fills: red = problem, "
     "amber = warning, green = OK. Yellow cells on this sheet are settings.", "p"),
    ("How to update records", "h"),
    ("1. Add a record by typing in the first empty row directly under a table; Excel extends the table automatically. "
     "Check that the calculated columns filled down; if not, copy the formula from the row above.", "p"),
    ("2. Use the drop-down lists for status, site and category fields so formulas recognise the values.", "p"),
    ("3. Asset IDs typed in RMA_Tracker or Asset_Audits must exist in Asset_Register (see each sheet's Data Check column).", "p"),
    ("4. When an RMA closes, set Case Status = Closed AND enter the Closure Date. When a request is invoiced, set "
     "Approval Stage = Invoiced, Invoice Status = Invoiced and the Invoice Date.", "p"),
    ("5. To report as of another day, change RefDate. Formula ranges reach row 1000 on each sheet.", "p"),
    ("How to refresh reports", "h"),
    ("Formulas, the dashboard and the charts recalculate automatically (if not, press F9, or Ctrl+Alt+F9 for a full "
     "recalculation). If the file opens in Protected View, click 'Enable Editing' first. PivotTables do NOT refresh "
     "automatically: use Data > Refresh All after editing records.", "p"),
    ("Worked example: failed GPU replaced from a spare, then returned under RMA", "h"),
    (f"1. GPU server {META['showcase_gpu_server']} ({SHOW_SITE}) logs uncorrectable ECC errors and a GPU falls off the bus. "
     f"The node is drained from the customer pool.", "p"),
    (f"2. The technician takes tagged spare GPU module {META['showcase_spare_gpu']} from the {SHOW_SITE} spare cage and installs "
     f"it. In Asset_Register that module changes from Available to Deployed and its Rack Position becomes the server's "
     f"rack position. The server returns to service, so it stays Deployed.", "p"),
    (f"3. Case {META['showcase_rma']} is opened in RMA_Tracker against the server, Component Type = GPU, with the failed "
     f"part's serial and the Spare Swap Asset ID. Warranty Status is looked up automatically.", "p"),
    ("4. The vendor approves advance replacement and ships a new GPU (tracking number recorded, Case Status = Awaiting "
     "Replacement). The failed GPU is returned to the vendor.", "p"),
    ("5. When the replacement arrives it is tagged as a new GPU Module asset (Available, spare cage) to restore spare "
     "cover; the case is set to Closed with a Closure Date, which stops the Days Open clock.", "p"),
    ("6. GPU modules are tagged assets (above the tagging threshold), so this workflow never touches Spare_Stock. "
     "Low-cost parts such as fans or DIMMs are issued from Spare_Stock instead (reduce On Hand).", "p"),
    ("Key assumptions", "h"),
    ("No double counting: tagged items (unit cost >= Tagging_Threshold_EUR) are only in Asset_Register; bulk spares "
     "below the threshold are only in Spare_Stock. A Spare_Stock check flags any line whose unit cost reaches the threshold.", "p"),
    ("Spare Stock On Order is the quantity on spare-parts POs in stage 'PO Issued'. When goods are Received, the store "
     "updates On Hand manually after the goods-receipt count (On Hand is not formula-linked).", "p"),
    ("Budget: invoiced = actual; PO Issued or Received but not invoiced = commitment; each request sits in one bucket only. "
     "Commitments and pipeline are assumed to be invoiced within FY2026.", "p"),
    ("Audit accuracy refers only to the audited sample (55 of the active assets), not to the whole fleet.", "p"),
    ("Prices are illustrative planning figures, not market quotations. Site codes, suppliers and serial numbers are invented.", "p"),
    ("Deliberate exceptions planted for analysis", "h"),
    (f"Duplicate serial number: {ex['duplicate_serial'][0]} and {ex['duplicate_serial'][1]}. Deployed switch with no "
     f"rack position: {ex['no_rack']}. Missing PO reference: {ex['no_po']}. Never audited: {ex['no_audit_date']}. "
     f"Retired with no disposal date: {ex['retired_no_disposal']}. Example of a late PO: {META['late_po_request']}. Plus overdue RMAs, "
     f"low-stock lines, audit discrepancies (including a spare GPU that could not be located) and budget overrun risks.", "p"),
]
r = 18
for t_, kind in text:
    c = ws.cell(row=r, column=2, value=t_)
    if kind == "h":
        c.font = F_H2
    else:
        c.font = F_BASE
        c.alignment = WRAP
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.row_dimensions[r].height = max(15, 15 * (len(t_) // 130 + 1))
    r += 1

# Workbook-wide polish
for sh in wb.worksheets:
    sh.sheet_properties.tabColor = {"README": "1F3864", "Management_Dashboard": "C00000",
                                    "Pivot_Analysis": "7F7F7F", "Data_Dictionary": "7F7F7F"}.get(sh.title, "2E6B5E")
wb.active = wb.sheetnames.index("README")
OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print(f"Saved {OUT}")
