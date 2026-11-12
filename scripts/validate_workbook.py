"""
validate_workbook.py
--------------------
Independently recomputes the key numbers with pandas (from the CSV data) and compares them with the
values Excel/LibreOffice calculated in the workbook.

The workbook must contain calculated values, i.e. it has been opened and saved in Excel
(or recalculated with LibreOffice) after build_workbook.py created it:

    python scripts/validate_workbook.py output/DC_Asset_Lifecycle_Financial_Dashboard.xlsx

Writes docs/Validation_Results.md and exits with code 1 if any check fails.
"""

import json
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
path = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "output" / "DC_Asset_Lifecycle_Financial_Dashboard.xlsx"
wb = load_workbook(path, data_only=True)
meta = json.loads((DATA / "meta.json").read_text())
REF = pd.Timestamp(meta["reference_date"])
RATE = meta["pending_approval_rate"]

D = lambda f, c=(): pd.read_csv(DATA / f, parse_dates=list(c))
A = D("asset_register.csv", ["Purchase Date", "Warranty Expiration", "Last Audit Date", "Disposal Date"])
S = D("spare_stock.csv")
R = D("rma_cases.csv", ["Date Reported", "Case Created Date", "Closure Date"])
P = D("procurement.csv", ["Request Date", "Expected Delivery Date", "Actual Delivery Date", "Invoice Date"])
U = D("asset_audits.csv", ["Audit Date"])
B = D("budget_inputs.csv")

results = []


def check(area, desc, expected, actual, tol=0.01):
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        ok = abs(float(expected) - float(actual)) <= tol
    else:
        ok = expected == actual
    results.append((area, desc, expected, actual, "PASS" if ok else "FAIL"))


def col_values(sheet, col, first=5):
    ws = wb[sheet]
    out = []
    r = first
    while ws[f"A{r}"].value not in (None, ""):
        out.append(ws[f"{col}{r}"].value)
        r += 1
    return out


def v(sheet, cell):
    return wb[sheet][cell].value


def labelled(sheet, label_col, label, value_col):
    """Value in value_col on the row whose label_col cell equals label (avoids hard-coded rows)."""
    ws = wb[sheet]
    for c in ws[label_col]:
        if c.value == label:
            return ws[f"{value_col}{c.row}"].value
    raise KeyError(label)


def grid_total(start_label, total_col):
    ws = wb["Pivot_Analysis"]
    start = next(c.row for c in ws["A"] if c.value == start_label)
    row = next(c.row for c in ws["A"] if c.row > start and c.value == "Total")
    return ws[f"{total_col}{row}"].value


# ---------------- independent calculations ----------------
active = A[A.Status != "Retired"]
days_to_exp = (active["Warranty Expiration"] - REF).dt.days
dup_serial_rows = A["Serial Number"].duplicated(keep=False).sum()

R["open"] = R["Case Status"] != "Closed"
R["days"] = (R["Closure Date"].fillna(REF) - R["Case Created Date"]).dt.days
P["value"] = P["Quantity"] * P["Est. Unit Price (EUR)"]
P["treat"] = "Pipeline"
P.loc[P["Approval Stage"].isin(["PO Issued", "Received"]), "treat"] = "Commitment"
P.loc[P["Approval Stage"] == "Cancelled", "treat"] = "Excluded"
P.loc[P["Invoice Status"] == "Invoiced", "treat"] = "Actual"

onorder = P[(P["Approval Stage"] == "PO Issued") & P.SKU.notna()].groupby(["SKU", "Requesting Site"]).Quantity.sum()
S["on_order"] = [onorder.get((k, s), 0) for k, s in zip(S.SKU, S.Site)]
S["pos"] = S["Current Stock (On Hand)"] + S["on_order"] - S["Reserved Stock"]
S["rop"] = S["Avg Monthly Consumption"] * S["Supplier Lead Time (Days)"] / meta["days_per_month"] + S["Safety Stock"]
S["flag"] = S.pos <= S.rop
S["qty"] = [max(0, -(-(t - p) // 1)) if f else 0 for t, p, f in zip(S["Target Stock"], S.pos, S.flag)]
S["cost"] = S.qty * S["Unit Cost (EUR)"]

U["match"] = (U["Recorded Site"] == U["Observed Site"]) & (U["Recorded Status"] == U["Observed Status"])

# ---------------- integrity ----------------
check("Integrity", "Asset IDs unique", True, A["Asset ID"].is_unique)
check("Integrity", "RMA IDs unique", True, R["RMA ID"].is_unique)
check("Integrity", "Request IDs unique", True, P["Request ID"].is_unique)
check("Integrity", "Audit IDs unique", True, U["Audit ID"].is_unique)
check("Integrity", "Every RMA Asset ID exists in register", True, R["Asset ID"].isin(A["Asset ID"]).all())
check("Integrity", "Every spare-swap Asset ID exists", True, R["Spare Swap Asset ID"].dropna().isin(A["Asset ID"]).all())
check("Integrity", "Every audit Asset ID exists", True, U["Asset ID"].isin(A["Asset ID"]).all())
check("Integrity", "No open RMA has a closure date", 0, int((R.open & R["Closure Date"].notna()).sum()))
check("Integrity", "Every closed RMA has a closure date", 0, int((~R.open & R["Closure Date"].isna()).sum()))
check("Integrity", "No RMA closed after the reference date", 0, int((R["Closure Date"] > REF).sum()))
check("Integrity", "No RMA reported before asset purchase", 0,
      int((R.merge(A, on="Asset ID").eval("`Date Reported` < `Purchase Date`")).sum()))
check("Integrity", "Every 'Under RMA' asset has an open case", True,
      set(A.loc[A.Status == "Under RMA", "Asset ID"]) <= set(R.loc[R.open, "Asset ID"]))
check("Integrity", "No audit dated after the reference date", 0, int((U["Audit Date"] > REF).sum()))
check("Integrity", "No invoice dated after the reference date", 0, int((P["Invoice Date"] > REF).sum()))
check("Integrity", "Negative inventory positions", 0, int((S.pos < 0).sum()))
check("Integrity", "Reserved stock exceeding on-hand", 0, int((S["Reserved Stock"] > S["Current Stock (On Hand)"]).sum()))
check("Integrity", "Spare SKUs at/above tagging threshold (double-count risk)", 0, int((S["Unit Cost (EUR)"] >= 1000).sum()))
check("Integrity", "Excel formula errors in workbook", 0,
      sum(1 for ws in wb for row in ws.iter_rows() for c in row
          if isinstance(c.value, str) and c.value.startswith("#") and c.value.endswith(("!", "?", "A"))))

# ---------------- asset register ----------------
check("Assets", "Total assets registered", len(A), v("Management_Dashboard", "B6"))
check("Assets", "Deployed assets", int((A.Status == "Deployed").sum()), v("Management_Dashboard", "C6"))
check("Assets", "Operational utilisation (deployed / active)", (A.Status == "Deployed").sum() / len(active),
      v("Management_Dashboard", "D6"), 1e-9)
check("Assets", "Warranties expiring within 90 days (active)", int(((days_to_exp >= 0) & (days_to_exp <= 90)).sum()),
      v("Management_Dashboard", "C11"))
check("Assets", "Duplicate-serial records flagged", int(dup_serial_rows), col_values("Asset_Register", "U").count("Duplicate"))
miss = (A[["Asset ID", "Serial Number", "Category", "Manufacturer", "Model", "Location", "Purchase Date",
           "Purchase Order ID", "Acquisition Cost (EUR)", "Warranty Expiration", "Status", "Assigned Team",
           "Last Audit Date"]].isna().sum(axis=1)
        + ((A.Status == "Deployed") & A["Rack Position"].isna()) + ((A.Status == "Retired") & A["Disposal Date"].isna()))
check("Assets", "Assets with missing mandatory data", int((miss > 0).sum()), v("Management_Dashboard", "I10"))

# ---------------- RMA ----------------
closed = R[~R.open]
check("RMA", "Total cases", len(R), v("RMA_Tracker", "X5"))
check("RMA", "Open cases", int(R.open.sum()), v("RMA_Tracker", "X6"))
check("RMA", "Overdue open cases (> 30 days)", int((R.open & (R.days > 30)).sum()), v("RMA_Tracker", "X7"))
check("RMA", "% closed", len(closed) / len(R), v("RMA_Tracker", "X9"), 1e-9)
check("RMA", "Average resolution days (closed)", closed.days.mean(), v("RMA_Tracker", "X10"), 1e-6)
check("RMA", "Dashboard open RMA card", int(R.open.sum()), v("Management_Dashboard", "F6"))

# ---------------- spares ----------------
check("Spares", "On-order quantity linked from Procurement", int(S.on_order.sum()), sum(col_values("Spare_Stock", "G")))
check("Spares", "Lines at/below reorder point", int(S.flag.sum()), v("Spare_Stock", "X6"))
check("Spares", "Critical lines (position <= safety stock)", int((S.pos <= S["Safety Stock"]).sum()), v("Spare_Stock", "X7"))
check("Spares", "Distinct SKUs needing replenishment", int(S.loc[S.flag, "SKU"].nunique()), v("Management_Dashboard", "B11"))
check("Spares", "Suggested order quantity", int(S.qty.sum()), v("Spare_Stock", "X9"))
check("Spares", "Estimated replenishment cost (EUR)", float(S.cost.sum()), v("Spare_Stock", "X10"))
check("Spares", "No negative suggested quantities", 0, sum(1 for q in col_values("Spare_Stock", "R") if q < 0))

# ---------------- procurement & budget ----------------
tot = P.groupby("treat").value.sum()
check("Procurement", "Actual (invoiced) total", float(tot.get("Actual", 0)), v("Procurement", "Z15"))
check("Procurement", "Commitments total", float(tot.get("Commitment", 0)), v("Procurement", "Z16"))
check("Procurement", "Pipeline total", float(tot.get("Pipeline", 0)), v("Procurement", "Z17"))
check("Procurement", "Buckets reconcile to total of all requests", 0.0, v("Procurement", "Z20"))
check("Procurement", "No request is both actual and commitment", 0,
      int(((pd.Series(col_values("Procurement", "R")) > 0) & (pd.Series(col_values("Procurement", "S")) > 0)).sum()))
check("Procurement", "Late POs", int(((P["Approval Stage"] == "PO Issued") & (P["Expected Delivery Date"] < REF)).sum()),
      v("Procurement", "Z22"))

bws = wb["Budget_Forecast"]
for i, row in B.iterrows():
    r = 6 + i
    sub = P[(P["Requesting Site"] == row.Site) & (P["Procurement Category"] == row["Budget Category"])]
    act = sub.loc[sub.treat == "Actual", "value"].sum()
    com = sub.loc[sub.treat == "Commitment", "value"].sum()
    add = (sub.loc[sub["Approval Stage"] == "Approved - No PO", "value"].sum()
           + sub.loc[sub["Approval Stage"] == "Pending Approval", "value"].sum() * RATE + row["Other Planned Spend (EUR)"])
    fc = act + com + add
    lbl = f"{row.Site} / {row['Budget Category']}"
    check("Budget", f"Forecast {lbl}", fc, bws[f"K{r}"].value)
    check("Budget", f"Variance {lbl}", row["Approved Annual Budget (EUR)"] - fc, bws[f"L{r}"].value)
check("Budget", "Approved budget total (dashboard)", float(B["Approved Annual Budget (EUR)"].sum()), v("Management_Dashboard", "D11"))
check("Budget", "Forecast total (dashboard)", float(labelled("Budget_Forecast", "A", "Region total", "K")), wb["Management_Dashboard"]["E11"].value)
check("Budget", "Variance total (dashboard)", float(labelled("Budget_Forecast", "A", "Region total", "L")), wb["Management_Dashboard"]["F11"].value)
check("Budget", "Site summary reconciles to budget table", 0.0, round(float(labelled("Budget_Forecast", "A", "Check: total equals budget table (should be 0)", "G")), 2))
check("Budget", "Monthly actuals reconcile to budget-table actuals", 0.0,
      round(float(labelled("Budget_Forecast", "A", "Check: monthly actuals equal budget-table actuals (should be 0)", "H")), 2))

# ---------------- audits ----------------
check("Audits", "Audit records", len(U), v("Asset_Audits", "T5"))
check("Audits", "Audit accuracy (matches / records)", U.match.mean(), v("Asset_Audits", "T7"), 1e-9)
check("Audits", "Discrepancies", int((~U.match).sum()), v("Asset_Audits", "T8"))
check("Audits", "Unresolved discrepancies", int(((~U.match) & (U["Corrective Action Status"] != "Resolved")).sum()),
      v("Asset_Audits", "T9"))
check("Audits", "Distinct assets audited", int(U["Asset ID"].nunique()), v("Asset_Audits", "T11"))
check("Audits", "Resolved discrepancies now match register", "all Yes",
      "all Yes" if all(x == "Yes" for x, s, res in zip(col_values("Asset_Audits", "N"), col_values("Asset_Audits", "M"),
                                                         col_values("Asset_Audits", "J")) if s == "Resolved") else "mismatch")

# ---------------- pivot cross-checks ----------------
pv = wb["Pivot_Analysis"]
check("Pivot grids", "Grid A grand total = asset count", len(A),
      grid_total("A. Asset count by location and category (all statuses)", "J"))
check("Pivot grids", "Grid B grand total = invoiced total", float(tot.get("Actual", 0)),
      grid_total("B. Invoiced expenditure by site and procurement category (EUR)", "E"))
check("Pivot grids", "Grid C grand total = RMA count", len(R),
      grid_total("C. RMA cases by component type and case status", "E"))

# ---------------- report ----------------
fails = [r for r in results if r[4] == "FAIL"]
lines = ["# Validation results", "",
         f"Workbook: `{path.name}`  |  Reference date: {meta['reference_date']}  |  "
         f"Checks: {len(results)}  |  Passed: {len(results) - len(fails)}  |  Failed: {len(fails)}", "",
         "Expected values are recomputed independently with pandas from the CSV source data; "
         "actual values are read from the calculated workbook.", "",
         "| Area | Check | Expected (pandas) | Workbook | Result |", "|---|---|---|---|---|"]


def fmt(x):
    if isinstance(x, float):
        return f"{x:,.4f}" if abs(x) < 10 else f"{x:,.2f}"
    return str(x)


for a, d, e, w, res in results:
    lines.append(f"| {a} | {d} | {fmt(e)} | {fmt(w)} | {res} |")
(ROOT / "docs").mkdir(exist_ok=True)
(ROOT / "docs" / "Validation_Results.md").write_text("\n".join(lines) + "\n")
print("\n".join(lines[-len(results):]) if fails else f"All {len(results)} checks passed.")
for f in fails:
    print("FAIL:", f)
sys.exit(1 if fails else 0)
