"""
verify_change_test.py
---------------------
Compares the calculated baseline workbook with the calculated change-test copy and checks every
dashboard figure against the change we expect from the four edits in change_test.py.

Both files must contain calculated values (open each in Excel and save, or recalculate with LibreOffice):

    python scripts/verify_change_test.py BASELINE.xlsx CHANGETEST.xlsx

Prints a before/after table; exit code 1 if any figure moves differently from expectation.
"""

import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
base_path, test_path = Path(sys.argv[1]), Path(sys.argv[2])
B = load_workbook(base_path, data_only=True)
T = load_workbook(test_path, data_only=True)

# Hand-calculated expected effect of each edit (see docs/Change_Test_Results.md for the reasoning).
# SP-OPT-400G HEL-01: position 11 <= reorder point 15.8 -> order 30 - 11 = 19 units x EUR 720 = EUR 13,680
INVOICE = 1_168_000          # REQ-26-034: 4 x EUR 292,000 moves from commitment to actual
# (label, sheet, cell, expected change, kind)  kind "delta" = after - before; "abs" = expected after value
CHECKS = [
    ("Dashboard: total assets registered", "Management_Dashboard", "B6", 1, "delta"),
    ("Dashboard: deployed assets", "Management_Dashboard", "C6", 0, "delta"),
    ("Dashboard: operational utilisation", "Management_Dashboard", "D6", 129 / 192, "abs"),
    ("Dashboard: audit accuracy", "Management_Dashboard", "E6", 0, "delta"),
    ("Dashboard: open RMA cases", "Management_Dashboard", "F6", 1, "delta"),
    ("Dashboard: spare part types to reorder", "Management_Dashboard", "B11", 1, "delta"),
    ("Dashboard: warranties expiring <= 90 days", "Management_Dashboard", "C11", 0, "delta"),
    ("Dashboard: approved budget", "Management_Dashboard", "D11", 0, "delta"),
    ("Dashboard: forecast year-end", "Management_Dashboard", "E11", 0, "delta"),
    ("Dashboard: projected variance", "Management_Dashboard", "F11", 0, "delta"),
    ("Watchlist: overdue open RMAs", "Management_Dashboard", "I5", 0, "delta"),
    ("Watchlist: critical spare lines", "Management_Dashboard", "I6", 0, "delta"),
    ("Watchlist: late purchase orders", "Management_Dashboard", "I9", -1, "delta"),
    ("Watchlist: assets under RMA without open case", "Management_Dashboard", "I13", 0, "abs"),
    ("Watchlist: record-level data checks failing", "Management_Dashboard", "I14", 0, "delta"),
    ("Chart data: HEL-01 active assets", "Management_Dashboard", "C51", 1, "delta"),
    ("Chart data: Under RMA assets", "Management_Dashboard", "C61", 1, "delta"),
    ("Chart data: open RMAs aged 0-15 days", "Management_Dashboard", "C66", 1, "delta"),
    ("Chart data: AMS-01 forecast year-end", "Management_Dashboard", "D74", 0, "delta"),
    ("RMA_Tracker: total cases", "RMA_Tracker", "X5", 1, "delta"),
    ("RMA_Tracker: % closed (20 / 37)", "RMA_Tracker", "X9", 20 / 37, "abs"),
    ("RMA_Tracker: average resolution of closed cases", "RMA_Tracker", "X10", 0, "delta"),
    ("Spare_Stock: lines at/below reorder point", "Spare_Stock", "X6", 1, "delta"),
    ("Spare_Stock: suggested order quantity", "Spare_Stock", "X9", 19, "delta"),
    ("Spare_Stock: estimated replenishment cost", "Spare_Stock", "X10", 19 * 720, "delta"),
    ("Procurement: actual (invoiced)", "Procurement", "Z15", INVOICE, "delta"),
    ("Procurement: outstanding commitments", "Procurement", "Z16", -INVOICE, "delta"),
    ("Procurement: buckets reconcile (difference)", "Procurement", "Z20", 0, "abs"),
    ("Budget_Forecast: AMS-01 equipment actual", "Budget_Forecast", "E9", INVOICE, "delta"),
    ("Budget_Forecast: AMS-01 equipment commitments", "Budget_Forecast", "F9", -INVOICE, "delta"),
    ("Budget_Forecast: AMS-01 equipment forecast", "Budget_Forecast", "K9", 0, "delta"),
    ("Budget_Forecast: region actual", "Budget_Forecast", "E18", INVOICE, "delta"),
    ("Budget_Forecast: November actual (monthly report)", "Budget_Forecast", "H49", INVOICE, "delta"),
    ("Budget_Forecast: monthly vs table check", "Budget_Forecast", "H52", 0, "abs"),
]

rows, fails = [], 0
for label, sheet, cell, exp, kind in CHECKS:
    b, t = B[sheet][cell].value, T[sheet][cell].value
    got = (t - b) if kind == "delta" else t
    ok = abs((got or 0) - exp) < 1e-6
    fails += not ok
    rows.append((label, f"{sheet}!{cell}", b, t, exp, kind, "PASS" if ok else "FAIL"))


def f(x):
    if isinstance(x, float) and abs(x) < 1.5:
        return f"{x:.2%}"
    return f"{x:,.0f}" if isinstance(x, (int, float)) else str(x)


md = ["| Figure | Cell | Before | After | Expected | Result |", "|---|---|---|---|---|---|"]
for label, ref, b, t, exp, kind, res in rows:
    exp_txt = (("+" if exp > 0 else "") + f(exp) + " change" if exp else "no change") if kind == "delta" else f(exp)
    md.append(f"| {label} | {ref} | {f(b)} | {f(t)} | {exp_txt} | {res} |")
print("\n".join(md))
print(f"\n{len(rows) - fails}/{len(rows)} figures moved as expected")
sys.exit(1 if fails else 0)
