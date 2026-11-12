"""
change_test.py
--------------
Simulates four everyday updates a user would type into the workbook, so you can check that the
dashboard follows the source records. The original workbook is not modified; a copy is written.

    python scripts/change_test.py
    -> output/DC_Asset_Lifecycle_Financial_Dashboard_ChangeTest.xlsx

Changes applied (as a user would make them in Excel):
 1. Add a faulty GPU: new tagged asset AST-0201 (H100 GPU module, HEL-01) with Status = Under RMA.
 2. Open an RMA: new case RMA-26-037 against AST-0201, Case Status = Open.
 3. Reduce spare stock: 8 units of SP-OPT-400G issued at HEL-01 (On Hand 19 -> 11).
 4. Change a purchase order: REQ-26-034 (PO-26-00010, 4 GPU servers for AMS-01) is delivered and
    invoiced, so its stage moves from PO Issued to Invoiced.

New rows get the same formulas as the row above (what Excel does when a table grows), and the
table ranges are extended to include them.
"""

from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "output" / "DC_Asset_Lifecycle_Financial_Dashboard.xlsx"
DST = ROOT / "output" / "DC_Asset_Lifecycle_Financial_Dashboard_ChangeTest.xlsx"

wb = load_workbook(SRC)


def append_row(ws, table_name, values):
    """Write input values into the first empty row under the table, copy formulas and styles from the
    row above, and extend the table range. values: {column letter: value} for input columns."""
    table = ws.tables[table_name]
    first, last = table.ref.split(":")
    last_col = "".join(ch for ch in last if ch.isalpha())
    last_row = int("".join(ch for ch in last if ch.isdigit()))
    new = last_row + 1
    for cell in ws[last_row]:
        target = ws.cell(row=new, column=cell.column)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            target.value = Translator(cell.value, origin=cell.coordinate).translate_formula(target.coordinate)
        target._style = cell._style
    for col, val in values.items():
        ws[f"{col}{new}"] = val
    table.ref = f"{first}:{last_col}{new}"
    return new


def find_row(ws, col, value):
    for c in ws[col]:
        if c.value == value:
            return c.row
    raise KeyError(value)


log = []

# 1. Faulty GPU added to the register
ws = wb["Asset_Register"]
r = append_row(ws, "tblAssets", {
    "A": "AST-0201", "B": "NVH26418832", "C": "GPU Module", "D": "NVIDIA", "E": "H100 SXM5 80GB FRU module",
    "F": "HEL-01", "G": "RMA-STAGING", "H": date(2026, 3, 16), "I": "PO-2026-51207", "J": 24800,
    "K": date(2029, 3, 16), "L": "Under RMA", "M": "GPU Compute Ops", "N": date(2026, 9, 22), "O": None})
log.append(f"Asset_Register row {r}: added AST-0201 (GPU Module, HEL-01, Under RMA)")

# 2. RMA opened for it
ws = wb["RMA_Tracker"]
r = append_row(ws, "tblRMA", {
    "A": "RMA-26-037", "B": "AST-0201", "D": "GPU", "E": "FP55102877",
    "H": "Double-bit ECC errors during burn-in; module fails NVIDIA diagnostics",
    "I": date(2026, 11, 10), "J": date(2026, 11, 11), "L": "Submitted", "M": "Open",
    "N": None, "O": None, "P": None})
log.append(f"RMA_Tracker row {r}: opened RMA-26-037 for AST-0201 (Open, created 11-Nov-2026)")

# 3. Spare stock reduced
ws = wb["Spare_Stock"]
r = next(c.row for c in ws["A"] if c.value == "SP-OPT-400G" and ws[f"B{c.row}"].value == "HEL-01")
before = ws[f"E{r}"].value
ws[f"E{r}"] = before - 8
log.append(f"Spare_Stock row {r}: SP-OPT-400G HEL-01 On Hand {before} -> {before - 8}")

# 4. Purchase order delivered and invoiced
ws = wb["Procurement"]
r = find_row(ws, "A", "REQ-26-034")
ws[f"J{r}"] = "Invoiced"
ws[f"N{r}"] = date(2026, 11, 6)
ws[f"O{r}"] = "Invoiced"
ws[f"P{r}"] = date(2026, 11, 11)
log.append(f"Procurement row {r}: REQ-26-034 PO Issued -> Invoiced (delivered 06-Nov, invoiced 11-Nov-2026)")

wb.save(DST)
print("\n".join(log))
print(f"Saved {DST}")
