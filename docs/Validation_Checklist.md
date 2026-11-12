# Validation Checklist

Two layers of checking were used: an automated comparison against independent pandas calculations, and manual spot checks you can repeat in Excel. Automated results are in [Validation_Results.md](Validation_Results.md) (81 checks, all passed at build time; the workbook was recalculated with LibreOffice and showed 0 formula errors across about 4,600 formulas).

## A. Automated checks (scripts/validate_workbook.py)

| # | Area | What is confirmed | Status |
|---|---|---|---|
| 1 | Integrity | Asset, RMA, request and audit IDs are unique | Pass |
| 2 | Integrity | Every Asset ID referenced in RMA_Tracker (including spare-swap IDs) and Asset_Audits exists in Asset_Register | Pass |
| 3 | Integrity | No open RMA has a closure date; every closed RMA has one; no closure after the reporting date | Pass |
| 4 | Integrity | No RMA reported before the asset was purchased; every "Under RMA" asset has an open case | Pass |
| 5 | Integrity | No audit or invoice dated after the reporting date | Pass |
| 6 | Integrity | No negative inventory position; reserved never exceeds on hand; no spare SKU at/above the tagging threshold | Pass |
| 7 | Integrity | Zero Excel error values (#N/A, #REF!, #DIV/0!, #NAME?, #VALUE!) | Pass |
| 8 | Assets | Total, deployed, utilisation, expiring warranties, duplicate serials, missing-data count match pandas | Pass |
| 9 | RMA | Total, open, overdue, % closed, average resolution time match pandas | Pass |
| 10 | Spares | On-order quantities equal PO Issued quantities; flagged lines, critical lines, order quantity and cost match | Pass |
| 11 | Procurement | Actual, commitment and pipeline totals match; buckets reconcile to total (difference 0); no request is both actual and committed | Pass |
| 12 | Budget | Forecast and variance for all 12 site/category lines match; dashboard totals equal the budget table; site summary and monthly actuals reconcile (difference 0) | Pass |
| 13 | Audits | Accuracy, discrepancies, unresolved count, distinct assets match; every resolved discrepancy now agrees with the register | Pass |
| 14 | Pivot grids | Cross-check grid totals equal asset count, invoiced total and RMA count | Pass |

To re-run: open the workbook in Excel, let it calculate, save it, then run
`python scripts/validate_workbook.py output/DC_Asset_Lifecycle_Financial_Dashboard.xlsx`.

## B. Manual spot checks in Excel (about 15 minutes)

Tick each one after you have done it yourself; this is also good interview preparation.

- [ ] **Recalculation.** Open the file, click Enable Editing, and confirm the dashboard shows 200 assets, 129 deployed and 16 open RMAs.
- [ ] **Reorder point.** Spare_Stock row 37 (SP-NVME-3T8, HEL-01): calculate 3.5 x (35/30) + 3 = 7.08 by hand and compare with column L.
- [ ] **Inventory position.** Same row: 2 on hand + 0 on order - 1 reserved = 1, so status is Critical and the suggested order is 15 - 1 = 14.
- [ ] **Forecast.** Budget_Forecast row 6: 1,307,500 + 1,270,120 + 28,800 = 2,606,420; 2,530,000 - 2,606,420 = -76,420.
- [ ] **Reconciliation.** Procurement cell Z20 ("Difference vs total of all requests") shows 0.
- [ ] **RMA ageing.** Pick an open case and calculate 12-Nov-2026 minus the Case Created Date. It should equal Days Open.
- [ ] **Audit accuracy.** Filter Asset_Audits column J to "Match": 46 of 55 records, which is 83.6%.
- [ ] **Lookup.** In RMA_Tracker type a non-existent Asset ID (e.g. AST-9999) in a copy of the file. Category shows "Not found" and Data Check shows "Invalid Asset ID".
- [ ] **Reference date.** Change README!C6 (RefDate) to 31-Dec-2026. Overdue RMAs and warranty flags should increase. Change it back afterwards.
- [ ] **PivotTables.** Build the three PivotTables (Pivot_Analysis sheet). Their grand totals should be 200, the invoiced total and 36.
- [ ] **Data validation.** Try typing "Deployd" in an Asset_Register Status cell; Excel should reject it.

## B2. Change test

Four realistic edits (faulty GPU added, RMA opened, spare stock issued, PO invoiced) were applied to a copy and the dashboard was rechecked: 34 of 34 figures moved exactly as predicted, and forecast and variance stayed unchanged when the PO moved from commitment to actual. See [Change_Test_Results.md](Change_Test_Results.md).

## C. Known deliberate exceptions (expected, not errors)

These were planted so the analysis has something to find. They appear on the dashboard watchlist and in the Data Check / Record Check columns:

- Duplicate serial number on two optical transceivers
- One deployed switch with no rack position, one asset with no PO reference, one asset never audited, one retired asset with no disposal date
- Overdue RMAs, low and critical spare-stock lines, late POs, unresolved audit discrepancies (including a spare GPU module not located) and three budget lines at overrun risk

The exact Asset IDs are listed at the bottom of the README sheet.

## D. Known limitations

- Native PivotTables are not included in the generated file; Python's openpyxl cannot create them. Build steps are provided, and the cross-check grids are clearly labelled as formulas.
- The workbook was verified with LibreOffice's calculation engine, not with Microsoft Excel itself. All functions used (COUNTIFS, SUMIFS, INDEX/MATCH, IFERROR, SUMPRODUCT, EOMONTH, ROUNDUP) behave the same in Excel. Opening it in Excel and running check B1 confirms this on your machine.
- When you add a row to a table, Excel normally fills the calculated columns down. If it does not, copy the formula from the row above.
