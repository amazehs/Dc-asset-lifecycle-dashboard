# Beginner's Guide to the Formulas and PivotTables

This guide explains the workbook's most important formulas in plain language, with worked examples taken from the simulated data. Work through it with the workbook open: click the cell named in each example and compare the formula bar with the explanation.

## 1. Building blocks used everywhere

### Excel Tables
Each data sheet holds an Excel Table (`tblAssets`, `tblSpares`, `tblRMA`, `tblProcurement`, `tblAudits`, `tblBudget`). A table grows automatically when you type in the row directly below it, keeps banded formatting, and is the recommended source for PivotTables, because a refresh picks up new rows.

### Named settings
The yellow cells on the README sheet have names: `RefDate`, `RMA_Target_Days`, `Warranty_Window_Days`, `Audit_Interval_Days`, `Refresh_Age_Years`, `Days_Per_Month`, `Pending_Approval_Rate`, `Tagging_Threshold_EUR`, `FY_Start`, `FY_End`. Formulas use the name instead of a typed number, so `=K5-RefDate` reads as "warranty expiry minus the reporting date". Change the setting once and every formula follows. This meets the requirement that no fixed date is hidden inside a formula.

### Named columns
Frequently used columns also have names, for example `AR_Status` = `Asset_Register!$L$5:$L$1000` and `RMA_Status` = `RMA_Tracker!$M$5:$M$1000`. See them under Formulas > Name Manager. The ranges run to row 1000 so that new records are counted without editing formulas.

### Colour code
Navy headers mark input columns, teal-green headers mark calculated columns, and olive headers mark lookups from another sheet. Blue text is typed data, black text is a formula, green text is a cross-sheet link.

## 2. Counting and summing with conditions

**COUNTIF / COUNTIFS** count rows that meet one or more conditions.

`=COUNTIF(AR_Status,"Deployed")` counts the assets whose status is Deployed (dashboard card 2).

`=COUNTIFS(AR_Location,"HEL-01",AR_Status,"<>Retired")` counts HEL-01 assets that are not retired. `"<>"` means "not equal to"; `"<>"` on its own means "not blank".

A common trap: because the ranges run to row 1000, `COUNTIFS(RMA_Status,"<>Closed")` would also count the empty rows. The workbook therefore adds `RMA_ID,"<>"` ("ID is not blank"): `=COUNTIFS(RMA_ID,"<>",RMA_Status,"<>Closed")`.

**SUMIFS** adds up a column for the rows that meet conditions. In Budget_Forecast row 6:

```
=SUMIFS(PR_Invoiced, PR_Site,A6, PR_Category,B6, PR_InvoiceDate,">="&FY_Start, PR_InvoiceDate,"<="&FY_End)
```

reads as: "add the invoiced amounts where the site equals A6, the category equals B6, and the invoice date falls inside the financial year". The `&` joins the operator text to the date value.

## 3. Looking up data from another sheet: INDEX/MATCH

RMA_Tracker does not retype the asset's category, manufacturer or site; it looks them up from the register:

```
=IFERROR(INDEX(AR_Category, MATCH($B5, AR_ID, 0)), "Not found")
```

* `MATCH($B5, AR_ID, 0)` finds the row position of the Asset ID in the register (0 = exact match).
* `INDEX(AR_Category, position)` returns the category in that position.
* `IFERROR(..., "Not found")` shows a readable message instead of `#N/A` when the ID does not exist.

**XLOOKUP equivalent.** In Microsoft 365 the same lookup can be written `=XLOOKUP($B5, AR_ID, AR_Category, "Not found")`. The workbook uses INDEX/MATCH because it works in every Excel version and in the LibreOffice engine used to verify the file. Being able to explain both is a good interview point.

## 4. Decision logic: IF, AND, OR

The warranty flag in Asset_Register (column R) is a chain of IFs evaluated left to right:

```
=IF(L5="Retired","N/A", IF(K5="","Missing", IF(K5<RefDate,"Expired",
   IF(K5-RefDate<=Warranty_Window_Days,"Expiring","In Warranty"))))
```

Retired assets are excluded first, then blanks, then expired, then those inside the 90-day window. Order matters: test the most specific case first.

## 5. Date arithmetic
Excel stores dates as day numbers, so subtracting two dates gives days.

* Days to warranty expiry: `=K5-RefDate` (negative means already expired).
* Asset age in years: `=ROUND((RefDate-H5)/365.25,1)`.
* RMA days open: `=IF(P5="", RefDate-J5, P5-J5)`. Open cases count up to the reporting date; closed cases stop at the closure date.
* Month end in the monthly report: `=EOMONTH(A39,0)` returns the last day of the month in A39.

## 6. Spare-parts replenishment: worked example

Line SP-NVME-3T8 (3.84TB NVMe SSD) at HEL-01, Spare_Stock row 37:

| Field | Value | Formula |
|---|---|---|
| On hand | 2 | input |
| Reserved | 1 | input |
| On order | 0 | `=SUMIFS(PR_Qty,PR_SKU,A37,PR_Site,B37,PR_Stage,"PO Issued")` |
| Avg monthly consumption | 3.5 | input |
| Lead time | 35 days = 1.17 months | `=I37/Days_Per_Month` |
| Safety stock | 3 | input |
| Reorder point | 3.5 x 1.17 + 3 = 7.08 | `=H37*J37+K37` |
| Inventory position | 2 + 0 - 1 = 1 | `=E37+G37-F37` |
| Status | Critical (1 <= safety stock 3) | `=IF(N37<=K37,"Critical",IF(N37<=L37,"Reorder","OK"))` |
| Target stock | 15 | input (policy: about reorder point + 2 months' use) |
| Suggested order | 15 - 1 = 14 | `=IF(Q37="YES",MAX(0,ROUNDUP(M37-N37,0)),0)` |
| Cost | 14 x €620 = €8,680 | `=R37*S37` |

Why inventory position rather than on-hand stock? Units already on order will arrive, and reserved units are already promised to work, so the position is the stock you can actually count on. `MAX(0, ...)` guarantees the suggestion is never negative, and `ROUNDUP` orders whole units.

**No double counting.** Items at or above the €1,000 tagging threshold (GPU modules, system boards, high-capacity PSUs, 800G optics) are tracked one by one in Asset_Register. Cheaper consumables live only in Spare_Stock. The Data Check column flags any Spare_Stock line whose unit cost reaches the threshold.

## 7. Budget and forecast: worked example

HEL-01 Equipment Purchase (Budget_Forecast row 6):

| Item | EUR | Source |
|---|---|---|
| Approved budget | 2,530,000 | input |
| Actual (invoiced) | 1,307,500 | SUMIFS of Invoiced Amount |
| Outstanding commitments | 1,270,120 | SUMIFS of Committed Amount (PO Issued or Received, not invoiced) |
| Approved pipeline, no PO | 28,800 | SUMIFS of Pipeline Amount, stage "Approved - No PO" |
| Pending pipeline x 60% | 0 | SUMIFS x `Pending_Approval_Rate` |
| Other planned spend | 0 | input estimate |
| Additional expected | 28,800 | `=G6+H6+I6` |
| **Forecast year-end** | **2,606,420** | `=E6+F6+J6` |
| **Projected variance** | **-76,420** | `=D6-K6` (negative = overrun risk) |

**Why nothing is double counted.** In Procurement, column Q gives each request exactly one Budget Treatment:

```
=IF(O5="Invoiced","Actual", IF(OR(J5="PO Issued",J5="Received"),"Commitment", IF(J5="Cancelled","Excluded","Pipeline")))
```

Columns R, S and T then put the request value into only one of Invoiced, Committed or Pipeline. When an invoice arrives the request moves from Commitment to Actual automatically. The reconciliation block on the Procurement sheet proves the buckets add back to the total of all requests (difference = 0).

**CAPEX vs OPEX.** Equipment is typically CAPEX (capitalised and depreciated), and maintenance services are typically OPEX (expensed in the period). Spare parts are labelled "Policy-dependent" because, depending on accounting policy, they may be capitalised or held as inventory and expensed when used. Finance makes that call.

## 8. Audit accuracy and its denominator

`Audit accuracy = matching audit records / all audit records` = 46 / 55 = 83.6%. A record matches only when both site and status agree with the register snapshot. The workbook shows coverage next to it (55 of 191 active assets = 28.8%), so nobody mistakes a sample result for a fleet-wide figure.

`Distinct assets audited` uses `=SUMPRODUCT((C5:C1000<>"")/COUNTIF(C5:C1000,C5:C1000&""))`, a standard trick: each asset contributes 1 divided by its number of occurrences, so duplicates add up to 1.

## 9. Conditional formatting
Rules are formula-based, for example `=R5="Expiring"` gives amber and `=T5>0` (missing fields) gives red. To view or edit them: Home > Conditional Formatting > Manage Rules, and choose "This Worksheet".

## 10. Data validation
Status, site, category and stage columns use drop-down lists (Data > Data Validation > List). They stop typos such as "Deployd" that would silently break COUNTIFS results.

## 11. Building the three PivotTables (step by step)

openpyxl, the Python library that builds the file, cannot create native PivotTables. You build them yourself in about five minutes, which is also good interview practice. The grids on the left of Pivot_Analysis are formula cross-checks, not PivotTables; use them to confirm your PivotTables are correct.

**PivotTable A (assets by location and category)**
1. Go to Asset_Register and click any cell inside the table.
2. Insert > PivotTable > From Table/Range. The range shows `tblAssets`.
3. Choose Existing Worksheet, Location: `Pivot_Analysis!$L$6`, then OK.
4. Drag **Location** to Rows, **Category** to Columns, **Asset ID** to Values. Because Asset ID is text, Excel shows "Count of Asset ID".
5. The grand total should be 200, the same as cross-check grid A.

**PivotTable B (invoiced spend by site and procurement category)**
1. Click inside the Procurement table, then Insert > PivotTable > From Table/Range (`tblProcurement`).
2. Location: `Pivot_Analysis!$L$20`.
3. Rows = **Requesting Site**, Columns = **Procurement Category**, Values = **Invoiced Amount (EUR)** (Sum).
4. Right-click a value > Value Field Settings > Number Format > Currency, € symbol, 0 decimals.
5. The grand total should equal cross-check grid B and the Actual figure on Budget_Forecast.

**PivotTable C (RMA cases by component and status)**
1. Click inside the RMA table, then Insert > PivotTable (`tblRMA`), Location `Pivot_Analysis!$L$34`.
2. Rows = **Component Type**, Columns = **Case Status**, Values = **RMA ID** (Count).
3. The grand total should be 36.

**Refreshing.** PivotTables do not recalculate like formulas. After editing data use Data > Refresh All (Ctrl+Alt+F5), or right-click the PivotTable > Refresh. To change a layout, drag fields between areas in the PivotTable Fields pane. To filter, add a field such as Status to Filters (e.g. exclude Retired).

## 12. Interview talking points

* **"Walk me through the workbook."** Start at the dashboard, then follow one number back to its source, for example open RMAs: card, then RMA_Tracker KPI cell, then the COUNTIFS, then the case rows.
* **"How do you avoid double counting?"** The tagging threshold separates serialized assets from bulk spares. Each procurement request sits in exactly one budget bucket, and a reconciliation cell proves it.
* **"How do you know the numbers are right?"** Every sheet has a Data Check column, the dashboard watchlist counts failures, and an independent Python script recomputes 81 figures and compares them with the workbook (`docs/Validation_Results.md`).
* **"What would you change in a real deployment?"** Feed the register from the DCIM/CMDB system instead of manual entry, take actuals from the ERP's invoice data, phase the budget to the purchase plan rather than straight-line, and agree the spare-parts capitalisation policy with Finance.
* **"What does the 83.6% audit accuracy mean?"** It applies to the 55 audited assets only (29% of the active fleet). It indicates process weaknesses in transfers and status updates; it does not mean 16% of all assets are wrong.
* **Link to your experience.** As a Unit Quartermaster you maintained equipment inventories, raised procurement requests and managed spare parts. This workbook applies the same controls (inventory accuracy, reorder discipline, accountability for every item) to data centre hardware. Keep simulated findings separate from real experience when you describe them.
