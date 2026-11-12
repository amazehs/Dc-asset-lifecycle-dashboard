# Change Test: does the dashboard follow the source records?

**Simulated data.** This test edits a copy of the workbook the way a user would, recalculates it, and checks that every affected figure moves by the amount worked out by hand beforehand. The original workbook is not changed.

## The four edits

1. **Faulty GPU added.** New asset AST-0201 (NVIDIA H100 SXM5 FRU module, HEL-01, €24,800) was entered at the bottom of the Asset_Register table with Status = Under RMA and Rack Position = RMA-STAGING.
2. **RMA opened.** New case RMA-26-037 for AST-0201: Component Type = GPU, failure "Double-bit ECC errors during burn-in", reported 10 Nov, created 11 Nov 2026, Vendor Status = Submitted, Case Status = Open. Category, manufacturer, site and warranty status were filled by the lookup formulas, not typed.
3. **Spare stock reduced.** 8 units of SP-OPT-400G (400G transceiver) were issued at HEL-01, so On Hand went from 19 to 11.
4. **Purchase order changed.** REQ-26-034 (PO-26-00010, 4 GPU servers for AMS-01, €1,168,000) was delivered on 6 Nov and invoiced on 11 Nov 2026: Approval Stage and Invoice Status were changed to Invoiced and the dates entered.

## What should happen (worked out before recalculating)

- Assets: total 200 to 201 and active 191 to 192. Deployed stays at 129, so utilisation falls to 129 / 192 = 67.19%. The Under RMA count rises from 13 to 14. HEL-01 active assets rise from 59 to 60.
- RMA: open cases 16 to 17 and total 36 to 37. Overdue stays at 6 because the new case is 1 day old; it lands in the 0-15 day bucket. % closed becomes 20 / 37 = 54.05%. Average resolution time does not change, because it uses closed cases only. The new faulty asset has an open case, so the "under RMA without an open case" watchlist line stays at 0.
- Spares: position = 11 + 0 - 0 = 11, which is at or below the reorder point of 15.8 (6.0/month x 42/30 months + safety 6) but above safety stock 6. The line becomes "Reorder" (not Critical): order 30 - 11 = 19 units x €720 = €13,680. SP-OPT-400G was not flagged at any site before, so part types to reorder go from 7 to 8.
- Procurement and budget: €1,168,000 moves from Commitment to Actual. Actual and commitments change by +/- €1,168,000, but the forecast (actual + commitments + additional) and the variance must not change. This confirms that nothing is double counted. The November column of the monthly report gains €1,168,000. Late POs fall from 4 to 3 because this PO is no longer outstanding.

## Results

| Figure | Cell | Before | After | Expected | Result |
|---|---|---|---|---|---|
| Dashboard: total assets registered | Management_Dashboard!B6 | 200 | 201 | +1 change | PASS |
| Dashboard: deployed assets | Management_Dashboard!C6 | 129 | 129 | no change | PASS |
| Dashboard: operational utilisation | Management_Dashboard!D6 | 67.54% | 67.19% | 67.19% | PASS |
| Dashboard: audit accuracy | Management_Dashboard!E6 | 83.64% | 83.64% | no change | PASS |
| Dashboard: open RMA cases | Management_Dashboard!F6 | 16 | 17 | +1 change | PASS |
| Dashboard: spare part types to reorder | Management_Dashboard!B11 | 7 | 8 | +1 change | PASS |
| Dashboard: warranties expiring <= 90 days | Management_Dashboard!C11 | 19 | 19 | no change | PASS |
| Dashboard: approved budget | Management_Dashboard!D11 | 8,130,000 | 8,130,000 | no change | PASS |
| Dashboard: forecast year-end | Management_Dashboard!E11 | 7,438,270 | 7,438,270 | no change | PASS |
| Dashboard: projected variance | Management_Dashboard!F11 | 691,730 | 691,730 | no change | PASS |
| Watchlist: overdue open RMAs | Management_Dashboard!I5 | 6 | 6 | no change | PASS |
| Watchlist: critical spare lines | Management_Dashboard!I6 | 3 | 3 | no change | PASS |
| Watchlist: late purchase orders | Management_Dashboard!I9 | 4 | 3 | -1 change | PASS |
| Watchlist: assets under RMA without open case | Management_Dashboard!I13 | 0 | 0 | 0 | PASS |
| Watchlist: record-level data checks failing | Management_Dashboard!I14 | 6 | 6 | no change | PASS |
| Chart data: HEL-01 active assets | Management_Dashboard!C51 | 59 | 60 | +1 change | PASS |
| Chart data: Under RMA assets | Management_Dashboard!C61 | 13 | 14 | +1 change | PASS |
| Chart data: open RMAs aged 0-15 days | Management_Dashboard!C66 | 6 | 7 | +1 change | PASS |
| Chart data: AMS-01 forecast year-end | Management_Dashboard!D74 | 2,522,390 | 2,522,390 | no change | PASS |
| RMA_Tracker: total cases | RMA_Tracker!X5 | 36 | 37 | +1 change | PASS |
| RMA_Tracker: % closed (20 / 37) | RMA_Tracker!X9 | 55.56% | 54.05% | 54.05% | PASS |
| RMA_Tracker: average resolution of closed cases | RMA_Tracker!X10 | 24 | 24 | no change | PASS |
| Spare_Stock: lines at/below reorder point | Spare_Stock!X6 | 10 | 11 | +1 change | PASS |
| Spare_Stock: suggested order quantity | Spare_Stock!X9 | 89 | 108 | +19 change | PASS |
| Spare_Stock: estimated replenishment cost | Spare_Stock!X10 | 34,125 | 47,805 | +13,680 change | PASS |
| Procurement: actual (invoiced) | Procurement!Z15 | 3,893,375 | 5,061,375 | +1,168,000 change | PASS |
| Procurement: outstanding commitments | Procurement!Z16 | 2,602,195 | 1,434,195 | -1,168,000 change | PASS |
| Procurement: buckets reconcile (difference) | Procurement!Z20 | 0 | 0 | 0 | PASS |
| Budget_Forecast: AMS-01 equipment actual | Budget_Forecast!E9 | 1,206,400 | 2,374,400 | +1,168,000 change | PASS |
| Budget_Forecast: AMS-01 equipment commitments | Budget_Forecast!F9 | 1,168,000 | 0 | -1,168,000 change | PASS |
| Budget_Forecast: AMS-01 equipment forecast | Budget_Forecast!K9 | 2,385,920 | 2,385,920 | no change | PASS |
| Budget_Forecast: region actual | Budget_Forecast!E18 | 3,893,375 | 5,061,375 | +1,168,000 change | PASS |
| Budget_Forecast: November actual (monthly report) | Budget_Forecast!H49 | 0 | 1,168,000 | +1,168,000 change | PASS |
| Budget_Forecast: monthly vs table check | Budget_Forecast!H52 | 0 | 0 | 0 | PASS |

**34 of 34 figures moved exactly as expected.** The recalculated copy had 0 formula errors (4,643 formulas). All four dashboard charts remained in place and updated: the HEL-01 bar, the Under RMA bar and the 0-15 day RMA bar each rose by one.

## How to repeat this yourself in Excel

1. Make a copy of the workbook and open it.
2. Type the four edits above directly into the sheets (start typing in the first empty row under each table, so Excel extends the table and fills the calculated columns).
3. Press F9 if the workbook is set to manual calculation, then open Management_Dashboard and compare it with the "After" column.
4. If you have built the PivotTables, they will still show the old numbers until you click Data > Refresh All. This is expected PivotTable behaviour, not an error.

Automated version:

```bash
python scripts/change_test.py            # writes output/..._ChangeTest.xlsx
# open both workbooks in Excel and save them (stores calculated values), then:
python scripts/verify_change_test.py output/DC_Asset_Lifecycle_Financial_Dashboard.xlsx output/DC_Asset_Lifecycle_Financial_Dashboard_ChangeTest.xlsx
```

Verification here used LibreOffice's calculation engine; Excel uses the same functions, so it should give the same results.
