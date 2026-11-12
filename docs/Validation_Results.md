# Validation results

Workbook: `DC_Asset_Lifecycle_Financial_Dashboard.xlsx` (recalculated copy)  |  Reference date: 2026-11-12  |  Checks: 81  |  Passed: 81  |  Failed: 0

Expected values are recomputed independently with pandas from the CSV source data; actual values are read from the calculated workbook.

| Area | Check | Expected (pandas) | Workbook | Result |
|---|---|---|---|---|
| Integrity | Asset IDs unique | True | True | PASS |
| Integrity | RMA IDs unique | True | True | PASS |
| Integrity | Request IDs unique | True | True | PASS |
| Integrity | Audit IDs unique | True | True | PASS |
| Integrity | Every RMA Asset ID exists in register | True | True | PASS |
| Integrity | Every spare-swap Asset ID exists | True | True | PASS |
| Integrity | Every audit Asset ID exists | True | True | PASS |
| Integrity | No open RMA has a closure date | 0 | 0 | PASS |
| Integrity | Every closed RMA has a closure date | 0 | 0 | PASS |
| Integrity | No RMA closed after the reference date | 0 | 0 | PASS |
| Integrity | No RMA reported before asset purchase | 0 | 0 | PASS |
| Integrity | Every 'Under RMA' asset has an open case | True | True | PASS |
| Integrity | No audit dated after the reference date | 0 | 0 | PASS |
| Integrity | No invoice dated after the reference date | 0 | 0 | PASS |
| Integrity | Negative inventory positions | 0 | 0 | PASS |
| Integrity | Reserved stock exceeding on-hand | 0 | 0 | PASS |
| Integrity | Spare SKUs at/above tagging threshold (double-count risk) | 0 | 0 | PASS |
| Integrity | Excel formula errors in workbook | 0 | 0 | PASS |
| Assets | Total assets registered | 200 | 200 | PASS |
| Assets | Deployed assets | 129 | 129 | PASS |
| Assets | Operational utilisation (deployed / active) | 0.6754 | 0.6754 | PASS |
| Assets | Warranties expiring within 90 days (active) | 19 | 19 | PASS |
| Assets | Duplicate-serial records flagged | 2 | 2 | PASS |
| Assets | Assets with missing mandatory data | 4 | 4 | PASS |
| RMA | Total cases | 36 | 36 | PASS |
| RMA | Open cases | 16 | 16 | PASS |
| RMA | Overdue open cases (> 30 days) | 6 | 6 | PASS |
| RMA | % closed | 0.5556 | 0.5556 | PASS |
| RMA | Average resolution days (closed) | 24.15 | 24.15 | PASS |
| RMA | Dashboard open RMA card | 16 | 16 | PASS |
| Spares | On-order quantity linked from Procurement | 120 | 120 | PASS |
| Spares | Lines at/below reorder point | 10 | 10 | PASS |
| Spares | Critical lines (position <= safety stock) | 3 | 3 | PASS |
| Spares | Distinct SKUs needing replenishment | 7 | 7 | PASS |
| Spares | Suggested order quantity | 89 | 89 | PASS |
| Spares | Estimated replenishment cost (EUR) | 34,125.00 | 34125 | PASS |
| Spares | No negative suggested quantities | 0 | 0 | PASS |
| Procurement | Actual (invoiced) total | 3,893,375.00 | 3893375 | PASS |
| Procurement | Commitments total | 2,602,195.00 | 2602195 | PASS |
| Procurement | Pipeline total | 1,240,900.00 | 1240900 | PASS |
| Procurement | Buckets reconcile to total of all requests | 0.0000 | 0 | PASS |
| Procurement | No request is both actual and commitment | 0 | 0 | PASS |
| Procurement | Late POs | 4 | 4 | PASS |
| Budget | Forecast HEL-01 / Equipment Purchase | 2,606,420.00 | 2606420 | PASS |
| Budget | Variance HEL-01 / Equipment Purchase | -76,420.00 | -76420 | PASS |
| Budget | Forecast HEL-01 / Spare Parts Replenishment | 37,300.00 | 37300 | PASS |
| Budget | Variance HEL-01 / Spare Parts Replenishment | 12,700.00 | 12700 | PASS |
| Budget | Forecast HEL-01 / Maintenance Services | 72,700.00 | 72700 | PASS |
| Budget | Variance HEL-01 / Maintenance Services | 7,300.00 | 7300 | PASS |
| Budget | Forecast AMS-01 / Equipment Purchase | 2,385,920.00 | 2385920 | PASS |
| Budget | Variance AMS-01 / Equipment Purchase | 484,080.00 | 484080 | PASS |
| Budget | Forecast AMS-01 / Spare Parts Replenishment | 63,270.00 | 63270 | PASS |
| Budget | Variance AMS-01 / Spare Parts Replenishment | -3,270.00 | -3270 | PASS |
| Budget | Forecast AMS-01 / Maintenance Services | 73,200.00 | 73200 | PASS |
| Budget | Variance AMS-01 / Maintenance Services | 6,800.00 | 6800 | PASS |
| Budget | Forecast FRA-01 / Equipment Purchase | 944,400.00 | 944400 | PASS |
| Budget | Variance FRA-01 / Equipment Purchase | 65,600.00 | 65600 | PASS |
| Budget | Forecast FRA-01 / Spare Parts Replenishment | 39,000.00 | 39000 | PASS |
| Budget | Variance FRA-01 / Spare Parts Replenishment | 11,000.00 | 11000 | PASS |
| Budget | Forecast FRA-01 / Maintenance Services | 73,940.00 | 73940 | PASS |
| Budget | Variance FRA-01 / Maintenance Services | -3,940.00 | -3940 | PASS |
| Budget | Forecast STO-01 / Equipment Purchase | 1,076,000.00 | 1076000 | PASS |
| Budget | Variance STO-01 / Equipment Purchase | 154,000.00 | 154000 | PASS |
| Budget | Forecast STO-01 / Spare Parts Replenishment | 37,680.00 | 37680 | PASS |
| Budget | Variance STO-01 / Spare Parts Replenishment | 12,320.00 | 12320 | PASS |
| Budget | Forecast STO-01 / Maintenance Services | 28,440.00 | 28440 | PASS |
| Budget | Variance STO-01 / Maintenance Services | 21,560.00 | 21560 | PASS |
| Budget | Approved budget total (dashboard) | 8,130,000.00 | 8130000 | PASS |
| Budget | Forecast total (dashboard) | 7,438,270.00 | 7438270 | PASS |
| Budget | Variance total (dashboard) | 691,730.00 | 691730 | PASS |
| Budget | Site summary reconciles to budget table | 0.0000 | 0.0000 | PASS |
| Budget | Monthly actuals reconcile to budget-table actuals | 0.0000 | 0.0000 | PASS |
| Audits | Audit records | 55 | 55 | PASS |
| Audits | Audit accuracy (matches / records) | 0.8364 | 0.8364 | PASS |
| Audits | Discrepancies | 9 | 9 | PASS |
| Audits | Unresolved discrepancies | 5 | 5 | PASS |
| Audits | Distinct assets audited | 55 | 55 | PASS |
| Audits | Resolved discrepancies now match register | all Yes | all Yes | PASS |
| Pivot grids | Grid A grand total = asset count | 200 | 200 | PASS |
| Pivot grids | Grid B grand total = invoiced total | 3,893,375.00 | 3893375 | PASS |
| Pivot grids | Grid C grand total = RMA count | 36 | 36 | PASS |
