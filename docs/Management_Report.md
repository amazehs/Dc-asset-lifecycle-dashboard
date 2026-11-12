# Management Report: Regional IT Asset Position at 12 November 2026

**Simulated case study.** Every figure below comes from the fictional data set in this repository (four invented sites: HEL-01, AMS-01, FRA-01, STO-01). The findings show how the workbook supports decisions; they are not observations from real operations. Figures refresh if the workbook data or the reporting date change.

## Headline position

The region has 200 registered assets (191 active, 9 retired). 129 are deployed, giving 67.5% operational utilisation of the active fleet; for production systems alone (GPU servers, storage servers and switches) utilisation is 87.5%, because tagged spare components (GPU modules, system boards, PSUs) sit in the denominator as idle stock. The FY2026 forecast is €7.44M against an approved €8.13M budget (91.5% utilisation, €0.69M projected underspend), but three budget lines are at overrun risk.

## Findings

**1. Inventory discrepancies.** 55 assets were physically audited (28.8% of the active fleet). 46 matched the register, an 83.6% match rate for that sample only. Of 9 discrepancies, 4 are resolved and 5 remain open, covering assets worth €602,920. The most serious are a spare H200 GPU module (AST-0181, STO-01) that could not be located, a GPU server recorded at HEL-01 but found at FRA-01 (AST-0021), a GPU server recorded as Deployed but found on a workbench (AST-0107), and a spare system board moved from AMS-01 to STO-01 and installed with no change record (AST-0143). The pattern points to inter-site moves and status changes not being booked at the time they happen. The register itself also carries one duplicate serial number (AST-0025 / AST-0049) and four records with missing mandatory data.

**2. Failure and RMA patterns.** There are 36 RMA cases in FY2026; 16 are open and 6 exceed the 30-day target (oldest: a PSU at STO-01 open 75 days, a storage controller at FRA-01 open 62 days). Closed cases took 24.2 days on average. GPUs are the most frequent failed component (13 of 36 cases), followed by PSUs (8) and system boards (5). Relative to installed base, storage servers (8 cases on 22 assets) and GPU servers (14 on 44) have the highest RMA rates. With samples this small these are indicators to monitor, not statistically proven failure rates.

**3. Spare-parts shortages.** 10 of 56 SKU-site lines are at or below their reorder point, across 7 part types; the suggested replenishment is 89 units costing about €34,125. AMS-01 has the most flagged lines (4), followed by HEL-01 (3). Three lines are critical, at or below safety stock: 400G DAC cables at FRA-01 and 2000W PSUs at AMS-01 (both zero stock), and 3.84TB NVMe drives at HEL-01 (one usable unit).

**4. Budget risks and favourable variances.** HEL-01 equipment is forecast at 103.0% of budget (€76,420 over), driven by €1.27M of GPU servers and 800G optics that have been received but not yet invoiced. AMS-01 spare parts (105.5%) and FRA-01 maintenance (105.6%) are small overruns in value. AMS-01 equipment shows the largest favourable variance (€484,080), although €1.17M of its commitment is a GPU-server PO already past its expected delivery date. €973,500 of requests await approval; the forecast includes 60% of that pipeline as an explicit assumption.

**5. Warranty exposure.** 19 active assets lose warranty within 90 days (combined acquisition value about €0.94M), and 46 active assets are already out of warranty.

## Recommendations

1. **Protect availability through spares and RMA turnaround.** Place the 10 flagged replenishment orders this week, critical lines first. Escalate the 6 overdue RMAs with each vendor's account team. Agree a minimum stock of tagged spare GPU modules per site, so that a GPU failure can always be fixed by a spare swap while the RMA runs.
2. **Close the audit gaps and tighten movement control.** Resolve the 5 open discrepancies within 30 days, starting with the missing GPU module (raise a loss report if it is not found within 10 working days). Require a transfer record for every inter-site move, and link status changes to the maintenance ticket. Extend the next audit wave to the 37 active assets flagged "Inspection Required".
3. **Control expenditure before year-end.** Hold or re-phase the remaining HEL-01 equipment pipeline, or request a formal budget transfer from AMS-01's favourable equipment variance. Chase the late AMS-01 GPU-server delivery so its €1.17M commitment converts to delivered capacity. Review the 19 expiring warranties against RMA history, then choose between extended warranty and self-maintenance with spares.

*Prepared from the workbook `DC_Asset_Lifecycle_Financial_Dashboard.xlsx` (reference date 12 Nov 2026). Simulated data; conclusions illustrate the analysis method only.*
