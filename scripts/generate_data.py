"""
generate_data.py
----------------
Generates the fictional, reproducible synthetic data set for the
"Regional Data Centre IT Asset Management and Financial Dashboard" case study.

Everything here is invented: sites, serial numbers, suppliers, prices and events.
Prices are illustrative planning figures, NOT verified market quotations.

Run:  python scripts/generate_data.py
Output: CSV files in ./data plus data/meta.json (IDs used in the README example).

The random seed is fixed, so every run produces identical data.
"""

import json
import math
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from dateutil.relativedelta import relativedelta

SEED = 42
REF = date(2026, 11, 12)          # reporting reference date (also stored in the workbook)
FY_START, FY_END = date(2026, 1, 1), date(2026, 12, 31)
PENDING_APPROVAL_RATE = 0.6       # share of pending requests expected to be approved & invoiced in FY
DAYS_PER_MONTH = 30

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
rng = random.Random(SEED)

SITES = ["HEL-01", "AMS-01", "FRA-01", "STO-01"]
SITE_WEIGHTS = [0.32, 0.26, 0.24, 0.18]


def rdate(start, end):
    """Random date between start and end inclusive."""
    return start + timedelta(days=rng.randint(0, (end - start).days))


# ---------------------------------------------------------------------------
# 1. ASSET REGISTER  (serialized assets, unit cost >= EUR 1,000 tagging threshold)
# ---------------------------------------------------------------------------
# category: (list of (manufacturer, model, unit cost EUR, serial prefix), warranty months,
#            status counts {Deployed, Available, Under Maintenance, Under RMA, Retired})
CATALOG = {
    "GPU Server": ([("Supermicro", "SYS-821GE-TNHR, 8x H100 SXM5 80GB", 286000, "SMC"),
                    ("Dell", "PowerEdge XE9680, 8x H100 SXM5 80GB", 294000, "DEL")], 36,
                   {"Deployed": 37, "Available": 1, "Under Maintenance": 2, "Under RMA": 2, "Retired": 2}),
    "GPU Module": ([("NVIDIA", "H100 SXM5 80GB FRU module", 24500, "NVH"),
                    ("NVIDIA", "H200 SXM 141GB FRU module", 31000, "NVJ")], 36,
                   {"Deployed": 9, "Available": 14, "Under Maintenance": 0, "Under RMA": 2, "Retired": 1}),
    "Storage Server": ([("Supermicro", "SSG-121E-NES24R, 24x NVMe", 46000, "SMS"),
                        ("HPE", "ProLiant DL380 Gen11, 12x LFF", 38000, "HPE")], 36,
                       {"Deployed": 18, "Available": 0, "Under Maintenance": 2, "Under RMA": 1, "Retired": 1}),
    "Network Switch": ([("Arista", "7060X6-64PE, 64x 800G", 58000, "ARS"),
                        ("NVIDIA", "SN5600 Spectrum-4, 64x 800G", 62000, "MLX"),
                        ("Arista", "7050SX3-48YC8 (management)", 14500, "ARM")], 36,
                       {"Deployed": 22, "Available": 2, "Under Maintenance": 0, "Under RMA": 1, "Retired": 1}),
    "Optical Transceiver": ([("NVIDIA", "OSFP 800G 2xSR4 multimode", 1350, "MTX"),
                             ("Finisar", "OSFP 800G DR8 single-mode", 1480, "FTL")], 24,
                            {"Deployed": 26, "Available": 6, "Under Maintenance": 0, "Under RMA": 3, "Retired": 1}),
    "Motherboard": ([("Supermicro", "X13DEG-OAD system board (GPU node)", 3900, "SMB"),
                     ("Dell", "XE9680 system board FRU", 4200, "DMB")], 36,
                    {"Deployed": 1, "Available": 10, "Under Maintenance": 0, "Under RMA": 2, "Retired": 1}),
    "Power Supply Unit": ([("Supermicro", "PWS-3K05A 3000W Titanium (GPU node)", 1150, "SPS"),
                           ("Delta", "5500W Titanium CRPS (HGX)", 1290, "DPS")], 36,
                          {"Deployed": 5, "Available": 10, "Under Maintenance": 0, "Under RMA": 2, "Retired": 1}),
    "Other Infrastructure": ([("Raritan", "PX4 intelligent rack PDU, 3-phase", 2800, "RPX"),
                              ("Opengear", "OM2248 out-of-band console server", 3200, "OGC"),
                              ("Dell", "PowerEdge R660 management server", 9500, "DMG")], 36,
                             {"Deployed": 11, "Available": 1, "Under Maintenance": 1, "Under RMA": 0, "Retired": 1}),
}

TEAM = {
    "GPU Server": "GPU Compute Ops", "GPU Module": "GPU Compute Ops", "Storage Server": "Storage Ops",
    "Network Switch": "Network Engineering", "Optical Transceiver": "Network Engineering",
    "Motherboard": "GPU Compute Ops", "Power Supply Unit": "GPU Compute Ops",
    "Other Infrastructure": "DC Facilities",
}

assets = []
used_serials = set()
for cat, (models, wmonths, counts) in CATALOG.items():
    for status, n in counts.items():
        for _ in range(n):
            mfr, model, cost, pfx = rng.choice(models)
            site = rng.choices(SITES, SITE_WEIGHTS)[0]
            if status == "Retired":
                pdate = rdate(date(2019, 6, 1), date(2021, 12, 31))
            elif status == "Under RMA":
                pdate = rdate(date(2024, 3, 1), date(2026, 3, 31))     # must still be in warranty
            else:
                pdate = rdate(date(2023, 1, 1), date(2026, 9, 30))
            while True:
                serial = f"{pfx}{pdate.year % 100:02d}{rng.randint(0, 999999):06d}"
                if serial not in used_serials:
                    used_serials.add(serial)
                    break
            assets.append({
                "Category": cat, "Manufacturer": mfr, "Model": model, "Location": site,
                "Purchase Date": pdate, "Acquisition Cost (EUR)": round(cost * rng.uniform(0.96, 1.04), -1),
                "Warranty Months": wmonths, "Status": status, "Serial Number": serial,
            })

rng.shuffle(assets)
for i, a in enumerate(assets, start=1):
    a["Asset ID"] = f"AST-{i:04d}"

# Force a realistic cluster of warranties expiring inside the 90-day window
non_retired = [a for a in assets if a["Status"] in ("Deployed", "Available")]
for a in rng.sample(non_retired, 5):
    expiry = rdate(REF + timedelta(days=5), REF + timedelta(days=85))
    a["Purchase Date"] = expiry - relativedelta(months=a["Warranty Months"])

# Purchase order IDs: one historical PO per category/site/year
po_map = {}
for a in assets:
    key = (a["Category"], a["Location"], a["Purchase Date"].year)
    if key not in po_map:
        po_map[key] = f"PO-{a['Purchase Date'].year}-{rng.randint(10000, 99999)}"
    a["Purchase Order ID"] = po_map[key]
    a["Warranty Expiration"] = a["Purchase Date"] + relativedelta(months=a["Warranty Months"])
    a["Assigned Team"] = "Hardware Spares Pool" if a["Status"] == "Available" else TEAM[a["Category"]]
    a["Disposal Date"] = None
    if a["Status"] == "Retired":
        a["Disposal Date"] = rdate(date(2025, 3, 1), date(2026, 10, 30))
        a["Assigned Team"] = "Asset Disposal"
    a["Rack Position"] = {
        "Deployed": f"H{rng.randint(1, 3)}-R{rng.randint(1, 24):02d}-U{rng.randint(1, 42):02d}",
        "Available": f"SPARE-CAGE-B{rng.randint(1, 20):02d}",
        "Under Maintenance": f"WORKBENCH-0{rng.randint(1, 3)}",
        "Under RMA": "RMA-STAGING",
        "Retired": None,
    }[a["Status"]]
    if a["Status"] == "Retired":
        a["Last Audit Date"] = a["Disposal Date"] - timedelta(days=rng.randint(10, 120))
    else:
        a["Last Audit Date"] = REF - timedelta(days=rng.randint(15, 250))
        if a["Last Audit Date"] < a["Purchase Date"]:
            a["Last Audit Date"] = a["Purchase Date"] + timedelta(days=rng.randint(1, 10))
            if a["Last Audit Date"] > REF:
                a["Last Audit Date"] = REF

by_id = {a["Asset ID"]: a for a in assets}

# ---------------------------------------------------------------------------
# 2. RMA TRACKER
# ---------------------------------------------------------------------------
FAILURES = {
    "GPU": ["Uncorrectable ECC errors; GPU fell off the bus (Xid 79)",
            "NVLink CRC errors and thermal throttling under load",
            "GPU not enumerated after firmware update"],
    "Motherboard": ["No POST; BMC reports CPU voltage regulator fault",
                    "PCIe link training failure on GPU baseboard slot"],
    "Power Supply Unit": ["PSU output failure; redundancy-lost alarm",
                          "PSU fan failure followed by over-temperature shutdown"],
    "Optical Transceiver": ["High pre-FEC BER and link flapping",
                            "Rx optical power below threshold after fibre cleaning"],
    "Network Switch": ["ASIC parity errors causing port resets"],
    "Storage Controller": ["NVMe backplane/controller fault; drives offline"],
    "Memory DIMM": ["Correctable ECC error threshold exceeded on DIMM"],
    "Fan Module": ["Fan RPM below threshold; chassis fan fault"],
    "NVMe Drive": ["Media errors; SMART critical warning"],
}
COMPONENT_FOR_CATEGORY = {"GPU Module": "GPU", "Optical Transceiver": "Optical Transceiver",
                          "Motherboard": "Motherboard", "Power Supply Unit": "Power Supply Unit",
                          "Network Switch": "Network Switch", "Storage Server": "Storage Controller"}

rma_rows = []


def add_rma(asset, component, created, closed=None, vendor=None, case=None, swap=None, failure=None):
    reported = created - timedelta(days=rng.randint(0, 2))
    rma_rows.append({
        "Asset ID": asset["Asset ID"], "Component Type": component,
        "Failed Part Serial": f"FP{rng.randint(10**7, 10**8 - 1)}",
        "Failure Description": failure or rng.choice(FAILURES[component]),
        "Date Reported": reported, "Case Created Date": created,
        "Vendor Status": vendor, "Case Status": case,
        "Spare Swap Asset ID": swap["Asset ID"] if swap else None,
        "Replacement Tracking No.": None, "Closure Date": closed,
    })


# 2a. Open cases for every asset currently "Under RMA"
open_ages = iter([62, 48, 41, 37, 33, 75, 26, 21, 18, 14, 11, 9, 6, 4, 3, 22])
for a in [a for a in assets if a["Status"] == "Under RMA"]:
    comp = COMPONENT_FOR_CATEGORY.get(a["Category"], "Motherboard")
    if a["Category"] == "GPU Server":
        comp = rng.choice(["GPU", "Motherboard"])
    age = next(open_ages)
    vendor = rng.choice(["Submitted", "Approved - Advance Replacement", "Replacement Shipped"])
    case = "Awaiting Replacement" if vendor == "Replacement Shipped" else "Open"
    add_rma(a, comp, REF - timedelta(days=age), vendor=vendor, case=case)

# 2b. GPU failures in deployed GPU servers, fixed by swapping in a spare GPU module.
#     Each deployed GPU module was installed as one of these swaps (9 cases: 3 open, 6 closed).
gpu_servers = [a for a in assets if a["Category"] == "GPU Server" and a["Status"] == "Deployed"]
swap_modules = [a for a in assets if a["Category"] == "GPU Module" and a["Status"] == "Deployed"]
showcase = None
for k, module in enumerate(swap_modules):
    candidates = [s for s in gpu_servers if s["Location"] == module["Location"]
                  and s["Warranty Expiration"] > REF and s["Purchase Date"] < date(2025, 12, 1)
                  and not any(r["Asset ID"] == s["Asset ID"] for r in rma_rows)]
    server = rng.choice(candidates)
    module["Rack Position"] = server["Rack Position"]       # module now sits in that server
    if k < 3:   # open
        age = next(open_ages)
        vendor = "Replacement Shipped" if k == 0 else rng.choice(["Submitted", "Approved - Advance Replacement"])
        case = "Awaiting Replacement" if vendor == "Replacement Shipped" else "Open"
        fail = "Uncorrectable ECC errors; GPU fell off the bus (Xid 79)" if k == 0 else None
        add_rma(server, "GPU", REF - timedelta(days=age), vendor=vendor, case=case, swap=module, failure=fail)
        module["Purchase Date"] = min(module["Purchase Date"], REF - timedelta(days=age + 60))
        if k == 0:
            showcase = (server["Asset ID"], module["Asset ID"])
    else:
        created = rdate(date(2026, 2, 1), date(2026, 9, 15))
        add_rma(server, "GPU", created, closed=created + timedelta(days=rng.randint(9, 40)),
                vendor="Replacement Received", case="Closed", swap=module)
        module["Purchase Date"] = min(module["Purchase Date"], created - timedelta(days=60))
    module["Warranty Expiration"] = module["Purchase Date"] + relativedelta(months=module["Warranty Months"])
    if module["Last Audit Date"] < module["Purchase Date"]:
        module["Last Audit Date"] = module["Purchase Date"] + timedelta(days=5)

# 2c. Two out-of-warranty claims rejected by the vendor; the units were then retired
for cat, comp in [("Motherboard", "Motherboard"), ("Power Supply Unit", "Power Supply Unit")]:
    a = next(x for x in assets if x["Category"] == cat and x["Status"] == "Retired")
    created = rdate(date(2026, 3, 1), date(2026, 6, 30))
    closed = created + timedelta(days=rng.randint(5, 12))
    add_rma(a, comp, created, closed=closed, vendor="Rejected - Out of Warranty", case="Closed")
    a["Disposal Date"] = closed + timedelta(days=rng.randint(7, 30))
    a["Last Audit Date"] = created - timedelta(days=rng.randint(20, 90))

# 2d. Other closed cases on in-warranty deployed equipment (advance replacement or repair)
pool = [a for a in assets if a["Status"] == "Deployed" and a["Category"] in
        ("GPU Server", "Storage Server", "Network Switch") and a["Warranty Expiration"] > date(2026, 10, 1)
        and a["Purchase Date"] < date(2025, 12, 1)
        and not any(r["Asset ID"] == a["Asset ID"] for r in rma_rows)]
comp_choice = {"GPU Server": ["Power Supply Unit", "Memory DIMM", "Fan Module", "Motherboard", "GPU"],
               "Storage Server": ["NVMe Drive", "Storage Controller", "Power Supply Unit"],
               "Network Switch": ["Optical Transceiver", "Power Supply Unit", "Fan Module"]}
for a in rng.sample(pool, 12):
    created = rdate(date(2026, 1, 10), date(2026, 10, 1))
    closed = min(created + timedelta(days=rng.randint(6, 48)), REF - timedelta(days=1))
    vendor = rng.choice(["Replacement Received", "Replacement Received", "Repaired"])
    add_rma(a, rng.choice(comp_choice[a["Category"]]), created, closed=closed, vendor=vendor, case="Closed")

rma_rows.sort(key=lambda r: r["Case Created Date"])
for i, r in enumerate(rma_rows, start=1):
    r["RMA ID"] = f"RMA-26-{i:03d}"
    if r["Vendor Status"] in ("Replacement Shipped", "Replacement Received"):
        r["Replacement Tracking No."] = f"1Z{rng.randint(10**9, 10**10 - 1)}EU"
showcase_rma = next(r["RMA ID"] for r in rma_rows if r["Spare Swap Asset ID"] == showcase[1])

# ---------------------------------------------------------------------------
# 3. ASSET AUDITS  (one audit per sampled asset; recorded values are a snapshot)
# ---------------------------------------------------------------------------
eligible = [a for a in assets if a["Status"] != "Retired"]
audit_assets = rng.sample(eligible, 55)
# Discrepancy plan: (type, resolved?)
plan = [("Location", True)] * 3 + [("Location", False)] * 2 + [("Status", True), ("Status", False)] \
       + [("Not Located", False), ("Location & Status", False)]
audit_rows = []
gpu_spares = [a for a in audit_assets if a["Category"] == "GPU Module" and a["Status"] == "Available"]
if not gpu_spares:   # guarantee the high-value "not located" example is a spare GPU
    swap_in = next(a for a in eligible if a["Category"] == "GPU Module" and a["Status"] == "Available")
    audit_assets[0] = swap_in
    gpu_spares = [swap_in]
not_located_asset = gpu_spares[0]
others = [a for a in audit_assets if a is not not_located_asset]
disc_assets = {}
deployed_others = [a for a in others if a["Status"] == "Deployed"]
avail_others = [a for a in others if a["Status"] == "Available" and a is not not_located_asset]
for typ, resolved in plan:
    if typ == "Not Located":
        disc_assets[not_located_asset["Asset ID"]] = (typ, resolved)
    elif typ == "Status":
        a = deployed_others.pop() if not resolved else avail_others.pop()
        disc_assets[a["Asset ID"]] = (typ, resolved)
    elif typ == "Location & Status":
        a = avail_others.pop()
        disc_assets[a["Asset ID"]] = (typ, resolved)
    else:
        a = deployed_others.pop()
        disc_assets[a["Asset ID"]] = (typ, resolved)

ACTIONS = {
    ("Location", True): "Register location corrected after physical verification and transfer record found",
    ("Location", False): "Confirm inter-site transfer paperwork; update register or return asset to recorded site",
    ("Status", True): "Status corrected: asset confirmed in spare cage, removed from production inventory",
    ("Status", False): "Asset found on workbench; confirm maintenance ticket and update status",
    ("Not Located", False): "Escalate: search spare cages, RMA shipments and transfer logs; raise loss report if not found in 10 days",
    ("Location & Status", False): "Spare moved and installed without change record; raise change ticket and update register",
}
for a in audit_assets:
    adate = rdate(date(2026, 8, 3), date(2026, 11, 10))
    rec_site, rec_status = a["Location"], a["Status"]
    obs_site, obs_status = a["Location"], a["Status"]
    action, astatus = "None required", "Not Required"
    if a["Asset ID"] in disc_assets:
        typ, resolved = disc_assets[a["Asset ID"]]
        other_site = rng.choice([s for s in SITES if s != a["Location"]])
        if typ == "Location" and resolved:
            rec_site = other_site                      # register was wrong, now corrected
        elif typ == "Location":
            obs_site = other_site                      # register still shows old site
        elif typ == "Status" and resolved:
            rec_status = "Deployed"                    # register said deployed, now corrected to Available
        elif typ == "Status":
            obs_status = "Under Maintenance"
        elif typ == "Not Located":
            obs_site, obs_status = "Not Located", "Not Located"
        elif typ == "Location & Status":
            obs_site, obs_status = other_site, "Deployed"
        action = ACTIONS[(typ, resolved)]
        astatus = "Resolved" if resolved else rng.choice(["Open", "In Progress"])
        if typ == "Not Located":
            astatus = "Open"
    a["Last Audit Date"] = max(adate, a["Purchase Date"] + timedelta(days=1))
    audit_rows.append({"Audit Date": a["Last Audit Date"], "Asset ID": a["Asset ID"],
                       "Recorded Site": rec_site, "Observed Site": obs_site,
                       "Recorded Status": rec_status, "Observed Status": obs_status,
                       "Auditor": rng.choice(["Site Ops Lead", "Asset Controller", "Regional Asset Manager"]),
                       "Corrective Action": action, "Corrective Action Status": astatus})
audit_rows.sort(key=lambda r: (r["Audit Date"], r["Asset ID"]))
for i, r in enumerate(audit_rows, start=1):
    r["Audit ID"] = f"AUD-26-{i:03d}"

# ---------------------------------------------------------------------------
# 4. Deliberate data-quality exceptions in the register (flagged by formulas)
# ---------------------------------------------------------------------------
audited_ids = {r["Asset ID"] for r in audit_rows}
rma_ids = {r["Asset ID"] for r in rma_rows} | {r["Spare Swap Asset ID"] for r in rma_rows if r["Spare Swap Asset ID"]}
free = [a for a in assets if a["Asset ID"] not in audited_ids | rma_ids]
optics = [a for a in free if a["Category"] == "Optical Transceiver" and a["Status"] == "Deployed"]
dup_a, dup_b = optics[0], optics[1]
dup_b["Serial Number"] = dup_a["Serial Number"]          # duplicate serial (transcription error)
sw = next(a for a in free if a["Category"] == "Network Switch" and a["Status"] == "Deployed")
sw["Rack Position"] = None                               # deployed but no rack position
nopo = next(a for a in free if a["Category"] == "Storage Server" and a["Status"] == "Deployed")
nopo["Purchase Order ID"] = None                         # missing PO reference
noaudit = next(a for a in free if a["Category"] == "Power Supply Unit" and a["Status"] == "Available")
noaudit["Last Audit Date"] = None                        # never audited since receipt
ret = next(a for a in free if a["Status"] == "Retired" and a["Category"] == "Other Infrastructure")
ret["Disposal Date"] = None                              # retired without disposal record
exceptions = {"duplicate_serial": [dup_a["Asset ID"], dup_b["Asset ID"]], "no_rack": sw["Asset ID"],
              "no_po": nopo["Asset ID"], "no_audit_date": noaudit["Asset ID"], "retired_no_disposal": ret["Asset ID"]}

# ---------------------------------------------------------------------------
# 5. SPARE STOCK (bulk, non-serialized parts below the EUR 1,000 tagging threshold)
# ---------------------------------------------------------------------------
SPARES = [  # SKU, category, description, base monthly consumption, lead days, unit cost
    ("SP-OPT-400G", "Optical Transceiver", "QSFP112 400G SR4 transceiver", 6, 42, 720),
    ("SP-OPT-100G", "Optical Transceiver", "QSFP28 100G SR4 transceiver", 4, 28, 160),
    ("SP-DAC-400G", "Cabling", "400G QSFP112 DAC cable, 2 m", 5, 21, 210),
    ("SP-DAC-BRK", "Cabling", "800G to 2x400G breakout DAC, 2 m", 3, 35, 390),
    ("SP-MPO-12", "Cabling", "MPO-12 OM4 fibre trunk, 10 m", 4, 14, 95),
    ("SP-PSU-2K", "Power Supply Unit", "2000W Titanium PSU (storage/switch)", 2, 35, 480),
    ("SP-FAN-GPU", "Cooling", "GPU server hot-swap fan module", 3, 28, 140),
    ("SP-DIMM-64", "Memory", "64GB DDR5-4800 RDIMM", 5, 28, 310),
    ("SP-NVME-3T8", "Storage Media", "3.84TB NVMe U.2 SSD", 3, 35, 620),
    ("SP-NVME-7T6", "Storage Media", "7.68TB NVMe U.2 SSD", 2, 42, 980),
    ("SP-HDD-20T", "Storage Media", "20TB SAS 7.2k HDD", 3, 30, 420),
    ("SP-NIC-OCP", "Network Adapter", "OCP 3.0 2x25GbE NIC", 1, 45, 390),
    ("SP-SSD-BOOT", "Storage Media", "960GB M.2 NVMe boot SSD", 2, 21, 140),
    ("SP-PWR-C19", "Cabling", "C19-C20 power cord, pack of 10", 2, 10, 45),
]
SITE_FACTOR = {"HEL-01": 1.3, "AMS-01": 1.0, "FRA-01": 0.9, "STO-01": 0.6}

spare_lines = []
for sku, cat, desc, base, lead, cost in SPARES:
    for site in SITES:
        cons = max(0.5, round(base * SITE_FACTOR[site] * rng.uniform(0.8, 1.2) * 2) / 2)
        safety = math.ceil(cons * 0.75)
        rop = cons * lead / DAYS_PER_MONTH + safety
        target = math.ceil(rop + 2 * cons)
        spare_lines.append({"SKU": sku, "Site": site, "Equipment Category": cat, "Description": desc,
                            "Avg Monthly Consumption": cons, "Supplier Lead Time (Days)": lead,
                            "Safety Stock": safety, "Target Stock": target, "Unit Cost (EUR)": cost,
                            "_rop": rop})

# ---------------------------------------------------------------------------
# 6. PROCUREMENT REQUESTS (FY2026)
# ---------------------------------------------------------------------------
SUPPLIERS = {
    "Equipment Purchase": ["Nordic Compute Distribution Oy", "EuroRack Systems BV", "Baltic Hardware Supply AB"],
    "Spare Parts Replenishment": ["PhotonLink Optics GmbH", "Baltic Hardware Supply AB", "EuroRack Systems BV"],
    "Maintenance Services": ["DataCare Field Services Ltd", "CoolFlow Facility Maintenance BV", "PowerGrid Services Oy"],
}
EQUIPMENT_ITEMS = [("GPU server, 8x H100 SXM5 (HGX platform)", 292000, (2, 4)),
                   ("Spare GPU FRU modules, H100 SXM5", 25000, (4, 8)),
                   ("800G leaf switch, 64-port", 59000, (2, 4)),
                   ("OSFP 800G transceivers", 1380, (32, 96)),
                   ("NVMe storage server, 24-bay", 46500, (2, 4)),
                   ("Spare GPU-node system boards", 4000, (4, 8)),
                   ("Management server, 1U", 9600, (2, 4))]
MAINT_ITEMS = [("Quarterly PDU/UPS preventive maintenance", 18000), ("Fibre inspection and cleaning service", 9500),
               ("Out-of-warranty GPU server board repair", 14500), ("Third-party storage maintenance (6 months)", 42000),
               ("CRAH/cooling preventive maintenance", 26000), ("Thermal survey and airflow remediation", 12500)]
STAGE_PLAN = {
    "Equipment Purchase": {"Invoiced": 7, "Received": 2, "PO Issued": 2, "Approved - No PO": 2,
                           "Pending Approval": 2, "Cancelled": 1},
    "Spare Parts Replenishment": {"Invoiced": 12, "Received": 3, "PO Issued": 6, "Approved - No PO": 2,
                                  "Pending Approval": 2, "Cancelled": 1},
    "Maintenance Services": {"Invoiced": 7, "Received": 1, "PO Issued": 1, "Approved - No PO": 1,
                             "Pending Approval": 2, "Cancelled": 1},
}
STAGE_DATES = {"Invoiced": (date(2026, 1, 8), date(2026, 8, 31)), "Received": (date(2026, 7, 1), date(2026, 10, 10)),
               "PO Issued": (date(2026, 10, 1), date(2026, 11, 8)), "Approved - No PO": (date(2026, 10, 15), date(2026, 11, 10)),
               "Pending Approval": (date(2026, 10, 20), date(2026, 11, 11)), "Cancelled": (date(2026, 2, 1), date(2026, 9, 30))}

SITE_CYCLE = ["HEL-01", "AMS-01", "FRA-01", "STO-01", "HEL-01", "AMS-01", "FRA-01", "HEL-01"]
proc = []
po_counter = 1
cycle_pos = {c: 0 for c in STAGE_PLAN}
for pcat, stages in STAGE_PLAN.items():
    for stage, n in stages.items():
        for _ in range(n):
            if pcat == "Spare Parts Replenishment":
                site = rng.choices(SITES, SITE_WEIGHTS)[0]
            else:   # equipment and maintenance rotate through sites so every site has activity
                site = SITE_CYCLE[cycle_pos[pcat] % len(SITE_CYCLE)]
                cycle_pos[pcat] += 1
            rdt = rdate(*STAGE_DATES[stage])
            sku = None
            if pcat == "Equipment Purchase":
                item, price, (lo, hi) = rng.choices(EQUIPMENT_ITEMS, weights=[4, 1, 1, 1, 1, 1, 1])[0]
                qty = rng.randint(lo, hi)
                lead = rng.randint(35, 70)
            elif pcat == "Spare Parts Replenishment":
                s = rng.choice(SPARES)
                sku, item, price = s[0], s[2], s[5]
                qty = rng.randint(2, 8) * 5
                lead = s[4]
            else:
                item, price = rng.choice(MAINT_ITEMS)
                price = round(price * rng.uniform(0.9, 1.1), -2)
                qty = 1
                lead = rng.randint(10, 30)
            row = {"Request Date": rdt, "Requesting Site": site, "Procurement Category": pcat, "SKU": sku,
                   "Item Description": item, "Quantity": qty, "Est. Unit Price (EUR)": price,
                   "Approval Stage": stage, "PO Reference": None, "Supplier": rng.choice(SUPPLIERS[pcat]),
                   "Expected Delivery Date": None, "Actual Delivery Date": None,
                   "Invoice Status": "Not Invoiced", "Invoice Date": None}
            if stage in ("PO Issued", "Received", "Invoiced"):
                row["PO Reference"] = f"PO-26-{po_counter:05d}"
                po_counter += 1
                row["Expected Delivery Date"] = rdt + timedelta(days=lead + 7)
            if stage in ("Received", "Invoiced"):
                row["Actual Delivery Date"] = min(rdt + timedelta(days=lead + rng.randint(0, 20)), REF - timedelta(days=2))
            if stage == "Invoiced":
                inv = row["Actual Delivery Date"] + timedelta(days=rng.randint(5, 30))
                row["Invoice Date"] = min(inv, REF - timedelta(days=1))
                row["Invoice Status"] = "Invoiced"
            if stage == "Cancelled":
                row["Invoice Status"] = "N/A"
            proc.append(row)

# Spare POs in "PO Issued" are what the Spare_Stock sheet counts as "On Order".
# Point them at specific spare lines so on-order quantities exist at those lines.
for row in proc:
    if row["Procurement Category"] == "Spare Parts Replenishment" and row["Approval Stage"] == "PO Issued":
        row["Requesting Site"] = rng.choice(SITES)
# Make one PO late: expected delivery already passed but nothing received
late = next(r for r in proc if r["Approval Stage"] == "PO Issued" and r["Procurement Category"] == "Equipment Purchase")
late["Expected Delivery Date"] = REF - timedelta(days=9)
late["Request Date"] = late["Expected Delivery Date"] - timedelta(days=55)
proc.sort(key=lambda r: r["Request Date"])
for i, r in enumerate(proc, start=1):
    r["Request ID"] = f"REQ-26-{i:03d}"


# Complete spare stock quantities: on order from POs, then on-hand/reserved to give planned flags
on_order = {}
for r in proc:
    if r["SKU"] and r["Approval Stage"] == "PO Issued":
        on_order[(r["SKU"], r["Requesting Site"])] = on_order.get((r["SKU"], r["Requesting Site"]), 0) + r["Quantity"]

flag_plan = {"AMS-01": 4, "HEL-01": 3, "FRA-01": 2, "STO-01": 1}
flagged = set()
for site, n in flag_plan.items():
    lines = [l for l in spare_lines if l["Site"] == site and (l["SKU"], site) not in on_order]
    for l in rng.sample(lines, n):
        flagged.add((l["SKU"], site))
critical = set(rng.sample(sorted(flagged), 3))

for l in spare_lines:
    key = (l["SKU"], l["Site"])
    oo = on_order.get(key, 0)
    reserved = rng.choice([0, 0, 0, 1, 1, 2, 3])
    rop = l["_rop"]
    if key in critical:
        position = rng.randint(0, max(0, l["Safety Stock"]))
    elif key in flagged:
        position = rng.randint(l["Safety Stock"] + 1, max(l["Safety Stock"] + 1, math.floor(rop)))
    else:
        position = rng.randint(math.floor(rop) + 1, l["Target Stock"] + math.ceil(l["Avg Monthly Consumption"]))
    on_hand = max(position - oo + reserved, reserved)
    l.update({"Current Stock (On Hand)": on_hand, "Reserved Stock": reserved, "Stock On Order": oo})
    del l["_rop"]

# ---------------------------------------------------------------------------
# 7. BUDGET INPUTS (approved budget + other planned spend), derived to give mixed variances
# ---------------------------------------------------------------------------
PCATS = list(STAGE_PLAN.keys())
OTHER_PLANNED = {"Equipment Purchase": 0, "Spare Parts Replenishment": 9000, "Maintenance Services": 15000}
OVERRUN = {("AMS-01", "Spare Parts Replenishment"): 0.90, ("FRA-01", "Maintenance Services"): 0.94,
           ("HEL-01", "Equipment Purchase"): 0.97}
budget = []
for site in SITES:
    for pcat in PCATS:
        rows = [r for r in proc if r["Requesting Site"] == site and r["Procurement Category"] == pcat]
        val = lambda r: r["Quantity"] * r["Est. Unit Price (EUR)"]
        actual = sum(val(r) for r in rows if r["Invoice Status"] == "Invoiced")
        commit = sum(val(r) for r in rows if r["Approval Stage"] in ("PO Issued", "Received"))
        appr = sum(val(r) for r in rows if r["Approval Stage"] == "Approved - No PO")
        pend = sum(val(r) for r in rows if r["Approval Stage"] == "Pending Approval") * PENDING_APPROVAL_RATE
        other = round(OTHER_PLANNED[pcat] * SITE_FACTOR[site], -2)
        forecast = actual + commit + appr + pend + other
        factor = OVERRUN.get((site, pcat), rng.uniform(1.05, 1.22))
        approved = max(round(forecast * factor, -4), 50000)
        budget.append({"Site": site, "Budget Category": pcat, "Approved Annual Budget (EUR)": approved,
                       "Other Planned Spend (EUR)": other})

# ---------------------------------------------------------------------------
# 8. WRITE CSVs
# ---------------------------------------------------------------------------
DATA.mkdir(exist_ok=True)
asset_cols = ["Asset ID", "Serial Number", "Category", "Manufacturer", "Model", "Location", "Rack Position",
              "Purchase Date", "Purchase Order ID", "Acquisition Cost (EUR)", "Warranty Expiration", "Status",
              "Assigned Team", "Last Audit Date", "Disposal Date"]
pd.DataFrame(sorted(assets, key=lambda a: a["Asset ID"]))[asset_cols].to_csv(DATA / "asset_register.csv", index=False)
pd.DataFrame(spare_lines)[["SKU", "Site", "Equipment Category", "Description", "Current Stock (On Hand)",
                           "Reserved Stock", "Stock On Order", "Avg Monthly Consumption", "Supplier Lead Time (Days)",
                           "Safety Stock", "Target Stock", "Unit Cost (EUR)"]].to_csv(DATA / "spare_stock.csv", index=False)
pd.DataFrame(rma_rows)[["RMA ID", "Asset ID", "Component Type", "Failed Part Serial", "Failure Description",
                        "Date Reported", "Case Created Date", "Vendor Status", "Case Status", "Spare Swap Asset ID",
                        "Replacement Tracking No.", "Closure Date"]].to_csv(DATA / "rma_cases.csv", index=False)
pd.DataFrame(proc)[["Request ID", "Request Date", "Requesting Site", "Procurement Category", "SKU",
                    "Item Description", "Quantity", "Est. Unit Price (EUR)", "Approval Stage", "PO Reference",
                    "Supplier", "Expected Delivery Date", "Actual Delivery Date", "Invoice Status",
                    "Invoice Date"]].to_csv(DATA / "procurement.csv", index=False)
pd.DataFrame(audit_rows)[["Audit ID", "Audit Date", "Asset ID", "Recorded Site", "Observed Site", "Recorded Status",
                          "Observed Status", "Auditor", "Corrective Action",
                          "Corrective Action Status"]].to_csv(DATA / "asset_audits.csv", index=False)
pd.DataFrame(budget).to_csv(DATA / "budget_inputs.csv", index=False)

meta = {"reference_date": REF.isoformat(), "seed": SEED, "showcase_gpu_server": showcase[0],
        "showcase_spare_gpu": showcase[1], "showcase_rma": showcase_rma, "exceptions": exceptions,
        "pending_approval_rate": PENDING_APPROVAL_RATE, "days_per_month": DAYS_PER_MONTH,
        "late_po_request": late["Request ID"]}
(DATA / "meta.json").write_text(json.dumps(meta, indent=2))
print(json.dumps(meta, indent=2))
print(f"assets={len(assets)} rma={len(rma_rows)} procurement={len(proc)} audits={len(audit_rows)} "
      f"spare_lines={len(spare_lines)} budget_lines={len(budget)}")
