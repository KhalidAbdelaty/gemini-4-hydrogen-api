"""Build the Northstar Retail Q4 workbook.

Writes two files:
  evaluation/northstar_reference.xlsx  correct formulas (evaluator only)
  northstar_q4_budget.xlsx             working copy with 3 planted errors

The working copy is recalculated with LibreOffice so its cached values match
its own (wrong) formulas, not the reference answers. evaluation/manifest.json
records the planted-error map and SHA-256 hashes before any API call.
The agent's tools only ever open northstar_q4_budget.xlsx.
"""
from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook as XLWorkbook
from openpyxl import load_workbook
from openpyxl.styles import Font

from formula_engine import Workbook, find_soffice

ROOT = Path(__file__).parent
WORKING = ROOT / "northstar_q4_budget.xlsx"
EVAL_DIR = ROOT / "evaluation"
REFERENCE = EVAL_DIR / "northstar_reference.xlsx"
MANIFEST = EVAL_DIR / "manifest.json"

SHEETS = ["Summary", "Forecast", "Assumptions", "Transactions"]  # Summary first: report layout
TX_LAST = 300  # SUMIFS ranges cover Transactions!2:300 so appended rows count

DEPARTMENTS = [("Fulfillment", 0.06), ("Marketing", 0.18), ("Customer Support", 0.10), ("Technology", 0.04)]
GROWTH_ROW = {name: 4 + i for i, (name, _) in enumerate(DEPARTMENTS)}  # Assumptions!B4:B7
VENDORS = [("ParcelPoint", 120_000, 0.04), ("CloudNimbus", 140_000, 0.03)]  # Assumptions rows 11-12
AD_CAP, BUDGET = 165_000, 1_415_000
CATEGORIES = ["Payroll", "Shipping", "Facilities", "Advertising", "Software",
              "Contractors", "Cloud hosting", "Payment processing"]

# Forecast lines: (department, category, monthly amounts by vendor)
LINES = [
    ("Fulfillment", "Payroll", [("Northstar Payroll", 70_000)]),
    ("Fulfillment", "Shipping", [("ParcelPoint", 40_000), ("RouteFast", 12_000)]),
    ("Fulfillment", "Facilities", [("Harbor Logistics Park", 30_000)]),
    ("Marketing", "Payroll", [("Northstar Payroll", 45_000)]),
    ("Marketing", "Advertising", [("AdReach", 50_000)]),
    ("Marketing", "Software", [("CampaignHub", 6_000)]),
    ("Customer Support", "Payroll", [("Northstar Payroll", 38_000)]),
    ("Customer Support", "Contractors", [("HelpDesk Partners", 15_000)]),
    ("Customer Support", "Software", [("TicketLoop", 4_000)]),
    ("Technology", "Payroll", [("Northstar Payroll", 60_000)]),
    ("Technology", "Cloud hosting", [("CloudNimbus", 50_000)]),
    ("Technology", "Payment processing", [("PayStream", 30_000)]),
]
CAPEX = [("2026-08", "Fulfillment", "Capital equipment", "Conveyor Systems Co", 85_000),
         ("2026-09", "Technology", "Capital equipment", "RackWorks", 40_000)]

RULES = [
    "R1. Each forecast line: Q4 projected = Q3 actual for that department and category x (1 + that department's Q4 growth rate).",
    "R2. Only the approved categories listed above are operating expenses. Capital equipment is capex and is excluded.",
    "R3. The Harbor Logistics Park warehouse lease (Fulfillment, Facilities) is a fixed-price contract: Q4 projected equals Q3 actual, with no growth.",
    "R4. Q4 paid advertising (Marketing, Advertising) is capped at the Q4 paid advertising cap above.",
    "R5. A vendor rebate applies when Q3 spend with that vendor is at or above the vendor's threshold. Rebate = Q3 vendor spend x rebate rate, credited against Q4 operating expenses.",
    "R6. Net Q4 operating expenses = total projected operating expenses for all approved categories - total vendor rebates.",
    "R7. Remaining budget = approved Q4 operating budget - net Q4 operating expenses. Status is Over budget when remaining budget is below zero.",
]

FORECAST_FIRST, FORECAST_LAST = 4, 4 + len(LINES) - 1  # rows 4-15
REBATE_ROWS = {"ParcelPoint": 20, "CloudNimbus": 21}
SUMMARY_FIRST = 4
SUMMARY_LAST = SUMMARY_FIRST + len(CATEGORIES) - 1  # rows 4-11

# The three planted errors (reader-facing only; never sent to the model).
PLANTED = {
    "Summary!B13": {"reference": f"=SUM(B{SUMMARY_FIRST}:B{SUMMARY_LAST})",
                    "broken": f"=SUM(B{SUMMARY_FIRST}:B{SUMMARY_LAST - 1})",
                    "kind": "excluded approved category (Payment processing) from the total"},
    "Forecast!D14": {"reference": f"=Assumptions!$B${GROWTH_ROW['Technology']}",
                     "broken": f"=Assumptions!$B${GROWTH_ROW['Marketing']}",
                     "kind": "Technology cloud hosting uses Marketing's growth rate"},
    "Forecast!E20": {"reference": "=IF(B20>=C20,B20*D20,0)",
                     "broken": "=IF(B20>C20,B20*D20,0)",
                     "kind": "rebate threshold uses > instead of >="},
}
# Two unusual formulas that follow documented rules and must not change.
EXCEPTIONS = {"Forecast!E6": "R3 fixed-price warehouse lease (no growth)",
              "Forecast!E8": "R4 paid advertising cap (MIN)"}


def transactions():
    rows, n = [], 1
    for month in ("2026-07", "2026-08", "2026-09"):
        for dept, cat, vendors in LINES:
            for vendor, amount in vendors:
                rows.append((f"T{n:03d}", month, dept, cat, vendor, amount))
                n += 1
        for m, dept, cat, vendor, amount in CAPEX:
            if m == month:
                rows.append((f"T{n:03d}", month, dept, cat, vendor, amount))
                n += 1
    return rows


def write_workbook(path: Path, broken: bool):
    wb = XLWorkbook()
    wb.remove(wb.active)
    ws = {name: wb.create_sheet(name) for name in SHEETS}
    bold = Font(bold=True)
    money, pct = '"$"#,##0', "0%"

    t = ws["Transactions"]
    t.append(["Txn ID", "Period", "Department", "Category", "Vendor", "Amount"])
    for row in transactions():
        t.append(list(row))
    for c in t["F"][1:]:
        c.number_format = money

    a = ws["Assumptions"]
    a["A1"] = "Northstar Retail: Q4 2026 assumptions and business rules"
    a["A3"], a["B3"] = "Department", "Q4 growth rate"
    for name, rate in DEPARTMENTS:
        r = GROWTH_ROW[name]
        a[f"A{r}"], a[f"B{r}"] = name, rate
        a[f"B{r}"].number_format = pct
    a["A9"] = "Vendor rebates"
    a["A10"], a["B10"], a["C10"] = "Vendor", "Q3 spend threshold", "Rebate rate"
    for i, (vendor, threshold, rate) in enumerate(VENDORS):
        r = 11 + i
        a[f"A{r}"], a[f"B{r}"], a[f"C{r}"] = vendor, threshold, rate
        a[f"B{r}"].number_format, a[f"C{r}"].number_format = money, pct
    a["A14"] = "Other parameters"
    a["A15"], a["B15"] = "Q4 paid advertising cap", AD_CAP
    a["A16"], a["B16"] = "Approved Q4 operating budget", BUDGET
    a["B15"].number_format = a["B16"].number_format = money
    a["A18"] = "Approved operating expense categories"
    for i, cat in enumerate(CATEGORIES):
        a[f"A{19 + i}"] = cat
    a["A28"] = "Business rules"
    for i, rule in enumerate(RULES):
        a[f"A{29 + i}"] = rule

    f = ws["Forecast"]
    f["A1"] = "Q4 2026 operating expense forecast"
    for col, head in zip("ABCDEF", ["Department", "Category", "Q3 actual", "Q4 growth", "Q4 projected", "Note"]):
        f[f"{col}3"] = head
    rng = lambda col: f"Transactions!${col}$2:${col}${TX_LAST}"
    for i, (dept, cat, _) in enumerate(LINES):
        r = FORECAST_FIRST + i
        f[f"A{r}"], f[f"B{r}"] = dept, cat
        f[f"C{r}"] = f"=SUMIFS({rng('F')},{rng('C')},$A{r},{rng('D')},$B{r})"
        if cat == "Facilities":
            f[f"D{r}"] = "fixed"
            f[f"E{r}"] = f"=C{r}"
            f[f"F{r}"] = "Lease is fixed-price (R3)"
        else:
            f[f"D{r}"] = f"=Assumptions!$B${GROWTH_ROW[dept]}"
            f[f"E{r}"] = f"=MIN(C{r}*(1+D{r}),Assumptions!$B$15)" if cat == "Advertising" else f"=C{r}*(1+D{r})"
            if cat == "Advertising":
                f[f"F{r}"] = "Capped (R4)"
        f[f"C{r}"].number_format = f[f"E{r}"].number_format = money
        f[f"D{r}"].number_format = pct
    f["A18"] = "Vendor rebates (credited in Q4)"
    for col, head in zip("ABCDE", ["Vendor", "Q3 qualifying spend", "Threshold", "Rate", "Rebate"]):
        f[f"{col}19"] = head
    for i, (vendor, _, _) in enumerate(VENDORS):
        r = REBATE_ROWS[vendor]
        f[f"A{r}"] = vendor
        f[f"B{r}"] = f"=SUMIFS({rng('F')},{rng('E')},A{r})"
        f[f"C{r}"] = f"=Assumptions!$B${11 + i}"
        f[f"D{r}"] = f"=Assumptions!$C${11 + i}"
        f[f"E{r}"] = f"=IF(B{r}>=C{r},B{r}*D{r},0)"
        for col in "BCE":
            f[f"{col}{r}"].number_format = money
        f[f"D{r}"].number_format = pct
    f["A22"], f["E22"] = "Total rebates", "=SUM(E20:E21)"
    f["E22"].number_format = money

    s = ws["Summary"]
    s["A1"] = "Northstar Retail: Q4 2026 operating expense summary"
    s["A3"], s["B3"] = "Category", "Q4 projected"
    for i, cat in enumerate(CATEGORIES):
        r = SUMMARY_FIRST + i
        s[f"A{r}"] = cat
        s[f"B{r}"] = f"=SUMIFS(Forecast!$E${FORECAST_FIRST}:$E${FORECAST_LAST},Forecast!$B${FORECAST_FIRST}:$B${FORECAST_LAST},A{r})"
    labels = {13: ("Total projected operating expenses", f"=SUM(B{SUMMARY_FIRST}:B{SUMMARY_LAST})"),
              14: ("Vendor rebates", "=Forecast!E22"),
              15: ("Net Q4 operating expenses", "=B13-B14"),
              16: ("Approved Q4 operating budget", "=Assumptions!B16"),
              17: ("Remaining budget", "=B16-B15"),
              18: ("Budget status", '=IF(B17<0,"Over budget","Within budget")')}
    for r, (label, formula) in labels.items():
        s[f"A{r}"], s[f"B{r}"] = label, formula
    for r in list(range(SUMMARY_FIRST, SUMMARY_LAST + 1)) + [13, 14, 15, 16, 17]:
        s[f"B{r}"].number_format = money

    if broken:
        for target, spec in PLANTED.items():
            sheet, coord = target.split("!")
            ws[sheet][coord] = spec["broken"]

    for sheet in ws.values():
        sheet["A1"].font = bold
        sheet.column_dimensions["A"].width = 38 if sheet.title != "Transactions" else 10
        for col in "BCDEF":
            sheet.column_dimensions[col].width = 20
    ws["Assumptions"].column_dimensions["A"].width = 60
    wb.save(path)


def libreoffice_recalc(path: Path) -> Path:
    """Open and re-save with LibreOffice headless so cached values exist."""
    out = Path(tempfile.mkdtemp())
    subprocess.run([find_soffice(), "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(out), str(path)],
                   check=True, capture_output=True, timeout=120)
    return out / path.name


def cross_check(xlsx_with_cache: Path, label: str) -> dict:
    """Compare LibreOffice's cached values with formula_engine on every formula cell."""
    ours = Workbook.from_xlsx(xlsx_with_cache)
    cached = load_workbook(xlsx_with_cache, data_only=True)
    mismatches, n = [], 0
    for sheet, coord in ours.formula_cells():
        n += 1
        lo, mine = cached[sheet][coord].value, ours.value(sheet, coord)
        same = (lo == mine) if isinstance(mine, str) or isinstance(lo, str) else abs(float(lo or 0) - float(mine or 0)) < 1e-6
        if not same:
            mismatches.append(f"{sheet}!{coord}: libreoffice={lo} engine={mine}")
    return {"workbook": label, "formula_cells": n, "mismatches": mismatches}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    EVAL_DIR.mkdir(exist_ok=True)
    tmp = Path(tempfile.mkdtemp())
    checks = []
    for broken, dest in ((False, REFERENCE), (True, WORKING)):
        raw = tmp / dest.name
        write_workbook(raw, broken)
        recalced = libreoffice_recalc(raw)
        # LibreOffice must not rewrite any formula text we wrote.
        before, after = Workbook.from_xlsx(raw), Workbook.from_xlsx(recalced)
        assert before.cells == after.cells, "LibreOffice changed workbook contents"
        shutil.copy(recalced, dest)
        checks.append(cross_check(dest, dest.name))
    ref, work = Workbook.from_xlsx(REFERENCE), Workbook.from_xlsx(WORKING)
    manifest = {
        "frozen_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "sha256": {"working_workbook": sha256(WORKING), "reference_workbook": sha256(REFERENCE),
                   "evaluator": sha256(ROOT / "evaluator.py")},
        "libreoffice_cross_check": checks,
        "planted_errors": PLANTED,
        "valid_exceptions": EXCEPTIONS,
        "summary_reference": {c: ref.value("Summary", c) for c in ("B13", "B14", "B15", "B16", "B17", "B18")},
        "summary_working": {c: work.value("Summary", c) for c in ("B13", "B14", "B15", "B16", "B17", "B18")},
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2))
    print(json.dumps({k: manifest[k] for k in ("libreoffice_cross_check", "summary_reference", "summary_working")}, indent=2))


if __name__ == "__main__":
    main()
