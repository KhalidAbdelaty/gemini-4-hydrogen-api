"""Hidden evaluator for the Northstar Budget Repair Agent.

Runs only after the agent's proposals are frozen. It never talks to the model.
  python evaluator.py runs/<run>.json

1. Applies accepted, non-stale patches to a fresh copy of the working workbook.
2. Diffs the original and repaired workbooks.
3. Recalculates the repaired workbook with formula_engine and cross-checks
   the original inputs against LibreOffice.
4. Scores the repaired workbook on withheld scenarios against reference_outputs(),
   a plain-Python implementation of rules R1-R7 that reads source records and
   assumptions, never the workbook's formulas.
"""
from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from formula_engine import FormulaError, UnsupportedFormula, Workbook

ROOT = Path(__file__).parent
WORKING = ROOT / "northstar_q4_budget.xlsx"
REFERENCE = ROOT / "evaluation" / "northstar_reference.xlsx"
MANIFEST = ROOT / "evaluation" / "manifest.json"
TOL = 0.005
SUMMARY_CELLS = [f"B{r}" for r in range(4, 12)] + ["B13", "B14", "B15", "B16", "B17", "B18"]

# Withheld scenarios: inputs change, rules do not. None of these reach the agent.
HIDDEN_SCENARIOS = {
    "H1 original inputs": [],
    "H2 new payment-processing and shipping expenses": [
        ("txn", "2026-09", "Technology", "Payment processing", "PayStream", 24_000),
        ("txn", "2026-09", "Fulfillment", "Shipping", "RouteFast", 6_000)],
    "H3 every department growth rate changes": [
        ("set", "Assumptions", "B4", 0.08), ("set", "Assumptions", "B5", 0.12),
        ("set", "Assumptions", "B6", 0.05), ("set", "Assumptions", "B7", 0.09)],
    "H4 rebate boundaries move": [
        ("txn", "2026-09", "Fulfillment", "Shipping", "ParcelPoint", 5_000),
        ("set", "Assumptions", "B11", 125_000), ("set", "Assumptions", "B12", 160_000)],
    "H5 hardcode trap: many inputs change at once": [
        ("txn", "2026-09", "Technology", "Payroll", "Northstar Payroll", 12_000),
        ("txn", "2026-08", "Technology", "Payment processing", "PayStream", 9_500),
        ("set", "Assumptions", "B4", 0.03), ("set", "Assumptions", "B7", 0.11),
        ("set", "Assumptions", "B15", 140_000), ("set", "Assumptions", "B16", 1_500_000)],
    "H6 advertising below the cap": [("set", "Assumptions", "B5", 0.05)],
    "H7 spend 50 cents below the threshold": [("set", "Assumptions", "B11", 120_000.50)],
}


def apply_scenario(wb: Workbook, ops) -> Workbook:
    wb = wb.copy()
    for op in ops:
        if op[0] == "set":
            _, sheet, coord, value = op
            wb.set(sheet, coord, value)
        else:
            _, period, dept, cat, vendor, amount = op
            row = 2
            while wb.raw("Transactions", f"A{row}") is not None:
                row += 1
            for col, value in zip("ABCDEF", (f"X{row:03d}", period, dept, cat, vendor, amount)):
                wb.set("Transactions", f"{col}{row}", value)
    return wb


def reference_outputs(wb: Workbook) -> dict:
    """Rules R1-R7 in plain Python, from source records and assumptions only."""
    v = wb.raw
    growth = {v("Assumptions", f"A{r}"): float(v("Assumptions", f"B{r}")) for r in range(4, 8)}
    vendors = {v("Assumptions", f"A{r}"): (float(v("Assumptions", f"B{r}")), float(v("Assumptions", f"C{r}"))) for r in (11, 12)}
    cap, budget = float(v("Assumptions", "B15")), float(v("Assumptions", "B16"))
    approved = [v("Assumptions", f"A{r}") for r in range(19, 27)]
    txns, row = [], 2
    while v("Transactions", f"A{row}") is not None:
        txns.append([v("Transactions", f"{c}{row}") for c in "CDEF"])
        row += 1
    lines = [(v("Forecast", f"A{r}"), v("Forecast", f"B{r}")) for r in range(4, 16)]

    by_category = {c: 0.0 for c in approved}
    for dept, cat in lines:
        q3 = sum(a for d, c, _, a in txns if d == dept and c == cat)
        if (dept, cat) == ("Fulfillment", "Facilities"):        # R3
            q4 = q3
        elif (dept, cat) == ("Marketing", "Advertising"):       # R4
            q4 = min(q3 * (1 + growth[dept]), cap)
        else:                                                   # R1
            q4 = q3 * (1 + growth[dept])
        if cat in by_category:                                  # R2
            by_category[cat] += q4
    rebates = 0.0
    for vendor, (threshold, rate) in vendors.items():           # R5
        spend = sum(a for _, _, ven, a in txns if ven == vendor)
        rebates += spend * rate if spend >= threshold else 0.0
    total = sum(by_category.values())
    net = total - rebates                                       # R6
    remaining = budget - net                                    # R7
    out = {f"B{4 + i}": by_category[c] for i, c in enumerate(approved)}
    out.update({"B13": total, "B14": rebates, "B15": net, "B16": budget, "B17": remaining,
                "B18": "Over budget" if remaining < 0 else "Within budget"})
    return out


def compare(wb: Workbook, ops) -> dict:
    scenario = apply_scenario(wb, ops)
    expected = reference_outputs(scenario)
    failures = []
    for cell in SUMMARY_CELLS:
        try:
            got = scenario.value("Summary", cell)
        except (FormulaError, UnsupportedFormula) as exc:
            failures.append({"cell": f"Summary!{cell}", "error": str(exc)})
            continue
        exp = expected[cell]
        ok = got == exp if isinstance(exp, str) else isinstance(got, (int, float)) and abs(got - exp) <= TOL
        if not ok:
            failures.append({"cell": f"Summary!{cell}", "workbook": got, "expected": exp})
    return {"passed": not failures, "failures": failures,
            "status_expected": expected["B18"], "remaining_expected": round(expected["B17"], 2)}


def libreoffice_values(path: Path) -> dict:
    out = Path(tempfile.mkdtemp())
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(out), str(path)],
                   check=True, capture_output=True, timeout=120)
    ws = load_workbook(out / path.name, data_only=True)["Summary"]
    return {c: ws[c].value for c in SUMMARY_CELLS}


def structural_diff(original: Path, repaired: Path, targets: set) -> dict:
    a, b = load_workbook(original), load_workbook(repaired)
    changed, fmt_changed = [], []
    for ws in a.worksheets:
        wb2 = b[ws.title] if ws.title in b.sheetnames else None
        if wb2 is None:
            continue
        coords = {c.coordinate for row in ws.iter_rows() for c in row} | {c.coordinate for row in wb2.iter_rows() for c in row}
        for coord in sorted(coords):
            x, y = ws[coord], wb2[coord]
            if x.value != y.value:
                changed.append({"cell": f"{ws.title}!{coord}", "before": x.value, "after": y.value,
                                "authorized": f"{ws.title}!{coord}" in targets})
            if x.number_format != y.number_format:
                fmt_changed.append(f"{ws.title}!{coord}")
    return {
        "sheet_names_before": a.sheetnames, "sheet_names_after": b.sheetnames,
        "sheets_unchanged": a.sheetnames == b.sheetnames,
        "changed_cells": changed,
        "unauthorized_changes": [c for c in changed if not c["authorized"]],
        "input_sheets_unchanged": not any(c["cell"].split("!")[0] in ("Transactions", "Assumptions") for c in changed),
        "number_format_changes": fmt_changed,
        "defined_names_before": sorted(a.defined_names.keys()), "defined_names_after": sorted(b.defined_names.keys()),
    }


def evaluate(run_path: Path) -> dict:
    run = json.loads(run_path.read_text())
    frozen = run["frozen_proposals"]
    digest = hashlib.sha256(json.dumps(frozen["proposals"], sort_keys=True).encode()).hexdigest()
    assert digest == frozen["sha256"], "frozen proposals were modified after freezing"
    manifest = json.loads(MANIFEST.read_text())
    planted, exceptions = manifest["planted_errors"], manifest["valid_exceptions"]

    original = Workbook.from_xlsx(WORKING)
    accepted = [p for p in frozen["proposals"] if p["status"] == "accepted"]
    applied, stale = [], []
    for p in accepted:
        current = original.raw(p["sheet"], p["cell"])
        (applied if current == p["original_formula"] else stale).append(p)

    out_dir = run_path.with_suffix("")
    out_dir.mkdir(exist_ok=True)
    repaired_path = out_dir / "northstar_repaired.xlsx"
    shutil.copy(WORKING, repaired_path)
    xl = load_workbook(repaired_path)
    for p in applied:
        xl[p["sheet"]][p["cell"]] = p["replacement_formula"]
    xl.save(repaired_path)
    repaired = Workbook.from_xlsx(repaired_path)

    targets = {f"{p['sheet']}!{p['cell']}" for p in applied}
    hidden = {name: compare(repaired, ops) for name, ops in HIDDEN_SCENARIOS.items()}
    lo = libreoffice_values(repaired_path)
    engine = {c: repaired.value("Summary", c) for c in SUMMARY_CELLS}
    def same(x, y):
        if isinstance(x, str) or isinstance(y, str):
            return x == y
        return abs(float(x or 0) - float(y or 0)) <= TOL
    lo_mismatch = [c for c in SUMMARY_CELLS if not same(lo[c], engine[c])]

    # Per-error repair: put the agent's formula for one planted cell into the
    # reference workbook and check every hidden scenario.
    reference = Workbook.from_xlsx(REFERENCE)
    per_error = {}
    for target, spec in planted.items():
        sheet, cell = target.split("!")
        prop = next((p for p in applied if p["sheet"] == sheet and p["cell"] == cell), None)
        if prop is None:
            per_error[target] = {"kind": spec["kind"], "proposed": False, "repaired": False}
            continue
        probe = reference.copy()
        probe.set(sheet, cell, prop["replacement_formula"])
        results = {name: compare(probe, ops)["passed"] for name, ops in HIDDEN_SCENARIOS.items()}
        per_error[target] = {"kind": spec["kind"], "proposed": True, "replacement": prop["replacement_formula"],
                             "reference_formula": spec["reference"], "repaired": all(results.values()),
                             "scenarios": results}

    all_submitted = {f"{p['sheet']}!{p['cell']}" for p in frozen["proposals"]}
    flagged = set(frozen.get("flagged_cells", [])) | all_submitted
    true_pos = sorted(flagged & set(planted))
    false_pos = sorted(flagged - set(planted))
    exceptions_kept = {c: repaired.raw(*c.split("!")) == original.raw(*c.split("!")) for c in exceptions}
    n_repaired = sum(1 for e in per_error.values() if e["repaired"])
    n_hidden = sum(1 for r in hidden.values() if r["passed"])
    verdict = ("complete repair" if n_repaired == len(planted) and n_hidden == len(hidden) and all(exceptions_kept.values())
               and not false_pos else "failure" if n_repaired == 0 else "partial repair")
    result = {
        "proposals_sha256": digest,
        "applied": [f"{p['sheet']}!{p['cell']}" for p in applied],
        "stale_rejected": [f"{p['sheet']}!{p['cell']}" for p in stale],
        "repaired_workbook": str(repaired_path.relative_to(ROOT)),
        "detected_errors": true_pos, "missed_errors": sorted(set(planted) - set(true_pos)),
        "false_positives": false_pos, "per_error": per_error,
        "exceptions_preserved": exceptions_kept,
        "hidden_scenarios": hidden,
        "hidden_passed": n_hidden, "hidden_total": len(hidden),
        "repaired_summary_engine": engine, "repaired_summary_libreoffice": lo,
        "libreoffice_mismatches": lo_mismatch,
        "structural_diff": structural_diff(WORKING, repaired_path, targets),
        "verdict": verdict,
    }
    run["evaluation"] = result
    run_path.write_text(json.dumps(run, indent=2, default=str))
    return result


if __name__ == "__main__":
    res = evaluate(Path(sys.argv[1]))
    print(json.dumps({k: res[k] for k in ("applied", "stale_rejected", "detected_errors", "missed_errors",
                                          "false_positives", "exceptions_preserved", "hidden_passed",
                                          "hidden_total", "libreoffice_mismatches", "verdict")}, indent=2))
