"""Streamlit front end for the Northstar Budget Repair Agent.

    streamlit run app_streamlit.py

"Repair the workbook" is the only control that calls the Gemini API; it is billed to
GEMINI_API_KEY in this project's .env. "Show the latest recorded run" opens the newest
runs/run_*.json with no API call.
"""
from __future__ import annotations

import base64
import html
import io
import json
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image

from agent import BILLING_TIER, CAPS, MODEL, PRICE, THINKING_LEVEL, load_api_key

HERE = Path(__file__).resolve().parent
ASSETS = HERE / "assets"
RUNS = HERE / "runs"

# Google palette, Partner Marketing Hub. White text needs 4.5:1, so three swatches are darkened with black.
BLUE = "#3974D4"   # Blue 500 #4285F4 + 13% black: white text 4.54:1 (Blue 500 itself: 3.56:1). Set in config.toml.
RED = "#D73E31"    # Red 500 #EA4335 + 8% black: white text 4.54:1 (Red 500 itself: 3.92:1)
GREEN = "#2A8642"  # Green 500 #34A853 + 20% black: white text 4.58:1 (Green 500 itself: 3.06:1)
YELLOW = "#FBBC05"  # Yellow 500 keeps black text: 12.3:1
GREY_700 = "#5F6368"  # secondary text: 6.05:1 on white
BORDER = "#D7D9DB"    # Grey 500 #9AA0A6 at 40% over white, borders only

FEATURES = [
    (":material/forum:", "Interactions API",
     "One stored interaction chain. tools, system_instruction, and generation_config are sent again on every turn."),
    (":material/table_chart:", "Native Excel input",
     "Turn 1 sends northstar_q4_budget.xlsx as an inline document. The Sheets tab compares what Hydrogen saw with Python."),
    (":material/build:", "Four Python tools",
     "inspect_cells, trace_dependencies, test_formula_patch, and submit_repair. Python validates and runs every call."),
    (":material/visibility_off:", "Hidden evaluator",
     "Seven withheld scenarios scored against a plain-Python version of rules R1 to R7. The model never calls it."),
    (":material/lock:", "Frozen repair proposals",
     "Proposals are hashed and saved before evaluation starts. No hidden result goes back to the model."),
    (":material/payments:", "Usage-based cost",
     "Cost comes from each turn's usage: uncached input, cached input, output, and thinking tokens."),
]

st.set_page_config(page_title="Northstar Budget Repair", page_icon=":material/table_chart:", layout="wide")
if (ASSETS / "datacamp-logo.png").exists():
    st.logo(str(ASSETS / "datacamp-logo.png"), link="https://www.datacamp.com", size="large")
else:
    st.info("assets/datacamp-logo.png is missing, so the page shows no DataCamp logo.")

st.markdown(
    f"""
    <style>
      .stApp::before {{ content: ""; position: fixed; top: 0; left: 0; right: 0; height: 3px; z-index: 1000000;
        background: linear-gradient(90deg, #EA4335, #FBBC05, #34A853, #4285F4); }}
      .stApp, .stApp p, .stApp li, .stApp label, .stApp button {{
        font-family: "Google Sans", "Product Sans", Arial, sans-serif; }}
      .block-container {{ padding-top: 2rem; max-width: 1180px; }}
      @media (max-width: 640px) {{ .block-container {{ padding-top: 4.5rem; }} }}
      [data-testid="stCaptionContainer"] {{ color: {GREY_700}; opacity: 1; }}
      .brand {{ display: flex; align-items: center; gap: .75rem; flex-wrap: wrap; margin-bottom: .25rem; }}
      .brand img {{ height: 26px; }}
      .brand .model {{ font-family: ui-monospace, Menlo, Consolas, monospace; font-size: .8rem; background: #FFFFFF;
        color: #000000; border: 1px solid {BORDER}; border-radius: .6rem; padding: .15rem .65rem; }}
      .pill {{ display: inline-block; border-radius: 999px; padding: .2rem .8rem; font-weight: 700;
        font-size: .9rem; margin: 0 .4rem .4rem 0; }}
      .pill-complete {{ background: {GREEN}; color: #FFFFFF; }}
      .pill-partial {{ background: {YELLOW}; color: #000000; }}
      .pill-failure {{ background: {RED}; color: #FFFFFF; }}
      .pill-neutral {{ background: #FFFFFF; color: #000000; border: 1px solid {BORDER}; }}
      .stats {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: .6rem;
        margin: .6rem 0 1rem; }}
      .stat {{ background: #FFFFFF; border: 1px solid {BORDER}; border-radius: .6rem; padding: .55rem .8rem; }}
      .stat .label {{ font-size: .78rem; color: {GREY_700}; }}
      .stat .value {{ font-size: 1.25rem; font-weight: 700; color: #000000;
        font-family: ui-monospace, Menlo, Consolas, monospace; }}
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data
def gemini_logo_b64() -> str | None:
    """The logo file has wide white margins; trim them in memory so 26px shows the mark, not padding."""
    path = ASSETS / "gemini-logo.png"
    if not path.exists():
        return None
    im = Image.open(path).convert("RGBA")
    box = im.getchannel("A").getbbox()
    buf = io.BytesIO()
    (im.crop(box) if box else im).save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()


logo = gemini_logo_b64()
if logo:
    st.markdown(f'<div class="brand"><img src="data:image/png;base64,{logo}" alt="Gemini logo"/>'
                f'<span class="model">{html.escape(MODEL)}</span></div>', unsafe_allow_html=True)
else:
    st.info("assets/gemini-logo.png is missing, so the brand row shows the model id only.")
    st.markdown(f'<div class="brand"><span class="model">{html.escape(MODEL)}</span></div>', unsafe_allow_html=True)
st.title("Northstar Budget Repair Agent")
st.markdown("Hydrogen investigates Northstar's Q4 workbook, and Python checks every proposed formula repair.")

# ----------------------------------------------------------------- sidebar
with st.sidebar:
    st.subheader("Run settings")
    st.markdown(
        f"- Model: `{MODEL}`\n"
        f"- `thinking_level`: `{THINKING_LEVEL}`\n"
        f"- App limits: {CAPS['max_api_turns']} API turns, {CAPS['max_tool_calls']} tool calls, "
        f"{CAPS['max_seconds'] // 60} min, no new turn after ${CAPS['max_spend_usd']:.2f}"
    )
    st.subheader("Price per million tokens")
    st.markdown(
        "Gemini Developer API, paid Standard, checked November 13, 2026:\n"
        f"- Input: ${PRICE['input']:.2f}\n"
        f"- Output, including thinking: ${PRICE['output_incl_thinking']:.2f}\n"
        f"- Cached input: ${PRICE['cached_input']:.3f}"
    )
    st.caption('The Developer API free tier is "Free of charge." Agent Platform rates are a different table.')
    if st.session_state.get("live_done"):
        if BILLING_TIER == "free":
            st.metric("Spent this session", "$0.00", delta_color="off")
            st.caption("Free tier: the paid Standard rate was not charged.")
        else:
            st.metric("Spent this session", f"${st.session_state.get('spend', 0.0):.4f}", delta_color="off")

# ---------------------------------------------------------------- features
st.subheader("What this run uses")
cols = st.columns(2)
for i, (icon, name, desc) in enumerate(FEATURES):
    with cols[i % 2], st.container(border=True):
        st.markdown(f"{icon} **{name}**")
        st.caption(desc)

# --------------------------------------------------------------- controls
running = st.session_state.get("running", False)


def start_run():
    st.session_state["running"] = True


st.write("")
left, right = st.columns(2)
with left:
    st.button("Repair the workbook", type="primary", icon=":material/build:", width="stretch",
              disabled=running, on_click=start_run, help="Starts a live run that calls the Gemini API.")
with right:
    recorded = st.button("Show the latest recorded run", icon=":material/history:", width="stretch",
                         disabled=running, help="Opens the newest file in runs/ without calling the API.")
st.caption("A live run calls the Gemini API and is billed to GEMINI_API_KEY in .env.")


# ---------------------------------------------------------------- helpers
def stats_html(items) -> str:
    cells = "".join(f"<div class='stat'><div class='label'>{html.escape(a)}</div>"
                    f"<div class='value'>{html.escape(b)}</div></div>" for a, b in items)
    return f"<div class='stats'>{cells}</div>"


def pill(text: str, kind: str) -> str:
    return f"<span class='pill pill-{html.escape(kind)}'>{html.escape(text)}</span>"


def short(value, limit=140) -> str:
    text = value if isinstance(value, str) else json.dumps(value)
    text = " ".join(text.split()).replace("`", "'")
    return text if len(text) <= limit else text[: limit - 3] + "..."


TOOL_BADGE = {"inspect_cells": ":blue-badge[read]", "trace_dependencies": ":blue-badge[read]",
              "test_formula_patch": ":violet-badge[test]", "submit_repair": ":orange-badge[submit]"}
ARG_KEYS = ("sheet", "cell", "cell_range", "replacement_formula")


def log_turn(log, t: dict) -> None:
    """Replay one recorded turn as log lines. Every badge carries a text label."""
    u = t["usage"]
    log.markdown(f"`turn {t['n']:>2}` :material/memory: **{t['phase']}** · `{t['status']}` · "
                 f"{u['total_input_tokens']:,} in ({u['total_cached_tokens']:,} cached) · "
                 f"{u['total_thought_tokens']:,} thinking · {u['total_output_tokens']:,} out · {t['latency_s']} s")
    for c in t["function_calls"]:
        args = short(", ".join(f"{k}={v}" for k, v in c["arguments"].items() if k in ARG_KEYS))
        if c["is_error"]:
            try:
                reason = json.loads(c["result"]).get("error", c["result"])
            except (ValueError, AttributeError):
                reason = c["result"]
            log.markdown(f":material/subdirectory_arrow_right: :red-badge[rejected] `{c['name']}` `{args}`: "
                         f"{short(reason, 160)}")
        else:
            log.markdown(f":material/subdirectory_arrow_right: {TOOL_BADGE.get(c['name'], ':blue-badge[tool]')} "
                         f"`{c['name']}` `{args}`")


LIVE_LINE = re.compile(r"^\s+tool (\w+)\((.*)\) -> (ok|error)$")


def log_live(log, line: str) -> None:
    """Format one progress line from agent.py while a live run streams."""
    m = LIVE_LINE.match(line)
    if m:
        name, args, result = m.groups()
        badge = ":red-badge[rejected]" if result == "error" else TOOL_BADGE.get(name, ":blue-badge[tool]")
        log.markdown(f":material/subdirectory_arrow_right: {badge} `{name}` `{short(args)}`")
    elif line.startswith("turn"):
        log.markdown(f":material/memory: {short(line, 220)}")
    else:
        log.markdown(f":material/info: {short(line, 220)}")


def latest_run() -> Path | None:
    files = sorted(RUNS.glob("run_*.json")) if RUNS.exists() else []
    return files[-1] if files else None


def yes(v) -> str:
    return "yes" if v else "no"


def render_run(rec: dict) -> None:
    ev, totals = rec.get("evaluation"), rec.get("totals", {})
    with st.status("Completed run", state="complete", expanded=False):
        log = st.container(height=420, border=True)
        for t in rec.get("turns", []):
            log_turn(log, t)
        fr = rec.get("frozen_proposals", {})
        if fr:
            log.markdown(f":green-badge[frozen] :material/lock: {len(fr.get('proposals', []))} proposals, "
                         f"sha256 `{fr.get('sha256', '')[:12]}`")
        if ev:
            log.markdown(f":material/rule: hidden evaluator: {ev['hidden_passed']} of {ev['hidden_total']} "
                         f"scenarios, verdict **{ev['verdict']}**")
    st.divider()

    verdict = ev["verdict"] if ev else None
    kinds = {"complete repair": "complete", "partial repair": "partial", "failure": "failure"}
    pills = pill(verdict, kinds[verdict]) if verdict else pill("not evaluated", "neutral")
    pills += pill(f"thinking_level: {rec.get('thinking_level', '')}", "neutral")
    pills += pill(f"Stopped by: {rec['stopped_by']}" if rec.get("stopped_by") else "Stopped by: model", "neutral")
    st.markdown(pills, unsafe_allow_html=True)

    free = rec.get("billing_tier") == "free"
    cost = totals.get("cost_usd_at_paid_standard", {}).get("total", 0.0)
    items = []
    if ev:
        n = len(ev["per_error"])
        items += [("Errors repaired", f"{sum(e['repaired'] for e in ev['per_error'].values())} of {n}"),
                  ("False positives", str(len(ev["false_positives"]))),
                  ("Exceptions kept", f"{sum(ev['exceptions_preserved'].values())} of {len(ev['exceptions_preserved'])}"),
                  ("Hidden scenarios", f"{ev['hidden_passed']} of {ev['hidden_total']}")]
    items += [("API turns", str(totals.get("api_turns", 0))), ("Tool calls", str(totals.get("tool_calls", 0))),
              ("Elapsed time", f"{totals.get('elapsed_s', 0):.1f} s"),
              ("Cost" if not free else "Cost (free tier)", "$0.00" if free else f"${cost:.4f}")]
    st.markdown(stats_html(items), unsafe_allow_html=True)

    tabs = st.tabs([":material/description: Report", ":material/table_chart: Sheets", ":material/edit_note: Patches",
                    ":material/rule: Evaluator", ":material/visibility_off: Hidden tests", ":material/payments: Cost"])
    reports = rec.get("phase_reports", {})
    with tabs[0]:
        st.caption("What Hydrogen said. Nothing on this tab is checked; the Evaluator and Hidden tests tabs are.")
        if reports.get("final"):
            with st.container(border=True):
                st.markdown(reports["final"])
        for phase, label in (("read_only", "Turn 1: what it saw in the workbook"),
                             ("diagnosis", "Turn 2: first diagnosis, no tools"),
                             ("investigate", "Investigation findings")):
            if reports.get(phase):
                with st.expander(label):
                    st.markdown(reports[phase])
        calls = [{"Turn": t["n"], "Phase": t["phase"], "Tool": c["name"],
                  "Arguments": ", ".join(f"{k}={v}" for k, v in c["arguments"].items() if k in ARG_KEYS),
                  "Result": "error" if c["is_error"] else "ok"}
                 for t in rec.get("turns", []) for c in t["function_calls"]]
        if calls:
            st.markdown("**Tool calls Python ran**")
            st.dataframe(pd.DataFrame(calls), hide_index=True, width="stretch")
    with tabs[1]:
        st.caption("Hydrogen's turn-1 inventory from the native .xlsx, next to openpyxl's inventory of the same file.")
        inv, seen = rec.get("python_inventory", {}), rec.get("read_only_json") or {}
        names = [w.get("name", "") for w in seen.get("worksheets", [])]
        st.dataframe(pd.DataFrame([{"Python: sheet": s["name"], "Python: dimensions": s["dimensions"],
                                    "Python: formula cells": s["formula_cells"],
                                    "Hydrogen: sheet as reported": names[i] if i < len(names) else "not reported"}
                                   for i, s in enumerate(inv.get("sheets", []))]), hide_index=True, width="stretch")
        left, right = st.columns(2)
        left.metric("Formula cells in the file", f"{sum(s['formula_cells'] for s in inv.get('sheets', []))}",
                    delta_color="off")
        right.metric("Hydrogen said it saw formulas", yes(seen.get("can_see_formulas")), delta_color="off")
        first = (rec.get("turns") or [{}])[0]
        if first.get("input_by_modality"):
            st.caption("Turn 1 input tokens by modality: " +
                       ", ".join(f"{m['modality']} {m['tokens']:,}" for m in first["input_by_modality"]))
    with tabs[2]:
        props = rec.get("frozen_proposals", {}).get("proposals", [])
        if props:
            st.dataframe(pd.DataFrame([{"Cell": f"{p['sheet']}!{p['cell']}", "Original": p["original_formula"],
                                        "Replacement": p["replacement_formula"], "Status": p["status"],
                                        "Expected effect": p.get("expected_effect", "")} for p in props]),
                         hide_index=True, width="stretch")
            with st.expander("Evidence Hydrogen gave for each submission"):
                for p in props:
                    st.markdown(f"**{p['sheet']}!{p['cell']}**: {p.get('evidence', '')}")
        else:
            st.info("No repair was submitted in this run.")
        st.caption(f"Unresolved cases Hydrogen reported: {len(rec.get('frozen_proposals', {}).get('unresolved', []))}")
    with tabs[3]:
        st.caption("Harness output. Python applied the frozen patches to a new copy and compared the two files.")
        if not ev:
            st.info("This run has no evaluation yet.")
        else:
            with st.container(border=True):
                st.markdown(":material/shield: **Harness-only: planted errors** (shown after the freeze)")
                st.dataframe(pd.DataFrame([{"Planted error": c, "What it was": e["kind"], "Proposed": yes(e["proposed"]),
                                            "Replacement": e.get("replacement", ""), "Repaired": yes(e["repaired"])}
                                           for c, e in ev["per_error"].items()]), hide_index=True, width="stretch")
            d = ev["structural_diff"]
            st.dataframe(pd.DataFrame([
                {"Check": "Patches applied", "Result": ", ".join(ev["applied"]) or "none"},
                {"Check": "Stale patches rejected", "Result": ", ".join(ev["stale_rejected"]) or "none"},
                {"Check": "Sheet names and order unchanged", "Result": yes(d["sheets_unchanged"])},
                {"Check": "Transactions and Assumptions unchanged", "Result": yes(d["input_sheets_unchanged"])},
                {"Check": "Unauthorized cell changes", "Result": str(len(d["unauthorized_changes"]))},
                {"Check": "Number-format changes", "Result": str(len(d["number_format_changes"]))},
                {"Check": "Valid exceptions unchanged",
                 "Result": ", ".join(f"{c}: {yes(v)}" for c, v in ev["exceptions_preserved"].items())},
                {"Check": "False positives", "Result": ", ".join(ev["false_positives"]) or "none"},
                {"Check": "LibreOffice versus engine mismatches", "Result": str(len(ev["libreoffice_mismatches"]))},
            ]), hide_index=True, width="stretch")
    with tabs[4]:
        if not ev:
            st.info("Hidden tests run only after the proposals are frozen.")
        else:
            st.caption("Withheld input scenarios. Expected values come from the plain-Python reference, never the workbook.")
            st.dataframe(pd.DataFrame([{"Scenario": name, "Passed": yes(r["passed"]),
                                        "Expected status": r["status_expected"],
                                        "Expected remaining budget": f"${r['remaining_expected']:,.2f}",
                                        "Failing cells": len(r["failures"])}
                                       for name, r in ev["hidden_scenarios"].items()]), hide_index=True, width="stretch")
    with tabs[5]:
        u, parts = totals.get("usage", {}), totals.get("cost_usd_at_paid_standard", {})
        uncached = u.get("total_input_tokens", 0) - u.get("total_cached_tokens", 0)
        k = st.columns(4)
        k[0].metric("Uncached input", f"{uncached:,}", f"${parts.get('input', 0):.4f}", delta_color="off")
        k[1].metric("Cached input", f"{u.get('total_cached_tokens', 0):,}", f"${parts.get('cached_input', 0):.4f}",
                    delta_color="off")
        k[2].metric("Output", f"{u.get('total_output_tokens', 0):,}", delta_color="off")
        k[3].metric("Thinking", f"{u.get('total_thought_tokens', 0):,}",
                    f"${parts.get('output_incl_thinking', 0):.4f} with output", delta_color="off")
        st.dataframe(pd.DataFrame([{"Turn": t["n"], "Phase": t["phase"], "Seconds": t["latency_s"],
                                    "Input": t["usage"]["total_input_tokens"], "Cached": t["usage"]["total_cached_tokens"],
                                    "Thinking": t["usage"]["total_thought_tokens"], "Output": t["usage"]["total_output_tokens"],
                                    "Tool calls": len(t["function_calls"]), "USD": round(t["cost_usd"], 4)}
                                   for t in rec.get("turns", [])]), hide_index=True, width="stretch")
        if free:
            st.caption("Free tier: $0.00 charged. Dollar figures show the paid Standard rate for reference only.")
        else:
            st.caption(f"Total ${parts.get('total', 0):.4f} at the paid Standard rate, from the usage object on each turn.")


# ------------------------------------------------------------------ actions
if running:
    try:
        load_api_key()
        key_ok = True
    except RuntimeError:
        key_ok = False
    if not key_ok:
        st.session_state["running"] = False
        st.session_state["error"] = "GEMINI_API_KEY was not found in this project's .env file, so no request was sent."
        st.rerun()
    from agent import Agent  # imported here so loading the page never touches the API

    with st.status("Repairing the workbook...", expanded=True) as status:
        log = st.container(height=420, border=True)
        try:
            path = Agent(progress=lambda line: log_live(log, line)).run()
            record = json.loads(path.read_text(encoding="utf-8"))
            st.session_state["spend"] = (st.session_state.get("spend", 0.0)
                                         + record["totals"]["cost_usd_at_paid_standard"]["total"])
            st.session_state["record"] = record
            st.session_state["live_done"] = True
            status.update(label="Run finished", state="complete", expanded=False)
        except Exception as exc:  # show the failure without a traceback
            st.session_state["error"] = f"The run stopped: {type(exc).__name__}: {exc}"
            status.update(label="Run failed", state="error")
    st.session_state["running"] = False
    st.rerun()

if st.session_state.get("error"):
    st.error(st.session_state.pop("error"))

if recorded:
    path = latest_run()
    if path is None:
        st.session_state["record"] = None
        st.info("No recorded runs yet. Start one with Repair the workbook.")
    else:
        st.session_state["record"] = json.loads(path.read_text(encoding="utf-8"))

if st.session_state.get("record"):
    render_run(st.session_state["record"])
