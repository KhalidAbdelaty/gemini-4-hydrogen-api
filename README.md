# Northstar Budget Repair Agent

A Python agent that uses Gemini 4 Hydrogen (`gemini-4-hydrogen`) through the Interactions API to find and repair silent formula errors in a four-sheet Excel workbook. Python runs every tool call, tests each proposed formula, freezes the proposals, and scores them on hidden scenarios the model never sees.

Tested on November 13, 2026 with Python 3.13, `google-genai` 3.3.0, `openpyxl` 3.1.5, and LibreOffice 24.2 on the Gemini Developer API.

## Files

| File | What it does |
| --- | --- |
| `northstar_q4_budget.xlsx` | The working workbook the agent investigates (three planted errors, two valid exceptions) |
| `build_workbook.py` | Rebuilds the working workbook and the evaluator-only reference in `evaluation/` |
| `formula_engine.py` | Restricted formula parser and evaluator (SUM, SUMIFS, IF, MIN, MAX, ROUND, arithmetic, comparisons) |
| `tools.py` | The four tools the model can request |
| `agent.py` | The Interactions API loop, application limits, run log, and freeze step |
| `evaluator.py` | Hidden scenarios, plain-Python reference rules, patch application, workbook diff |
| `app_streamlit.py` | Streamlit interface for one investigation |

## Setup

```bash
git clone https://github.com/KhalidAbdelaty/gemini-4-hydrogen-api.git
cd gemini-4-hydrogen-api
python -m venv .venv
.venv\Scripts\Activate.ps1          # macOS or Linux: source .venv/bin/activate
pip install -r requirements.txt
```

Create a `.env` file with one line, `GEMINI_API_KEY=your-key`. `agent.py` reads only that variable from that file. If your key is on the free tier, set `BILLING_TIER = "free"` in `agent.py`.

Install [LibreOffice](https://www.libreoffice.org/download/) and make `soffice` reachable. On Windows:

```powershell
$env:Path += ";C:\Program Files\LibreOffice\program"
```

## Run

```bash
python agent.py                    # live run: calls the Gemini API, then runs evaluator.py
python evaluator.py runs/<run>.json   # re-score a frozen run, no API call
streamlit run app_streamlit.py     # interface: live run or the latest recorded run
```

`python build_workbook.py` rebuilds the workbook and writes a new `evaluation/manifest.json`. Only run it if you want a fresh baseline; the agent refuses to start if the workbook no longer matches the manifest.

The planted errors are listed in `evaluation/manifest.json` and `build_workbook.py`. The agent's tools never open either file.
