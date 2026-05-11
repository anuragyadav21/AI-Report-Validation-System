# AI Report Validation System

**Workbook pipeline:** prompt comparison, LLM reports, rubric validation, and statistics.

This repository implements the **AI Report Validation System** — a **systems-engineering prompt experiment**: three prompt variants (**Prompt A**, **Prompt B**, **Prompt C**) evaluated on the **same scenarios**, with **LLM-generated reports**, **structured rubric scoring (D1–D6)**, and **statistical comparison** of overall scores (repeated-measures design).

### Documentation (validation, experiment, statistics, usage)

For assignment-style documentation — **criteria table**, **design**, **statistical tests**, **system architecture**, **technical details**, and **step-by-step usage** — see **[docs/validation-system.md](docs/validation-system.md)**.

---

## What you get

| Layer | Description |
|--------|-------------|
| **Workbook** | Excel (`.xlsx`) with three prompt sheets plus a **Rubric** reference sheet; columns **A–D** fixed (metadata, scenario, composed prompt, LLM response); **E–L** rubric scores, overall, notes. |
| **Generation** | Optional OpenAI calls to fill **column D** (responses) per sheet/row. |
| **Validation** | OpenAI “judge” returns **JSON** scores for six dimensions; script writes **E–J** and **L** (preserves **D** unless you force overwrite). |
| **Simulation** | Optional script copies a workbook and applies controlled perturbations (down-score rubric integers, synthetic **Overall** in **K**) for method demos. |
| **Analysis** | Descriptive stats, **figures** (boxplot, RM line plot, mean ± 95% CI), **repeated-measures ANOVA**, partial η², **Holm** pairwise tests, **cluster-robust OLS** regression → **Word report** (`.docx`). |

---

## Repository layout

```
project/                         # repository root (your folder name may differ)
├── README.md                      # This file
├── docs/
│   └── validation-system.md       # Validation criteria, experiment design, stats, usage (detailed)
├── .env.example                   # Template for OPENAI_API_KEY (copy to `.env`; `.env` is gitignored)
├── requirements-analysis.txt      # All deps including openai + python-dotenv (see file)
├── prompt-worksheet.xlsx          # Typical main workbook (template / filled data; not always committed)
├── prompt-worksheet_AB.xlsx       # Default input for significance report (example: A+B synthetic edits)
├── scripts/
│   ├── experiment_constants.py    # Column indices, row ranges, sheet names, rubric column headers
│   ├── rubric_sheet.py            # Builds the Excel **Rubric** sheet (anchors + D1–D6 blurbs)
│   ├── workbook_layout.py         # Inserts/syncs rubric columns (E–L), data validation, formulas
│   ├── build_prompt_worksheet.py  # Create template or sync layout on existing file
│   ├── migrate_add_rubric_columns.py   # One-off migrate older workbook → add E–L + Rubric
│   ├── openai_env.py                   # Load `.env` + resolve OPENAI_API_KEY (shared by OpenAI scripts)
│   ├── openai_fill_llm_responses.py    # Fill column D via OpenAI
│   ├── openai_validate_rubric.py       # Score D vs B; write E–J, L (JSON rubric in script)
│   ├── clone_prompt_scores_workbook.py # Copy workbook; optional −2 on D1–D6; random K per sheet
│   ├── analyze_rm_anova.py        # RM ANOVA + Friedman (CLI); exports long CSV; `build_long_frame`
│   └── analyze_prompt_significance.py  # Full homework report → .docx (+ optional stdout / CSV)
└── .mpl_cache/                    # Matplotlib config/cache (created at runtime; gitignored)
```

**Rubric source of truth (two places — keep aligned):**

1. **`scripts/openai_validate_rubric.py`** — string **`SCORING_RUBRIC_BLOCK`**: text actually sent to the API.  
2. **`scripts/rubric_sheet.py`** — **`populate_rubric_sheet()`**: human-readable **Rubric** tab in the workbook.

---

## Workbook column model

| Columns | Content |
|---------|---------|
| **A** | Serial |
| **B** | Scenario text |
| **C** | Prompt + scenario (composed) |
| **D** | LLM response |
| **E–J** | D1–D6 (1–5 Likert-style integers) |
| **K** | Overall (typically `=AVERAGE(E:J)`; may be replaced by a numeric cell for analysis) |
| **L** | Notes / validator comments |

Sheets: **`Prompt A`**, **`Prompt B`**, **`Prompt C`**, **`Rubric`**.  
Row range for data rows is defined in **`experiment_constants.py`** (`FIRST_DATA_ROW` … `LAST_DATA_ROW`).

**Analysis rule for “Overall”:** `build_long_frame()` in `analyze_rm_anova.py` uses **numeric column K** when present; otherwise it falls back to the mean of **E–J**. If **K** is a formula, open the file in Excel and **save** once so cached values exist for `openpyxl` (`data_only=True`).

---

## Setup

### 1. Python

Use Python **3.10+** (3.11/3.12 recommended). A virtual environment is recommended.

```bash
cd "/path/to/ai-report-validation-system"   # or your clone path
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
```

### 2. Dependencies

```bash
pip install -r requirements-analysis.txt
```

Includes **pandas**, **openpyxl**, **pingouin**, **scipy**, **statsmodels**, **python-docx**, **matplotlib**, **openai**, and **python-dotenv**.

### 3. OpenAI API key (`.env`)

Scripts read **`OPENAI_API_KEY`** from a **`.env`** file at the project root (same folder as `README.md`). You do **not** need to `export` the key in the terminal for normal runs.

```bash
cp .env.example .env
# Edit .env and set: OPENAI_API_KEY=sk-...
```

- **`.env`** is listed in **`.gitignore`** — never commit real keys.  
- If `OPENAI_API_KEY` is already set in your shell (e.g. CI), that value is kept; `.env` only fills missing variables unless you change `override` in `scripts/openai_env.py`.

---

## Typical workflows

### A. Create or refresh the workbook template

```bash
python3 scripts/build_prompt_worksheet.py --help
# New file if missing:
python3 scripts/build_prompt_worksheet.py --output prompt-worksheet.xlsx
# Sync rubric columns / Rubric sheet without wiping A–D:
python3 scripts/build_prompt_worksheet.py --output prompt-worksheet.xlsx
# DANGER: full blank rebuild (wipes workbook):
python3 scripts/build_prompt_worksheet.py --output prompt-worksheet.xlsx --force-rebuild
```

### B. Migrate an old workbook (add E–L + Rubric)

```bash
python3 scripts/migrate_add_rubric_columns.py --input prompt-worksheet.xlsx
# Writes prompt-worksheet_with_rubric.xlsx by default, or use --inplace / --output ...
```

### C. Fill LLM responses (column D)

```bash
python3 scripts/openai_fill_llm_responses.py --workbook prompt-worksheet.xlsx
# Optional: --overwrite, --model gpt-4o-mini, --dry-run
```

### D. Validate / score with the rubric (E–J, L)

```bash
python3 scripts/openai_validate_rubric.py --workbook prompt-worksheet.xlsx
# Optional: --overwrite-scores, --dry-run
```

### E. Clone workbook with synthetic score tweaks (optional lab demo)

```bash
python3 scripts/clone_prompt_scores_workbook.py --help
# Default: copy from prompt-worksheet.xlsx → prompt-worksheet_promptA_minus2.xlsx
# With --prompt-b: also adjust Prompt B; random Overall ranges configurable
```

### F. Statistical report (Word + figures)

```bash
python3 scripts/analyze_prompt_significance.py --workbook prompt-worksheet_AB.xlsx
# Creates: prompt-worksheet_AB_significance_report.docx next to the workbook

# Terminal-only (no figures section; no matplotlib temp plots):
python3 scripts/analyze_prompt_significance.py --stdout-only --workbook prompt-worksheet_AB.xlsx

# Export long table for external tools:
python3 scripts/analyze_prompt_significance.py --workbook prompt-worksheet_AB.xlsx --csv-out experiment_long.csv
```

### G. RM ANOVA only (terminal)

```bash
python3 scripts/analyze_rm_anova.py --workbook prompt-worksheet.xlsx
python3 scripts/analyze_rm_anova.py --workbook prompt-worksheet.xlsx --csv-out experiment_long.csv
```

---

## Statistics implemented (`analyze_prompt_significance.py`)

- **Descriptive:** mean and SD of **Overall** by prompt (complete cases: all three prompts non-missing per scenario).  
- **Figures (embedded in .docx):** boxplot by prompt; **repeated-measures** line plot (one line per scenario); mean ± **95% CI**.  
- **Omnibus:** one-way **repeated-measures ANOVA** (Pingouin; Greenhouse–Geisser / sphericity as applicable).  
- **Effect size:** partial η² from SS.  
- **Post-hoc:** paired **t** tests with **Holm** adjustment across three pairwise contrasts.  
- **Regression:** OLS **Overall ~ Prompt** (reference A) with **cluster-robust** standard errors (cluster = scenario).

---

## Design notes & pitfalls

1. **Rubric drift:** If you edit dimension text, update both `openai_validate_rubric.py` and `rubric_sheet.py`.  
2. **Formula cache:** Analysis reads cached cell values; re-save in Excel if **K** looks empty in Python.  
3. **Git:** Add `.env`, API keys, and large private datasets to **`.gitignore`** (already ignores `.venv/`, `__pycache__/`, `.mpl_cache/`, Office lock files `~$*`).  
4. **Office lock files:** Close Excel/Word before scripts overwrite the same `.xlsx` / `.docx`.

---

## License / course use

Adapt this README if your institution requires a license block or academic integrity statement. Course-specific rubrics belong in your own submission; this repo holds the **implementation** only.
