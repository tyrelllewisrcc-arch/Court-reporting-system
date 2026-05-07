# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A single-file Streamlit app (`app.py`) that automates filling a 9-sheet Excel court statistics report for San Pedro Magistrate Court. Users upload a raw court data workbook plus a blank template; the app classifies each case by charge type and disposition, then writes tallies into the correct cells of the template.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

There are no tests, no linter config, and no build step.

## Architecture

The entire application lives in `app.py` (~320 lines), structured in four sections:

### 1. Input parsing — `smart_read_excel(file)`
Scans the first 20 rows of the uploaded data file to find the header row (looks for `COURT BOOK` + `CHARGE`/`OFFENCE` in the same row). Normalises all matched columns to canonical names: `CASEID`, `CHARGE`, `VICTIM`, `DATE_ARR`, `DATE_DISP`, `AGE`, `GENDER`, `SENTENCE`, `CASE_STATUS`, `REMARK`. All downstream code uses only these canonical names.

### 2. Classifiers (pure functions, return hardcoded row/column positions)

| Function | Purpose | Returns |
|---|---|---|
| `classify_crime_sheet1(charge, victim)` | Maps a charge string to a template row number (22–59). Police-as-victim is a special case → row 25. | int |
| `classify_statutory_sheet8(charge)` | Maps to statutory row 12–18 for Sheets 8 & 9. | int |
| `parse_disposition(remark)` | `CONVICTED` / `DISMISSED` / `NOLLE` / `OTHER` | str |
| `parse_sentence(sentence_text)` | `FINE` / `PRISON` / `PROBATION` / `REFORMATORY` / `OTHER` | str |
| `is_juvenile(age)` | `True` if age ≤ 16 | bool |
| `get_age_col_sheet5(age, gender)` | Returns Excel column letter B–K for the age/gender demographic cell. | str or None |

### 3. Template filler — `fill_all_sheets(template_file, df, mode)`
Loads the blank template with `openpyxl`, increments cells by reference (e.g. `ws["D25"]`), and returns the modified workbook. The `mode` argument is either `"New"` or `"Disposed"`:

- **New**: writes to Sheets 1, 3 (column D) and Sheet 8 (column C).
- **Disposed**: writes to Sheets 1, 3 (column J), plus Sheets 2, 4, 5, 6, 7, 8, and 9.

**Critical hardcoded offsets** — these must stay in sync with the physical Excel template:
- Sheet 1/3: row numbers come directly from `classify_crime_sheet1`.
- Sheet 5: uses `r_num - 11` to convert from the Sheet 1 row space to the Sheet 5 row space.
- Sheets 6/7 (juveniles): uses `r_num - 14`.

Sheet names in the template must be exactly `Sheet1` through `Sheet9`.

### 4. Streamlit UI (bottom of file)
Sidebar controls: data file upload, template upload, mode radio (`New Cases` / `Disposed Cases`), month/year selectors, full-year checkbox. The main "Process & Fill Report" button drives the pipeline and exposes a download button for the filled workbook.

## Key Constraints

- **Column D vs J**: `D` = new/arraigned cases, `J` = disposed cases in Sheets 1 & 3. Swapping these produces silent data corruption.
- **Row number derivation**: all row numbers are derived from `classify_crime_sheet1`, even for sheets that logically cover different content. Adding a new crime category requires updating that one function plus checking the `-11` and `-14` offsets still hold in the physical template.
- **No deduplication on persons**: Sheet 3 counts each defendant row; Sheet 1 deduplicates by `CASEID` to count unique cases.
- **`xlsxwriter` is listed in requirements but not used** — `openpyxl` handles all writes. Do not replace `openpyxl` with `xlsxwriter` without a full rewrite of the filler.
