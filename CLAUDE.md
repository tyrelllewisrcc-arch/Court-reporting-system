# CLAUDE.md — Court Reporting System

## Project Overview

**San Pedro Auto-Filler Pro** is a single-page Streamlit web application that automates the population of official 9-sheet court statistics Excel reports for the San Pedro Magistrate's Court in Belize. Users upload a raw court book data file and a blank report template; the app processes the data and produces a filled Excel report ready for submission.

## Repository Structure

```
Court-reporting-system/
├── app.py              # Entire application — UI, logic, and data processing
├── requirements.txt    # Python dependencies
└── README.md           # One-line project description
```

This is an intentionally minimal, single-file application. All business logic lives in `app.py`.

## Technology Stack

| Library      | Purpose                                      |
|--------------|----------------------------------------------|
| `streamlit`  | Web UI framework (widgets, file upload, download) |
| `pandas`     | Reading Excel data files, date filtering     |
| `openpyxl`   | Reading/writing the Excel template workbook  |
| `xlsxwriter` | Listed as dependency; not used directly in code |

Python version: not pinned — any modern Python 3.8+ should work.

## Running the Application

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app runs locally on `http://localhost:8501` by default.

## Application Architecture

The app is structured in four logical sections, marked by comments in `app.py`:

### 1. Smart Column Mapping (`smart_read_excel`) — `app.py:21`

Reads the raw court data Excel file with intelligent header detection:
- Scans the first 20 rows looking for a header row that contains both `COURT BOOK` and either `CHARGE` or `OFFENCE`
- Normalises all column names to uppercase
- Maps raw column names to canonical internal names:

| Raw Column Contains        | Internal Name  |
|----------------------------|----------------|
| `COURT BOOK`               | `CASEID`       |
| `CHARGE` / `OFFENCE`       | `CHARGE`       |
| `COMPLAINANT` / `VICTIM`   | `VICTIM`       |
| `ARRAINGMENT` / `ARRAIGNMENT` | `DATE_ARR`  |
| `CONCLUDED` / `DISPOSAL`   | `DATE_DISP`    |
| `AGE`                      | `AGE`          |
| `SEX` / `GENDER`           | `GENDER`       |
| `FURTHER`                  | `SENTENCE`     |
| `STATUS`                   | `CASE_STATUS`  |
| `REMARK`                   | `REMARK`       |

### 2. Intelligent Parsers — `app.py:56`

Four pure functions that classify text values into categories:

- **`classify_crime_sheet1(charge, victim)`** — Returns the Excel row number (integer) for the crime category in Sheets 1/2/3/4/5. Uses keyword matching on the charge string and victim string. Police-related victims map to row 25. Falls back to row 59 for unrecognised charges.

- **`classify_statutory_sheet8(charge)`** — Returns the row number for Sheets 8/9. Maps charge keywords to statutory categories: Dangerous Drugs (12), Firearms (13), Liquor (14), Police Act (15), Gambling (16), Traffic/Motor/Licence (17), Other (18).

- **`parse_disposition(remark)`** — Returns one of `CONVICTED`, `DISMISSED`, `NOLLE`, or `OTHER` based on keywords in the remark field.

- **`parse_sentence(sentence_text)`** — Returns one of `FINE`, `PRISON`, `PROBATION`, `REFORMATORY`, or `OTHER`.

- **`is_juvenile(age)`** — Returns `True` if age ≤ 16.

- **`get_age_col_sheet5(age, gender)`** — Returns the Excel column letter (`B`–`K`) for Sheet 5 age/gender demographic breakdown.

### 3. Template Filler (`fill_all_sheets`) — `app.py:140`

Core function that takes an `openpyxl` workbook and a filtered DataFrame, then populates cells by incrementing counts. The function respects the `mode` parameter:

**Mode: `"New"` (Arraignments)**
- Fills Sheet 1 column `D` with unique case counts by crime row
- Fills Sheet 3 column `D` with person counts by crime row
- Fills Sheet 8 column `C` with new statutory case counts

**Mode: `"Disposed"` (Concluded Cases)**
- Fills Sheet 1 column `J` and Sheet 3 column `J` (disposed counts)
- Fills Sheet 2 with disposal breakdown by outcome (Dismissed→`C`, Nolle→`D`, Convicted→`E`)
- For **convicted** cases only:
  - Sheet 4: sentence type by gender (columns `D`–`I`)
  - Sheet 5: age/gender demographic (columns `B`–`K`, row offset: `r_num - 11`)
  - Sheets 6 & 7: juvenile cases (row offset: `r_num - 14`)
  - Sheet 9: statutory punishment by gender
- Fills Sheet 8 columns `E`/`F` with convicted/dismissed statutory counts

### 4. Main Interface — `app.py:265`

Streamlit sidebar controls:
- File uploader for raw data Excel file
- File uploader for blank template Excel file
- Radio: **New Cases (Arraignments)** or **Disposed Cases (Concluded)**
- Checkbox: **Full Year Report** (if unchecked, shows month selector)
- Year number input (default 2025)

On button click, the app:
1. Parses the data file via `smart_read_excel`
2. Filters rows by the selected month/year using `DATE_ARR` (New mode) or `DATE_DISP` (Disposed mode)
3. Calls `fill_all_sheets` with the filtered DataFrame
4. Offers a `st.download_button` for the filled `.xlsx` output

## Key Conventions and Quirks

### Row Number Mapping
All sheet-filling uses **direct Excel row numbers as integers**. The `classify_*` functions return row numbers that correspond to specific crime/offence categories in the template. These row numbers are tightly coupled to the physical layout of the Excel template. If the template layout changes, all `classify_*` return values must be updated.

### Deduplication Logic (Sheet 1 vs Sheet 3)
- **Sheet 1** counts unique **cases** (deduplicated by `CASEID`)
- **Sheet 3** counts **persons** (every row, including co-accused sharing a `CASEID`)

### Gender Detection
Gender is detected by the absence of `'F'` in the gender string (`'F' not in str(gender).upper()`). This means any value not containing `F` is treated as male. Ensure source data uses `M`/`F` or `Male`/`Female` consistently.

### Age Offsets
- Sheet 5 uses row offset: `r_num - 11`
- Sheets 6/7 use row offset: `r_num - 14`

These hardcoded offsets assume specific template structure. They are approximate and may need adjustment if the template is revised.

### Silent Failures
The `fill_all_sheets` function wraps all cell writes in bare `except: pass` clauses. Invalid row/column references are silently ignored. Always verify output against source data when debugging discrepancies.

### Date Column Requirement
The app enforces that the appropriate date column exists:
- New cases mode requires `DATE_ARR` (Arraignment date)
- Disposed cases mode requires `DATE_DISP` (Disposal/Conclusion date)

If the column is missing, the app halts with an error before processing.

### Month Filter Hard-coding
The period name display uses `datetime(2025, report_month, 1)` regardless of the `report_year` input (`app.py:298`). This is a display-only bug; the actual data filter uses `report_year` correctly.

## Data Flow Summary

```
User uploads:
  ├── Data file (court book Excel)      → smart_read_excel() → normalised DataFrame
  └── Template (blank 9-sheet Excel)   → openpyxl.load_workbook()

User selects: mode (New/Disposed), period (month+year or full year)

DataFrame filtered by date column and period
          ↓
fill_all_sheets(template_wb, filtered_df, mode)
  ├── classify_crime_sheet1()   → row numbers for Sheets 1,2,3,4,5,6,7
  ├── classify_statutory_sheet8() → row numbers for Sheets 8,9
  ├── parse_disposition()       → CONVICTED / DISMISSED / NOLLE / OTHER
  ├── parse_sentence()          → FINE / PRISON / PROBATION / REFORMATORY / OTHER
  ├── is_juvenile()             → routes to Sheets 6,7
  └── get_age_col_sheet5()      → column letter for Sheet 5
          ↓
Filled workbook saved to BytesIO → st.download_button
```

## Development Guidelines

### Making Changes to Crime Classification
When adding new charge keywords to `classify_crime_sheet1` or `classify_statutory_sheet8`:
- Add the new keyword check **before** the generic fallback (`return 59` / `return 18`)
- Verify the target row number against the physical Excel template
- More specific patterns should appear earlier in the if-chain to avoid being shadowed by broader patterns (e.g., check `'ATTEMPT' in charge and 'MURDER' in charge` before `'MURDER' in charge`)

### Adding New Sheet Support
1. Add the sheet name check pattern: `if 'SheetN' in wb.sheetnames:`
2. Follow the existing `(ws[f"{col}{row}"].value or 0) + 1` increment pattern
3. Handle both `"New"` and `"Disposed"` modes as appropriate

### Modifying Column Mappings
Edit the `col_map` dictionary inside `smart_read_excel`. Keys are substrings matched against uppercased column names; values are the canonical internal column names used throughout the rest of the code.

### Testing Locally
There is no automated test suite. Manual testing requires:
1. A sample court book Excel file with a recognisable header row
2. A blank 9-sheet template Excel file matching the expected layout
3. Running `streamlit run app.py` and exercising both New and Disposed modes

### Dependencies
Do not remove `xlsxwriter` from `requirements.txt` even though it is not directly imported in `app.py` — it may be required by `pandas` for Excel export operations in future enhancements.
