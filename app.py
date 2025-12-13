import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

# --- PAGE CONFIG ---
st.set_page_config(page_title="San Pedro Auto-Filler", layout="wide")
st.title("📝 Auto-Fill 'San Pedro Annual Statistics' Report")
st.markdown("""
**Instructions:**
1. Upload your **Data File** (Active Cases or Returns).
2. Upload your **Blank Report Template** (`_San Pedro Annual Statistics 2025_.xlsx`).
3. Select the **Month**.
4. The system will **Auto-Fill Sheet 1 (Cases) and Sheet 3 (Persons)**.
""")

# --- 1. SMART COLUMN MAPPING (FIXED) ---
def smart_read_excel(file):
    """
    Scans an Excel file, finds the header, and strictly safely renames columns.
    """
    if not file: return None
    
    # 1. Find Header Row
    df_preview = pd.read_excel(file, header=None, nrows=20)
    header_row = 0
    found = False
    
    for idx, row in df_preview.iterrows():
        # Convert row to string safely to search keywords
        row_text = " ".join([str(x).upper() for x in row.values])
        
        has_id = any(x in row_text for x in ['COURT BOOK', 'CASE NO', 'CASE #'])
        has_charge = any(x in row_text for x in ['CHARGE', 'OFFENCE'])
        
        if has_id and has_charge:
            header_row = idx
            found = True
            break
            
    if not found:
        st.error(f"Could not find headers in {file.name}. Ensure it has 'Court Book' and 'Charge' columns.")
        return None

    # 2. Read Data
    df = pd.read_excel(file, header=header_row)
    
    # 3. Clean Column Names (Fixes the TypeError)
    # Convert all columns to String, Strip Whitespace, Uppercase
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    # 4. Map to Standard Internal Names
    col_map = {}
    for c in df.columns:
        if 'COURT BOOK' in c: col_map[c] = 'CASEID'
        elif 'CHARGE' in c or 'OFFENCE' in c: col_map[c] = 'CHARGE'
        elif 'COMPLAINANT' in c or 'VICTIM' in c: col_map[c] = 'VICTIM'
        elif 'STATUS' in c or 'REMARK' in c: col_map[c] = 'STATUS'
        elif 'ARRAINGMENT' in c or 'ARRAIGNMENT' in c: col_map[c] = 'DATE_ARR'
        elif 'CONCLUDED' in c or 'DISPOSAL' in c: col_map[c] = 'DATE_DISP'
        elif 'AGE' in c: col_map[c] = 'AGE'
        elif 'SEX' in c or 'GENDER' in c: col_map[c] = 'GENDER'
        elif 'DEFENDANT' in c: col_map[c] = 'DEFENDANT'
        elif 'FURTHER' in c: col_map[c] = 'SENTENCE'
    
    return df.rename(columns=col_map)

# --- 2. CATEGORIZATION ENGINE ---
def classify_crime(charge, victim):
    charge = str(charge).upper()
    victim = str(victim).upper()
    
    # Returns: (Category Name, Row Number)
    
    # POLICE RULE
    if any(k in victim for k in ['POLICE', 'PC ', 'CPL ', 'WPC ', 'GOB']) and 'MINOR' not in victim:
        if any(x in charge for x in ['ASSAULT', 'RESIST']): return "AGAINST LAWFUL AUTHORITY", 25 
        return "AGAINST LAWFUL AUTHORITY", 25

    # 2. AGAINST LAWFUL AUTHORITY (Rows 21-25)
    if 'ESCAPE' in charge: return "AGAINST LAWFUL AUTHORITY", 24
    if 'PERJURY' in charge: return "AGAINST LAWFUL AUTHORITY", 23
    if any(x in charge for x in ['DISORDERLY', 'ABUSIVE', 'THREAT']): return "AGAINST LAWFUL AUTHORITY", 22
    
    # 3. AGAINST PUBLIC MORALITY (Rows 26-31)
    if 'RAPE' in charge: return "AGAINST PUBLIC MORALITY", 27
    if 'SEXUAL ASSAULT' in charge: return "AGAINST PUBLIC MORALITY", 28
    if 'UNLAWFUL SEXUAL' in charge: return "AGAINST PUBLIC MORALITY", 29
    if 'UNNATURAL' in charge: return "AGAINST PUBLIC MORALITY", 30
    
    # 4. AGAINST THE PERSON (Rows 32-41)
    if 'ATTEMPT' in charge and 'MURDER' in charge: return "AGAINST THE PERSON", 35
    if 'MURDER' in charge: return "AGAINST THE PERSON", 33
    if 'MANSLAUGHTER' in charge: return "AGAINST THE PERSON", 34
    if 'GRIEVOUS' in charge: return "AGAINST THE PERSON", 36
    if 'WOUNDING' in charge: return "AGAINST THE PERSON", 37
    if 'HARM' in charge: return "AGAINST THE PERSON", 38
    if 'AGGRAVATED ASSAULT' in charge: return "AGAINST THE PERSON", 39
    if 'COMMON ASSAULT' in charge: return "AGAINST THE PERSON", 40
    
    # 5. AGAINST PROPERTY (Rows 42-50)
    if 'ROBBERY' in charge: return "AGAINST PROPERTY", 43
    if 'BURGLARY' in charge: return "AGAINST PROPERTY", 44
    if 'THEFT' in charge: return "AGAINST PROPERTY", 45
    if 'DECEPTION' in charge or 'FRAUD' in charge: return "AGAINST PROPERTY", 46
    if 'HANDLING' in charge: return "AGAINST PROPERTY", 47
    if 'DAMAGE' in charge: return "AGAINST PROPERTY", 48
    if 'ARSON' in charge: return "AGAINST PROPERTY", 49
    
    # 6. OTHERS (Rows 51-59)
    if 'FORGERY' in charge: return "OTHERS", 52
    if 'DRUG' in charge or 'CANNABIS' in charge: return "OTHERS", 54
    if 'PIPE' in charge: return "OTHERS", 57
    if 'VEHICLE' in charge: return "OTHERS", 58
    if 'TRAFFIC' in charge or 'MOTOR' in charge or 'LICENSE' in charge: return "OTHERS", 59
    if 'FIREARM' in charge or 'AMMUNITION' in charge: return "OTHERS", 59
    
    return "OTHERS", 59 

# --- 3. TEMPLATE FILLER LOGIC ---
def fill_template(template_file, case_counts, person_counts, month_col='D'):
    wb = openpyxl.load_workbook(template_file)
    
    # --- FILL SHEET 1 (CASES) ---
    if 'Sheet1' in wb.sheetnames:
        ws = wb['Sheet1']
        for row_num, count in case_counts.items():
            try:
                current = ws[f"{month_col}{row_num}"].value
                if not isinstance(current, (int, float)): current = 0
                ws[f"{month_col}{row_num}"] = current + count
            except: pass
            
    # --- FILL SHEET 3 (PERSONS) ---
    # Sheet 3 usually has same structure as Sheet 1. 
    # If Sheet3 exists, we assume same row mappings.
    if 'Sheet3' in wb.sheetnames:
        ws3 = wb['Sheet3']
        for row_num, count in person_counts.items():
            try:
                current = ws3[f"{month_col}{row_num}"].value
                if not isinstance(current, (int, float)): current = 0
                ws3[f"{month_col}{row_num}"] = current + count
            except: pass

    return wb

# --- 4. MAIN APP ---
st.sidebar.header("1. Upload Files")
data_file = st.sidebar.file_uploader("Upload Data File (Excel)", type=['xlsx'])
template_file = st.sidebar.file_uploader("Upload 'San Pedro Statistics' Template", type=['xlsx'])

st.sidebar.header("2. Settings")
mode = st.sidebar.radio("What data is this?", ["New Cases (Arraignments)", "Disposed Cases (Concluded)"])
report_month = st.sidebar.selectbox("Select Month", range(1, 13), format_func=lambda x: datetime(2025, x, 1).strftime('%B'))
report_year = st.sidebar.number_input("Year", value=2025)

if st.button("Run Auto-Fill"):
    if not data_file or not template_file:
        st.error("Please upload BOTH files.")
        st.stop()
        
    # 1. Read Data
    df = smart_read_excel(data_file)
    if df is None: st.stop()
    
    # 2. Filter by Date
    date_col = 'DATE_ARR' if mode == "New Cases (Arraignments)" else 'DATE_DISP'
    
    if date_col not in df.columns:
        st.error(f"Could not find date column. Ensure '{'Arraignment' if mode.startswith('New') else 'Concluded'}' column exists.")
        st.stop()
        
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    mask = (df[date_col].dt.month == report_month) & (df[date_col].dt.year == report_year)
    df_filtered = df[mask].copy()
    
    st.success(f"Processing {len(df_filtered)} rows for {datetime(2025, report_month, 1).strftime('%B')}.")
    
    # 3. Calculate Stats
    
    # A. Person Counts (Raw Rows)
    person_counts = {}
    for _, row in df_filtered.iterrows():
        _, row_idx = classify_crime(row.get('CHARGE', ''), row.get('VICTIM', ''))
        person_counts[row_idx] = person_counts.get(row_idx, 0) + 1
        
    # B. Case Counts (Unique Court Book #)
    case_counts = {}
    # Drop duplicates so 1 Case ID = 1 Count
    if 'CASEID' in df_filtered.columns:
        df_unique_cases = df_filtered.drop_duplicates(subset=['CASEID'])
        for _, row in df_unique_cases.iterrows():
            _, row_idx = classify_crime(row.get('CHARGE', ''), row.get('VICTIM', ''))
            case_counts[row_idx] = case_counts.get(row_idx, 0) + 1
    else:
        st.warning("No 'Court Book No.' found. Counting rows as cases.")
        case_counts = person_counts

    # 4. Fill Template
    target_col = 'D' if mode == "New Cases (Arraignments)" else 'J'
    
    try:
        wb_filled = fill_template(template_file, case_counts, person_counts, month_col=target_col)
        
        output = io.BytesIO()
        wb_filled.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download Filled Report",
            data=output,
            file_name=f"San_Pedro_Stats_FILLED_{report_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # 5. Audit Trail
        st.write("---")
        st.write("### 🔍 Audit Trail (Data Extracted)")
        
        # Display key columns including Age/Sex if found
        display_cols = ['CASEID', 'DEFENDANT', 'AGE', 'GENDER', 'CHARGE', 'STATUS', date_col]
        # Filter to show only cols that actually exist in df
        final_cols = [c for c in display_cols if c in df_filtered.columns]
        
        st.dataframe(df_filtered[final_cols])
        
    except Exception as e:
        st.error(f"Error processing template: {e}")
