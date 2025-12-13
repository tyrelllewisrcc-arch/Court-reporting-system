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
4. The system will read your data and **fill in the numbers** into your template automatically.
""")

# --- 1. SMART COLUMN MAPPING ---
def smart_read_excel(file):
    """
    Scans an Excel file to find the header row and maps columns to standard names.
    """
    if not file: return None
    
    # 1. Find Header Row (Scan first 20 rows)
    df_preview = pd.read_excel(file, header=None, nrows=20)
    header_row = 0
    found = False
    
    for idx, row in df_preview.iterrows():
        row_str = row.astype(str).str.upper().values
        # Look for key columns to identify the header
        if any(x in row_str for x in ['COURT BOOK', 'CASE NO', 'CASE #']) and \
           any(x in row_str for x in ['CHARGE', 'OFFENCE', 'OFFENSE']):
            header_row = idx
            found = True
            break
            
    if not found:
        st.error(f"Could not find headers in {file.name}. Ensure it has 'Court Book' and 'Charge' columns.")
        return None

    # 2. Read Data
    df = pd.read_excel(file, header=header_row)
    df.columns = df.columns.str.strip().str.upper() # Standardize to UPPERCASE
    
    # 3. Rename Columns to Standard Names
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
        elif 'FURTHER' in c: col_map[c] = 'SENTENCE' # For sentence details
    
    return df.rename(columns=col_map)

# --- 2. CATEGORIZATION ENGINE (Matches your Sheet Structure) ---
def classify_crime(charge, victim):
    charge = str(charge).upper()
    victim = str(victim).upper()
    
    # --- ROW MAPPING FOR SHEET 1 ---
    # We return: (Category Name, Excel Row Number for Sheet 1)
    
    # 1. POLICE RULE
    if any(k in victim for k in ['POLICE', 'PC ', 'CPL ', 'WPC ', 'GOB']) and 'MINOR' not in victim:
        if any(x in charge for x in ['ASSAULT', 'RESIST']): return "AGAINST LAWFUL AUTHORITY", 25 # Others/Assault Police
        return "AGAINST LAWFUL AUTHORITY", 25

    # 2. AGAINST LAWFUL AUTHORITY (Rows 21-25)
    if 'ESCAPE' in charge: return "AGAINST LAWFUL AUTHORITY", 24
    if 'PERJURY' in charge: return "AGAINST LAWFUL AUTHORITY", 23
    if any(x in charge for x in ['DISORDERLY', 'ABUSIVE', 'THREATENING WORDS']): return "AGAINST LAWFUL AUTHORITY", 22
    
    # 3. AGAINST PUBLIC MORALITY (Rows 26-31)
    if 'RAPE' in charge: return "AGAINST PUBLIC MORALITY", 27
    if 'SEXUAL ASSAULT' in charge: return "AGAINST PUBLIC MORALITY", 28
    if 'UNLAWFUL SEXUAL' in charge: return "AGAINST PUBLIC MORALITY", 29
    if 'UNNATURAL' in charge: return "AGAINST PUBLIC MORALITY", 30
    
    # 4. AGAINST THE PERSON (Rows 32-41)
    if 'ATTEMPT' in charge and 'MURDER' in charge: return "AGAINST THE PERSON", 35
    if 'MURDER' in charge: return "AGAINST THE PERSON", 33
    if 'MANSLAUGHTER' in charge: return "AGAINST THE PERSON", 34
    if 'GRIEVOUS' in charge or 'DANGEROUS HARM' in charge: return "AGAINST THE PERSON", 36
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
    if 'VEHICLE' in charge and 'TAKING' in charge: return "OTHERS", 58
    if 'TRAFFIC' in charge or 'MOTOR' in charge or 'LICENSE' in charge: return "OTHERS", 59
    if 'FIREARM' in charge or 'AMMUNITION' in charge: return "OTHERS", 59
    
    return "OTHERS", 59 # Default catch-all

# --- 3. TEMPLATE FILLER LOGIC ---
def fill_template(template_file, monthly_stats, month_col='D'):
    """
    Opens the template and writes counts into specific rows.
    month_col: 'D' for New Cases (Sheet 1), 'J' for Disposed (Sheet 1)
    """
    wb = openpyxl.load_workbook(template_file)
    
    # --- FILL SHEET 1 (Main Stats) ---
    if 'Sheet1' in wb.sheetnames:
        ws = wb['Sheet1']
        
        # Iterate through our calculated stats and write to cells
        # stats format: { Row_Number: Count }
        for row_num, count in monthly_stats.items():
            # Write to the specific cell (e.g., D33 for Murder New Cases)
            try:
                current_val = ws[f"{month_col}{row_num}"].value or 0
                ws[f"{month_col}{row_num}"] = current_val + count
            except:
                pass # Skip if row number is invalid
    
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
        st.error("Please upload BOTH your Data File and the Template File.")
        st.stop()
        
    # 1. Read Data
    df = smart_read_excel(data_file)
    if df is None: st.stop()
    
    # 2. Filter by Date
    date_col = 'DATE_ARR' if mode == "New Cases (Arraignments)" else 'DATE_DISP'
    
    if date_col not in df.columns:
        st.error(f"Could not find a date column for {mode}. Looked for headers like 'Arraignment' or 'Concluded'. Found: {list(df.columns)}")
        st.stop()
        
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    mask = (df[date_col].dt.month == report_month) & (df[date_col].dt.year == report_year)
    df_filtered = df[mask].copy()
    
    st.success(f"Found {len(df_filtered)} cases for {datetime(2025, report_month, 1).strftime('%B')}.")
    
    # 3. Calculate Stats (Map to Row Numbers)
    # Dictionary to store {Row_Index: Count}
    row_counts = {}
    
    for _, row in df_filtered.iterrows():
        cat_name, row_idx = classify_crime(row.get('CHARGE', ''), row.get('VICTIM', ''))
        
        if row_idx in row_counts:
            row_counts[row_idx] += 1
        else:
            row_counts[row_idx] = 1
            
    # 4. Fill Template
    # Column D = New Cases, Column J = Disposed Cases (Based on Sheet 1 structure)
    target_col = 'D' if mode == "New Cases (Arraignments)" else 'J'
    
    try:
        wb_filled = fill_template(template_file, row_counts, month_col=target_col)
        
        # 5. Save & Download
        output = io.BytesIO()
        wb_filled.save(output)
        output.seek(0)
        
        st.write("### ✅ Template Filled Successfully!")
        st.write("The system mapped your cases to the specific rows in Sheet 1.")
        
        st.download_button(
            label="📥 Download Filled Report",
            data=output,
            file_name=f"San_Pedro_Stats_FILLED_{report_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Audit Preview
        st.write("---")
        st.write("### Audit Trail (What was counted)")
        st.dataframe(df_filtered[['CASEID', 'CHARGE', 'VICTIM', 'STATUS']])
        
    except Exception as e:
        st.error(f"Error writing to template: {e}")
