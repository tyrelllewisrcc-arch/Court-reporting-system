import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- CONFIGURATION ---
st.set_page_config(page_title="Belize Court Stats System", layout="wide")
st.title("🇧🇿 Flexible Court Statistical Reporting System")

# --- 1. SMART HEADER DETECTION ---
def load_data_flexibly(uploaded_file):
    """
    Reads an Excel file and automatically finds the header row 
    by looking for key columns like 'Court Book' or 'Charge'.
    """
    if not uploaded_file:
        return None
    
    # Read first 10 rows to scan for headers
    df_preview = pd.read_excel(uploaded_file, header=None, nrows=10)
    
    header_row_index = None
    for idx, row in df_preview.iterrows():
        row_str = row.astype(str).str.upper().values
        # Key words to identify the header row
        if any("COURT BOOK" in x for x in row_str) and any("CHARGE" in x for x in row_str):
            header_row_index = idx
            break
    
    if header_row_index is None:
        st.error(f"Could not find headers in {uploaded_file.name}. Please ensure columns 'Court Book No.' and 'Charge' exist.")
        return None
    
    # Reload with correct header
    df = pd.read_excel(uploaded_file, header=header_row_index)
    
    # Normalize Columns (Strip spaces, handle slight naming variations)
    df.columns = df.columns.str.strip()
    
    # Map varied column names to standard internal names
    col_map = {}
    for c in df.columns:
        cu = c.upper()
        if "COURT BOOK" in cu: col_map[c] = "CaseID"
        elif "ARRAINGMENT" in cu or "ARRAIGNMENT" in cu: col_map[c] = "ArraignmentDate"
        elif "CONCLUDED" in cu: col_map[c] = "DisposalDate"
        elif "COMPLAINANT" in cu or "VICTIM" in cu: col_map[c] = "Victim"
        elif "CHARGE" in cu: col_map[c] = "Charge"
        elif "REMARK" in cu and "FURTHER" not in cu: col_map[c] = "Remark"
        elif "SEX" in cu or "GENDER" in cu: col_map[c] = "Gender"
        elif "AGE" in cu: col_map[c] = "Age"
    
    df = df.rename(columns=col_map)
    return df

# --- 2. CATEGORIZATION ENGINE (UNCHANGED) ---
def get_category(charge, complainant):
    charge = str(charge).upper()
    complainant = str(complainant).upper()

    # POLICE RULE
    police_keywords = ['POLICE', 'PC ', 'WPC ', 'CPL ', 'SGT ', 'GOB', 'DEPARTMENT']
    if any(k in complainant for k in police_keywords) and "MINOR" not in complainant:
        return "AGAINST LAWFUL AUTHORITY"

    # MAPPING
    if any(x in charge for x in ['ESCAPE', 'RESCUE', 'PUBLIC TERROR', 'DISORDERLY', 'ABUSIVE', 'PERJURY', 'RESIST', 'OBSTRUCT']):
        return "AGAINST LAWFUL AUTHORITY"
    if any(x in charge for x in ['RAPE', 'UNLAWFUL SEXUAL', 'SEXUAL ASSAULT', 'UNNATURAL']):
        return "AGAINST PUBLIC MORALITY"
    if any(x in charge for x in ['MURDER', 'MANSLAUGHTER', 'HARM', 'WOUNDING', 'ASSAULT', 'THREAT']):
        return "AGAINST THE PERSON"
    if any(x in charge for x in ['ROBBERY', 'BURGLARY', 'THEFT', 'DECEPTION', 'FRAUD', 'HANDLING', 'DAMAGE', 'ARSON']):
        return "AGAINST PROPERTY"
    if any(x in charge for x in ['DRUG', 'CANNABIS', 'COCAINE', 'PIPE', 'FIREARM', 'AMMUNITION', 'TRAFFIC', 'MOTOR', 'LICENSE']):
        return "OTHERS"
    
    return "OTHERS" # Default

# --- 3. MAIN LOGIC ---

st.sidebar.header("Settings")
report_month = st.sidebar.selectbox("Select Report Month", range(1, 13), format_func=lambda x: datetime(2025, x, 1).strftime('%B'))
report_year = st.sidebar.number_input("Year", value=2025)

st.sidebar.header("Upload Files (Either or Both)")
file1 = st.sidebar.file_uploader("Upload File A (Active or Returns)", type=['xlsx'], key="f1")
file2 = st.sidebar.file_uploader("Upload File B (Optional)", type=['xlsx'], key="f2")

if st.button("Process Data"):
    if not file1 and not file2:
        st.warning("Please upload at least one file.")
        st.stop()

    # Consolidate uploads into one list
    files_to_process = [f for f in [file1, file2] if f is not None]
    
    all_new_cases = []
    all_disposed_cases = []
    all_audit_data = []

    for f in files_to_process:
        df = load_data_flexibly(f)
        if df is None: continue

        # --- A. COUNT NEW CASES (Arraignment Date matches Report Month) ---
        if 'ArraignmentDate' in df.columns:
            df['ArraignmentDate'] = pd.to_datetime(df['ArraignmentDate'], errors='coerce')
            
            mask_new = (df['ArraignmentDate'].dt.month == report_month) & \
                       (df['ArraignmentDate'].dt.year == report_year)
            
            df_new = df[mask_new].copy()
            
            for _, row in df_new.iterrows():
                cat = get_category(row.get('Charge', ''), row.get('Victim', ''))
                all_new_cases.append({
                    'Category': cat,
                    'CaseID': row.get('CaseID'),
                    'Charge': row.get('Charge'),
                    'Date': row['ArraignmentDate']
                })
                # Add to detailed audit list
                all_audit_data.append({
                    'Source File': f.name,
                    'Type': 'New Case (Arraigned)',
                    'CaseID': row.get('CaseID'),
                    'Category': cat,
                    'Charge': row.get('Charge'),
                    'Date': row['ArraignmentDate'].strftime('%Y-%m-%d'),
                    'Status': row.get('Status', '')
                })

        # --- B. COUNT DISPOSED CASES (Disposal Date matches Report Month) ---
        if 'DisposalDate' in df.columns:
            df['DisposalDate'] = pd.to_datetime(df['DisposalDate'], errors='coerce')
            
            mask_disposed = (df['DisposalDate'].dt.month == report_month) & \
                            (df['DisposalDate'].dt.year == report_year)
            
            df_disp = df[mask_disposed].copy()
            
            for _, row in df_disp.iterrows():
                cat = get_category(row.get('Charge', ''), row.get('Victim', ''))
                all_disposed_cases.append({
                    'Category': cat,
                    'CaseID': row.get('CaseID'),
                    'Charge': row.get('Charge'),
                    'Date': row['DisposalDate']
                })
                # Add to detailed audit list
                all_audit_data.append({
                    'Source File': f.name,
                    'Type': 'Disposed Case',
                    'CaseID': row.get('CaseID'),
                    'Category': cat,
                    'Charge': row.get('Charge'),
                    'Date': row['DisposalDate'].strftime('%Y-%m-%d'),
                    'Status': row.get('Remark', 'Concluded')
                })

    # --- C. AGGREGATE STATS ---
    df_new_stats = pd.DataFrame(all_new_cases)
    df_disp_stats = pd.DataFrame(all_disposed_cases)
    df_audit = pd.DataFrame(all_audit_data)

    # Count by Category
    stats_new = df_new_stats['Category'].value_counts() if not df_new_stats.empty else pd.Series(dtype=int)
    stats_disp = df_disp_stats['Category'].value_counts() if not df_disp_stats.empty else pd.Series(dtype=int)

    # Combine into one table
    final_stats = pd.DataFrame({
        'New Cases (Arraigned)': stats_new,
        'Disposed Cases (Concluded)': stats_disp
    }).fillna(0).astype(int)

    # Sort to look standard
    sort_order = ["AGAINST LAWFUL AUTHORITY", "AGAINST PUBLIC MORALITY", "AGAINST THE PERSON", "AGAINST PROPERTY", "OTHERS"]
    final_stats = final_stats.reindex(sort_order).fillna(0).astype(int)
    
    # Calculate Total
    total_row = pd.DataFrame(final_stats.sum()).T
    total_row.index = ['TOTAL']
    final_stats = pd.concat([final_stats, total_row])

    # --- D. DISPLAY ---
    st.subheader(f"Statistics for {datetime(2025, report_month, 1).strftime('%B %Y')}")
    st.dataframe(final_stats)

    st.subheader("Individual Cases List (Audit Trail)")
    st.write("These are the specific rows detected for this month:")
    st.dataframe(df_audit)

    # Excel Download
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_stats.to_excel(writer, sheet_name='Summary_Statistics')
        df_audit.to_excel(writer, sheet_name='Individual_Cases_List', index=False)
    
    st.download_button(
        label="📥 Download Report (.xlsx)",
        data=output.getvalue(),
        file_name=f"Court_Report_{report_month}_{report_year}.xlsx",
        mime="application/vnd.ms-excel"
    )
