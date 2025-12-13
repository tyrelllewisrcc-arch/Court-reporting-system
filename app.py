import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# --- CONFIGURATION ---
st.set_page_config(page_title="Belize Court Reporting System", layout="wide")
st.title("🇧🇿 Automated Court Statistical Reporting System (CSRS)")

# --- 1. CRIME CATEGORIZATION ENGINE ---
def get_category_and_subcategory(charge, complainant):
    """
    Maps a charge description to the San Pedro Stats Categories.
    Prioritizes the 'Police Rule' (Against Lawful Authority).
    """
    charge = str(charge).upper().strip()
    complainant = str(complainant).upper().strip()

    # --- RULE 1: THE POLICE RULE ---
    # If complainant is Police/GOB, it is AUTOMATICALLY Against Lawful Authority
    police_keywords = ['POLICE', 'PC ', 'WPC ', 'CPL ', 'SGT ', 'INSP ', 'GOB', 'DEPARTMENT']
    if any(k in complainant for k in police_keywords) and "MINOR" not in complainant:
        if any(x in charge for x in ['ASSAULT', 'RESIST', 'OBSTRUCT']):
            return "AGAINST LAWFUL AUTHORITY", "Assault/Resist Police"
        return "AGAINST LAWFUL AUTHORITY", "Other Police Offenses"

    # --- RULE 2: CATEGORY MAPPING ---
    
    # A. AGAINST LAWFUL AUTHORITY (Non-Police victim specific)
    if any(x in charge for x in ['ESCAPE', 'RESCUE']): return "AGAINST LAWFUL AUTHORITY", "Escape and Rescue"
    if any(x in charge for x in ['PUBLIC TERROR', 'DISORDERLY', 'ABUSIVE', 'THREATENING WORDS']): return "AGAINST LAWFUL AUTHORITY", "Against public order"
    if any(x in charge for x in ['PERJURY']): return "AGAINST LAWFUL AUTHORITY", "Perjury"

    # B. AGAINST PUBLIC MORALITY
    if any(x in charge for x in ['RAPE']): return "AGAINST PUBLIC MORALITY", "Rape"
    if any(x in charge for x in ['UNLAWFUL SEXUAL']): return "AGAINST PUBLIC MORALITY", "Unlawful Sexual intercourse"
    if any(x in charge for x in ['SEXUAL ASSAULT']): return "AGAINST PUBLIC MORALITY", "Sexual Assault"
    if any(x in charge for x in ['UNNATURAL']): return "AGAINST PUBLIC MORALITY", "Unnatural offences"

    # C. AGAINST THE PERSON
    if any(x in charge for x in ['MURDER']) and 'ATTEMPT' not in charge: return "AGAINST THE PERSON", "Murder"
    if any(x in charge for x in ['MANSLAUGHTER']): return "AGAINST THE PERSON", "Manslaughter"
    if any(x in charge for x in ['ATTEMPT MURDER', 'ATTEMPT TO MURDER']): return "AGAINST THE PERSON", "Attempted Murder"
    if any(x in charge for x in ['GRIEVOUS HARM', 'DANGEROUS HARM']): return "AGAINST THE PERSON", "Grievous Harm"
    if any(x in charge for x in ['WOUNDING']): return "AGAINST THE PERSON", "Wounding"
    if 'HARM' in charge: return "AGAINST THE PERSON", "Harm"
    if 'AGGRAVATED ASSAULT' in charge: return "AGAINST THE PERSON", "Aggravated Assault"
    if 'COMMON ASSAULT' in charge: return "AGAINST THE PERSON", "Common Assault"
    
    # D. AGAINST PROPERTY
    if any(x in charge for x in ['ROBBERY']): return "AGAINST PROPERTY", "Robbery"
    if any(x in charge for x in ['BURGLARY']): return "AGAINST PROPERTY", "Burglary"
    if any(x in charge for x in ['THEFT']): return "AGAINST PROPERTY", "Theft"
    if any(x in charge for x in ['DECEPTION', 'FRAUD', 'FALSE PRETENSE']): return "AGAINST PROPERTY", "False Pretence/Fraud"
    if any(x in charge for x in ['HANDLING']): return "AGAINST PROPERTY", "Handling Stolen Goods"
    if any(x in charge for x in ['DAMAGE TO PROPERTY']): return "AGAINST PROPERTY", "Damage to Property"
    if any(x in charge for x in ['ARSON']): return "AGAINST PROPERTY", "Arson"

    # E. OTHERS (Drugs, Traffic, Firearms)
    if any(x in charge for x in ['DRUG', 'CANNABIS', 'COCAINE', 'PIPE']): return "OTHERS", "Drugs"
    if any(x in charge for x in ['FIREARM', 'AMMUNITION', 'GANG']): return "OTHERS", "Firearms/Gang"
    if any(x in charge for x in ['TRAFFIC', 'MOTOR', 'LICENSE', 'INSURANCE', 'DRIVE', 'DRIVING']): return "OTHERS", "Traffic"
    if any(x in charge for x in ['FORGERY']): return "OTHERS", "Forgery"

    return "OTHERS", "Other Offenses"

# --- 2. DATA PROCESSING HELPERS ---

def load_file(uploaded_file):
    if uploaded_file:
        # Docs say Header is Row 4 (Index 3)
        return pd.read_excel(uploaded_file, header=3)
    return None

def determine_age_group(age):
    try:
        age = int(age)
    except:
        return "Unknown"
    
    if age <= 16: return "Juvenile (<=16)"
    if 17 <= age <= 25: return "17-25"
    if 26 <= age <= 35: return "26-35"
    if 36 <= age <= 45: return "36-45"
    if age >= 46: return "46+"
    return "Unknown"

def clean_gender(sex):
    s = str(sex).upper().strip()
    if s.startswith('M'): return 'Male'
    if s.startswith('F'): return 'Female'
    return 'Unknown'

# --- 3. MAIN APP INTERFACE ---

st.sidebar.header("1. Upload Data")
returns_file = st.sidebar.file_uploader("Upload 'Returns' (Concluded Cases)", type=['xlsx'])
active_file = st.sidebar.file_uploader("Upload 'Active Cases' (Pending)", type=['xlsx'])

st.sidebar.header("2. Report Settings")
report_month = st.sidebar.selectbox("Select Month", range(1, 13), format_func=lambda x: datetime(2025, x, 1).strftime('%B'))
report_year = st.sidebar.number_input("Year", value=2025)

if st.sidebar.button("Generate Statistical Report"):
    if not returns_file or not active_file:
        st.error("Please upload both the Returns and Active Cases files.")
    else:
        # --- A. LOAD DATA ---
        df_returns = load_file(returns_file)
        df_active = load_file(active_file)

        # Standardize Column Names (Strip whitespace)
        df_returns.columns = df_returns.columns.str.strip()
        df_active.columns = df_active.columns.str.strip()

        # --- B. FILTER BY MONTH (Returns File) ---
        # Ensure Date Concluded is datetime
        df_returns['Date Concluded'] = pd.to_datetime(df_returns['Date Concluded'], errors='coerce')
        
        # Filter Logic
        mask = (df_returns['Date Concluded'].dt.month == report_month) & \
               (df_returns['Date Concluded'].dt.year == report_year)
        df_monthly = df_returns[mask].copy()

        st.success(f"Processing {len(df_monthly)} concluded cases for {datetime(2025, report_month, 1).strftime('%B')}.")

        # --- C. ENRICH DATA (Apply Categorization) ---
        
        # Process Returns Data
        categories = []
        subcategories = []
        age_groups = []
        clean_genders = []
        is_convicted = []
        
        for index, row in df_monthly.iterrows():
            # 1. Categorize
            cat, sub = get_category_and_subcategory(row.get('Charge', ''), row.get('Complainant/Victim', ''))
            categories.append(cat)
            subcategories.append(sub)
            
            # 2. Age Group
            age_groups.append(determine_age_group(row.get('Age', 0)))
            
            # 3. Gender
            clean_genders.append(clean_gender(row.get('Sex', '')))

            # 4. Conviction Status
            remark = str(row.get('Remark', '')).upper()
            is_convicted.append("CONVICTED" in remark)

        df_monthly['Category'] = categories
        df_monthly['Subcategory'] = subcategories
        df_monthly['AgeGroup'] = age_groups
        df_monthly['CleanGender'] = clean_genders
        df_monthly['IsConvicted'] = is_convicted

        # Process Active Data (For Pending Counts)
        active_cats = []
        for index, row in df_active.iterrows():
            # Use 'Charge (2)' as per docs, or fallback to 'Charge'
            chg = row.get('Charge (2)', row.get('Charge', '')) 
            cat, sub = get_category_and_subcategory(chg, row.get('Complainant/Victim', ''))
            active_cats.append(cat)
        df_active['Category'] = active_cats

        # --- D. GENERATE STATISTICS ---

        # 1. SHEET 1: CASES (New vs Disposed vs Pending)
        # Note: Sheet 1 counts CASES (Unique Court Book No), not Persons.
        
        # New Cases (from Returns as per doc instructions)
        new_cases = df_monthly.drop_duplicates(subset=['Court Book No.'])['Category'].value_counts()
        
        # Disposed Cases
        disposed_cases = df_monthly.drop_duplicates(subset=['Court Book No.'])['Category'].value_counts()
        
        # Pending Cases (From Active File)
        pending_cases = df_active.drop_duplicates(subset=['Court Book No.'])['Category'].value_counts()

        sheet1_data = pd.DataFrame({
            'New Cases (This Month)': new_cases,
            'Disposed (This Month)': disposed_cases,
            'Pending (Total Active)': pending_cases
        }).fillna(0)

        # 2. SHEET 3: PERSONS (Count all rows)
        sheet3_data = df_monthly['Category'].value_counts().rename("Total Persons Involved")

        # 3. SHEET 5: CONVICTED BY AGE (Convicted Only)
        df_convicted = df_monthly[df_monthly['IsConvicted'] == True]
        
        sheet5_pivot = pd.pivot_table(
            df_convicted, 
            index=['Category'], 
            columns=['AgeGroup', 'CleanGender'], 
            values='Court Book No.', 
            aggfunc='count', 
            fill_value=0
        )

        # 4. SHEET 6: JUVENILES (Age <= 16, Convicted)
        df_juvenile = df_convicted[df_convicted['AgeGroup'] == "Juvenile (<=16)"]
        sheet6_data = df_juvenile.groupby(['Category', 'CleanGender']).size().unstack(fill_value=0)

        # --- E. DISPLAY & EXPORT ---
        
        st.subheader("📊 Preview: Main Statistics (Sheet 1)")
        st.dataframe(sheet1_data)

        st.subheader("👦 Preview: Juvenile Convictions (Sheet 6)")
        if not df_juvenile.empty:
            st.dataframe(sheet6_data)
        else:
            st.info("No juvenile convictions found for this month.")

        # Excel Export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            sheet1_data.to_excel(writer, sheet_name='Sheet1_MainStats')
            sheet3_data.to_excel(writer, sheet_name='Sheet3_Persons')
            sheet5_pivot.to_excel(writer, sheet_name='Sheet5_AgeGroups')
            if not df_juvenile.empty:
                sheet6_data.to_excel(writer, sheet_name='Sheet6_Juveniles')
            
            # Raw Data Dump for checking
            df_monthly.to_excel(writer, sheet_name='Verified_Data_Dump')

        st.download_button(
            label="📥 Download Generated Statistics (.xlsx)",
            data=output.getvalue(),
            file_name=f"San_Pedro_Stats_{report_year}_{report_month}.xlsx",
            mime="application/vnd.ms-excel"
        )
