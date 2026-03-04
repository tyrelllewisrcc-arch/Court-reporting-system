import streamlit as st
import pandas as pd
import io
import openpyxl
from datetime import datetime

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="San Pedro Court Reporting System",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- GLOBAL CSS ---
st.markdown("""
<style>
    /* ---- Global Reset & Font ---- */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* ---- Hide default Streamlit chrome ---- */
    #MainMenu, footer, header { visibility: hidden; }

    /* ---- Sidebar ---- */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
        border-right: 1px solid #334155;
    }
    [data-testid="stSidebar"] * {
        color: #e2e8f0 !important;
    }
    [data-testid="stSidebar"] .stRadio label,
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stFileUploader label {
        color: #94a3b8 !important;
        font-size: 0.82rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }

    /* ---- Hero Banner ---- */
    .hero {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 50%, #0f4c75 100%);
        padding: 2.5rem 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        border: 1px solid #1e40af55;
        box-shadow: 0 8px 32px rgba(0,0,0,0.3);
    }
    .hero h1 {
        color: #f8fafc;
        font-size: 2.2rem;
        font-weight: 700;
        margin: 0 0 0.4rem 0;
    }
    .hero p {
        color: #94a3b8;
        font-size: 1.05rem;
        margin: 0;
    }
    .hero .badge {
        display: inline-block;
        background: #22c55e22;
        color: #22c55e;
        border: 1px solid #22c55e44;
        padding: 3px 12px;
        border-radius: 20px;
        font-size: 0.78rem;
        font-weight: 600;
        margin-bottom: 0.8rem;
        letter-spacing: 0.05em;
    }

    /* ---- Stat Cards ---- */
    .stat-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 1rem;
        margin-bottom: 2rem;
    }
    .stat-card {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 1.2rem 1.4rem;
        text-align: center;
        transition: transform 0.2s, border-color 0.2s;
    }
    .stat-card:hover {
        transform: translateY(-2px);
        border-color: #3b82f6;
    }
    .stat-card .icon { font-size: 1.8rem; margin-bottom: 0.4rem; }
    .stat-card .number {
        font-size: 2rem;
        font-weight: 700;
        color: #f1f5f9;
        line-height: 1;
    }
    .stat-card .label {
        font-size: 0.78rem;
        color: #64748b;
        margin-top: 0.3rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
    }

    /* ---- Section Cards ---- */
    .section-card {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1.2rem;
    }
    .section-card h3 {
        color: #e2e8f0;
        font-size: 1.05rem;
        font-weight: 600;
        margin: 0 0 0.5rem 0;
    }
    .section-card p {
        color: #94a3b8;
        font-size: 0.9rem;
        line-height: 1.6;
        margin: 0;
    }

    /* ---- Sheet Badge List ---- */
    .sheet-badge {
        display: inline-block;
        background: #1e40af22;
        color: #60a5fa;
        border: 1px solid #1e40af55;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.78rem;
        font-weight: 500;
        margin: 3px 3px 3px 0;
    }

    /* ---- Step Guide ---- */
    .step-row {
        display: flex;
        align-items: flex-start;
        gap: 1rem;
        margin-bottom: 1.2rem;
    }
    .step-num {
        background: linear-gradient(135deg, #3b82f6, #1d4ed8);
        color: white;
        width: 32px;
        height: 32px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 0.9rem;
        flex-shrink: 0;
    }
    .step-content h4 {
        color: #e2e8f0;
        font-size: 0.95rem;
        font-weight: 600;
        margin: 0 0 0.25rem 0;
    }
    .step-content p {
        color: #94a3b8;
        font-size: 0.85rem;
        margin: 0;
        line-height: 1.5;
    }

    /* ---- Nav Pills ---- */
    .nav-pill {
        display: inline-block;
        padding: 6px 16px;
        border-radius: 8px;
        font-size: 0.88rem;
        font-weight: 500;
        cursor: pointer;
        transition: background 0.2s;
    }
    .nav-pill.active {
        background: #3b82f6;
        color: white;
    }
    .nav-pill.inactive {
        background: #1e293b;
        color: #94a3b8;
        border: 1px solid #334155;
    }

    /* ---- Process Button ---- */
    .stButton > button {
        background: linear-gradient(135deg, #3b82f6, #1d4ed8) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.5rem !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        letter-spacing: 0.02em !important;
        transition: opacity 0.2s !important;
        width: 100%;
    }
    .stButton > button:hover { opacity: 0.88 !important; }

    /* ---- Download Button ---- */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #22c55e, #15803d) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.5rem !important;
        font-weight: 600 !important;
        width: 100%;
    }

    /* ---- Dataframe ---- */
    .stDataFrame { border-radius: 8px; overflow: hidden; }

    /* ---- Divider ---- */
    hr { border-color: #334155 !important; }

    /* ---- Sidebar nav items ---- */
    .sidebar-nav-item {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 10px 14px;
        border-radius: 8px;
        margin-bottom: 4px;
        cursor: pointer;
        font-size: 0.92rem;
        font-weight: 500;
        color: #94a3b8;
        transition: background 0.2s;
    }
    .sidebar-nav-item.active {
        background: #3b82f622;
        color: #60a5fa !important;
        border-left: 3px solid #3b82f6;
    }
    .sidebar-logo {
        padding: 1.2rem 1rem 1rem 1rem;
        border-bottom: 1px solid #334155;
        margin-bottom: 1rem;
    }
    .sidebar-logo h2 {
        color: #f1f5f9 !important;
        font-size: 1.1rem;
        font-weight: 700;
        margin: 0;
    }
    .sidebar-logo span {
        color: #60a5fa !important;
        font-size: 0.78rem;
    }

    /* ---- Alert boxes ---- */
    .alert-info {
        background: #1e40af18;
        border: 1px solid #3b82f644;
        border-left: 4px solid #3b82f6;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        color: #93c5fd;
        font-size: 0.88rem;
        margin-bottom: 1rem;
    }
    .alert-success {
        background: #15803d18;
        border: 1px solid #22c55e44;
        border-left: 4px solid #22c55e;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        color: #86efac;
        font-size: 0.88rem;
        margin-bottom: 1rem;
    }

    /* ---- Table of supported sheets ---- */
    .sheet-table { width: 100%; border-collapse: collapse; }
    .sheet-table th {
        background: #0f172a;
        color: #64748b;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        padding: 8px 12px;
        text-align: left;
        border-bottom: 1px solid #334155;
    }
    .sheet-table td {
        padding: 8px 12px;
        font-size: 0.85rem;
        color: #cbd5e1;
        border-bottom: 1px solid #1e293b;
    }
    .sheet-table tr:hover td { background: #1e293b88; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════════

def smart_read_excel(file):
    """Intelligently reads court data Excel, auto-detecting headers."""
    if not file:
        return None
    df_preview = pd.read_excel(file, header=None, nrows=20)
    header_row = 0
    found = False
    for idx, row in df_preview.iterrows():
        row_text = " ".join([str(x).upper() for x in row.values])
        if 'COURT BOOK' in row_text and ('CHARGE' in row_text or 'OFFENCE' in row_text):
            header_row = idx
            found = True
            break

    if not found:
        st.error(f"Could not find headers in {file.name}. Expected 'COURT BOOK' and 'CHARGE'/'OFFENCE' columns.")
        return None

    df = pd.read_excel(file, header=header_row)
    df.columns = [str(c).strip().upper() for c in df.columns]

    col_map = {}
    for c in df.columns:
        if 'COURT BOOK' in c:              col_map[c] = 'CASEID'
        elif 'CHARGE' in c or 'OFFENCE' in c: col_map[c] = 'CHARGE'
        elif 'COMPLAINANT' in c or 'VICTIM' in c: col_map[c] = 'VICTIM'
        elif 'ARRAINGMENT' in c or 'ARRAIGNMENT' in c: col_map[c] = 'DATE_ARR'
        elif 'CONCLUDED' in c or 'DISPOSAL' in c: col_map[c] = 'DATE_DISP'
        elif 'AGE' in c:                   col_map[c] = 'AGE'
        elif 'SEX' in c or 'GENDER' in c:  col_map[c] = 'GENDER'
        elif 'FURTHER' in c:               col_map[c] = 'SENTENCE'
        elif 'STATUS' in c:                col_map[c] = 'CASE_STATUS'
        elif 'REMARK' in c:                col_map[c] = 'REMARK'

    df = df.rename(columns=col_map)
    return df.loc[:, ~df.columns.duplicated()]


def classify_crime_sheet1(charge, victim):
    charge = str(charge).upper()
    victim = str(victim).upper()
    if any(k in victim for k in ['POLICE', 'PC ', 'CPL ', 'GOB']) and 'MINOR' not in victim:
        return 25
    if 'ESCAPE' in charge:                          return 24
    if 'PERJURY' in charge:                         return 23
    if any(x in charge for x in ['DISORDERLY', 'ABUSIVE', 'THREAT']): return 22
    if 'RAPE' in charge:                            return 27
    if 'SEXUAL ASSAULT' in charge:                  return 28
    if 'UNLAWFUL SEXUAL' in charge:                 return 29
    if 'UNNATURAL' in charge:                       return 30
    if 'ATTEMPT' in charge and 'MURDER' in charge:  return 35
    if 'MURDER' in charge:                          return 33
    if 'MANSLAUGHTER' in charge:                    return 34
    if 'GRIEVOUS' in charge:                        return 36
    if 'WOUNDING' in charge:                        return 37
    if 'HARM' in charge:                            return 38
    if 'AGGRAVATED ASSAULT' in charge:              return 39
    if 'COMMON ASSAULT' in charge:                  return 40
    if 'ROBBERY' in charge:                         return 43
    if 'BURGLARY' in charge:                        return 44
    if 'THEFT' in charge:                           return 45
    if 'DECEPTION' in charge or 'FRAUD' in charge:  return 46
    if 'HANDLING' in charge:                        return 47
    if 'DAMAGE' in charge:                          return 48
    if 'ARSON' in charge:                           return 49
    if 'FORGERY' in charge:                         return 52
    if 'DRUG' in charge or 'CANNABIS' in charge:    return 54
    if 'PIPE' in charge:                            return 57
    if 'VEHICLE' in charge:                         return 58
    if 'TRAFFIC' in charge or 'MOTOR' in charge or 'LICENSE' in charge: return 59
    if 'FIREARM' in charge or 'AMMUNITION' in charge: return 59
    return 59


def classify_statutory_sheet8(charge):
    c = str(charge).upper()
    if 'DRUG' in c or 'CANNABIS' in c:               return 12
    if 'FIREARM' in c or 'AMMUNITION' in c:           return 13
    if 'LIQUOR' in c:                                 return 14
    if 'POLICE' in c:                                 return 15
    if 'GAMBLING' in c:                               return 16
    if 'TRAFFIC' in c or 'MOTOR' in c or 'LICENSE' in c: return 17
    return 18


def parse_disposition(remark):
    r = str(remark).upper()
    if any(x in r for x in ['CONVICTED', 'GUILTY', 'FINE', 'PRISON']): return 'CONVICTED'
    if any(x in r for x in ['ACQUITTED', 'DISMISSED', 'STRUCK', 'DISCHARGED']): return 'DISMISSED'
    if any(x in r for x in ['WITHDRAWN', 'NOLLE']): return 'NOLLE'
    return 'OTHER'


def parse_sentence(sentence_text):
    s = str(sentence_text).upper()
    if 'FINE' in s or '$' in s:                                              return 'FINE'
    if any(x in s for x in ['PRISON', 'IMPRISONMENT', 'CONFINEMENT', 'MONTHS', 'YEARS']): return 'PRISON'
    if 'PROBATION' in s or 'BOND' in s:                                      return 'PROBATION'
    if 'REFORM' in s or 'SCHOOL' in s:                                       return 'REFORMATORY'
    return 'OTHER'


def is_juvenile(age):
    try:
        return int(age) <= 16
    except:
        return False


def get_age_col_sheet5(age, gender):
    g = str(gender).upper()
    is_male = 'F' not in g
    try:
        a = int(age)
    except:
        return None
    if a <= 16:       return 'B' if is_male else 'C'
    if 17 <= a <= 25: return 'D' if is_male else 'E'
    if 26 <= a <= 35: return 'F' if is_male else 'G'
    if 36 <= a <= 45: return 'H' if is_male else 'I'
    if a >= 46:       return 'J' if is_male else 'K'
    return None


def fill_all_sheets(template_file, df, mode):
    """Core engine: fills all 9 sheets of the statistical template."""
    wb = openpyxl.load_workbook(template_file)

    seen_cases = set()
    rows_sheet1 = []
    rows_sheet3 = []

    for idx, row in df.iterrows():
        r_num = classify_crime_sheet1(row.get('CHARGE', ''), row.get('VICTIM', ''))
        rows_sheet3.append(r_num)
        case_id = row.get('CASEID', idx)
        if case_id not in seen_cases:
            rows_sheet1.append(r_num)
            seen_cases.add(case_id)

    # Sheet 1 — Cases
    if 'Sheet1' in wb.sheetnames:
        ws = wb['Sheet1']
        col = 'D' if mode == "New" else 'J'
        for r in rows_sheet1:
            try:
                curr = ws[f"{col}{r}"].value or 0
                ws[f"{col}{r}"] = curr + 1
            except:
                pass

    # Sheet 3 — Persons
    if 'Sheet3' in wb.sheetnames:
        ws = wb['Sheet3']
        col = 'D' if mode == "New" else 'J'
        for r in rows_sheet3:
            try:
                curr = ws[f"{col}{r}"].value or 0
                ws[f"{col}{r}"] = curr + 1
            except:
                pass

    # Sheet 8 — Statutory Cases
    if 'Sheet8' in wb.sheetnames:
        ws = wb['Sheet8']
        for _, row in df.iterrows():
            stat_row = classify_statutory_sheet8(row.get('CHARGE', ''))
            if mode == "New":
                try:
                    ws[f"C{stat_row}"] = (ws[f"C{stat_row}"].value or 0) + 1
                except:
                    pass
            elif mode == "Disposed":
                disp = parse_disposition(row.get('REMARK', ''))
                if disp == 'CONVICTED':
                    try: ws[f"E{stat_row}"] = (ws[f"E{stat_row}"].value or 0) + 1
                    except: pass
                elif disp == 'DISMISSED':
                    try: ws[f"F{stat_row}"] = (ws[f"F{stat_row}"].value or 0) + 1
                    except: pass

    if mode == "Disposed":
        # Sheet 2 — Disposals
        if 'Sheet2' in wb.sheetnames:
            ws = wb['Sheet2']
            for _, row in df.iterrows():
                r_num = classify_crime_sheet1(row.get('CHARGE', ''), row.get('VICTIM', ''))
                disp = parse_disposition(row.get('REMARK', ''))
                target_col = {'CONVICTED': 'E', 'DISMISSED': 'C', 'NOLLE': 'D'}.get(disp)
                if target_col:
                    try: ws[f"{target_col}{r_num}"] = (ws[f"{target_col}{r_num}"].value or 0) + 1
                    except: pass

        for _, row in df.iterrows():
            if parse_disposition(row.get('REMARK', '')) != 'CONVICTED':
                continue
            r_num = classify_crime_sheet1(row.get('CHARGE', ''), row.get('VICTIM', ''))
            gender = row.get('GENDER', 'M')
            is_male = 'F' not in str(gender).upper()
            age = row.get('AGE', 0)
            sent_type = parse_sentence(row.get('SENTENCE', ''))

            # Sheet 4 — Sentence breakdown
            if 'Sheet4' in wb.sheetnames:
                ws = wb['Sheet4']
                s_col = {
                    'PRISON':    'D' if is_male else 'E',
                    'PROBATION': 'F' if is_male else 'G',
                    'FINE':      'H' if is_male else 'I'
                }.get(sent_type)
                if s_col:
                    try: ws[f"{s_col}{r_num}"] = (ws[f"{s_col}{r_num}"].value or 0) + 1
                    except: pass

            # Sheet 5 — Age demographics
            if 'Sheet5' in wb.sheetnames:
                ws = wb['Sheet5']
                a_col = get_age_col_sheet5(age, gender)
                if a_col:
                    try: ws[f"{a_col}{r_num-11}"] = (ws[f"{a_col}{r_num-11}"].value or 0) + 1
                    except: pass

            # Sheets 6 & 7 — Juveniles
            if is_juvenile(age):
                juv_row = r_num - 14
                if 'Sheet6' in wb.sheetnames:
                    ws_6 = wb['Sheet6']
                    try: ws_6[f"F{juv_row}"] = (ws_6[f"F{juv_row}"].value or 0) + 1
                    except: pass
                if 'Sheet7' in wb.sheetnames:
                    ws = wb['Sheet7']
                    sent_col = {
                        'PRISON': 'B', 'PROBATION': 'C',
                        'FINE': 'D', 'REFORMATORY': 'E'
                    }.get(sent_type)
                    if sent_col:
                        try: ws[f"{sent_col}{juv_row}"] = (ws[f"{sent_col}{juv_row}"].value or 0) + 1
                        except: pass

            # Sheet 9 — Statutory punishment
            if 'Sheet9' in wb.sheetnames:
                stat_row = classify_statutory_sheet8(row.get('CHARGE', ''))
                ws = wb['Sheet9']
                s_col = {
                    'PRISON':    'D' if is_male else 'E',
                    'PROBATION': 'B' if is_male else 'C',
                    'FINE':      'F' if is_male else 'G'
                }.get(sent_type)
                if s_col:
                    try: ws[f"{s_col}{stat_row}"] = (ws[f"{s_col}{stat_row}"].value or 0) + 1
                    except: pass

    return wb


# ═══════════════════════════════════════════════════
#  SIDEBAR NAVIGATION
# ═══════════════════════════════════════════════════

with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
        <h2>⚖️ San Pedro</h2>
        <span>Court Reporting System</span>
    </div>
    """, unsafe_allow_html=True)

    page = st.radio(
        "Navigation",
        ["🏠  Home", "📊  Generate Report", "📋  How It Works"],
        label_visibility="collapsed"
    )

    st.markdown("---")
    st.markdown("""
    <div style="padding: 0 4px;">
        <p style="color:#475569; font-size:0.75rem; text-transform:uppercase; letter-spacing:0.07em; margin-bottom:8px;">System Info</p>
        <p style="color:#64748b; font-size:0.8rem; margin:4px 0;">
            Version <strong style="color:#94a3b8;">2.0</strong>
        </p>
        <p style="color:#64748b; font-size:0.8rem; margin:4px 0;">
            Sheets supported: <strong style="color:#94a3b8;">9</strong>
        </p>
        <p style="color:#64748b; font-size:0.8rem; margin:4px 0;">
            Status: <span style="color:#22c55e; font-weight:600;">● Active</span>
        </p>
    </div>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════
#  PAGE: HOME
# ═══════════════════════════════════════════════════

if page == "🏠  Home":
    st.markdown("""
    <div class="hero">
        <div class="badge">● SYSTEM ACTIVE</div>
        <h1>San Pedro Court Reporting System</h1>
        <p>Automated statistical report generation for San Pedro Magistrate Court.<br>
        Upload your court data and receive a complete, formatted 9-sheet Excel report in seconds.</p>
    </div>
    """, unsafe_allow_html=True)

    # Stat cards
    st.markdown("""
    <div class="stat-grid">
        <div class="stat-card">
            <div class="icon">📄</div>
            <div class="number">9</div>
            <div class="label">Report Sheets</div>
        </div>
        <div class="stat-card">
            <div class="icon">🔍</div>
            <div class="number">30+</div>
            <div class="label">Crime Categories</div>
        </div>
        <div class="stat-card">
            <div class="icon">⚡</div>
            <div class="number">Auto</div>
            <div class="label">Column Detection</div>
        </div>
        <div class="stat-card">
            <div class="icon">📥</div>
            <div class="number">xlsx</div>
            <div class="label">Export Format</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="section-card">
            <h3>📊 Supported Report Sheets</h3>
            <p style="margin-bottom:0.8rem;">The system automatically fills all 9 sheets of the official statistical template:</p>
            <table class="sheet-table">
                <thead>
                    <tr><th>Sheet</th><th>Description</th><th>Mode</th></tr>
                </thead>
                <tbody>
                    <tr><td>Sheet 1</td><td>Main Crimes — Cases</td><td>Both</td></tr>
                    <tr><td>Sheet 2</td><td>Disposal Breakdown</td><td>Disposed</td></tr>
                    <tr><td>Sheet 3</td><td>Main Crimes — Persons</td><td>Both</td></tr>
                    <tr><td>Sheet 4</td><td>Convicted (Sentence Type)</td><td>Disposed</td></tr>
                    <tr><td>Sheet 5</td><td>Convicted (Age Groups)</td><td>Disposed</td></tr>
                    <tr><td>Sheet 6</td><td>Juvenile Offenses</td><td>Disposed</td></tr>
                    <tr><td>Sheet 7</td><td>Juvenile Sentences</td><td>Disposed</td></tr>
                    <tr><td>Sheet 8</td><td>Statutory Offenses</td><td>Both</td></tr>
                    <tr><td>Sheet 9</td><td>Statutory Punishment</td><td>Disposed</td></tr>
                </tbody>
            </table>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="section-card">
            <h3>⚡ Key Features</h3>
            <div class="step-row" style="margin-top:0.8rem;">
                <div class="step-num">✓</div>
                <div class="step-content">
                    <h4>Smart Header Detection</h4>
                    <p>Automatically identifies column headers across different Excel formats — no manual mapping required.</p>
                </div>
            </div>
            <div class="step-row">
                <div class="step-num">✓</div>
                <div class="step-content">
                    <h4>Intelligent Crime Classification</h4>
                    <p>Classifies 30+ offence types including drugs, firearms, assault, sexual offences, and traffic violations.</p>
                </div>
            </div>
            <div class="step-row">
                <div class="step-num">✓</div>
                <div class="step-content">
                    <h4>Monthly & Annual Reports</h4>
                    <p>Filter data by a specific month or generate a full-year statistical summary.</p>
                </div>
            </div>
            <div class="step-row">
                <div class="step-num">✓</div>
                <div class="step-content">
                    <h4>Juvenile Case Handling</h4>
                    <p>Automatically segments juvenile offenders (age ≤ 16) into dedicated sheets 6 & 7.</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div class="alert-info">
        💡 <strong>Quick Start:</strong> Navigate to <em>Generate Report</em> in the sidebar, upload your data file and blank template, configure your settings, and click Process.
    </div>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════
#  PAGE: GENERATE REPORT
# ═══════════════════════════════════════════════════

elif page == "📊  Generate Report":
    st.markdown("""
    <div class="hero">
        <div class="badge">REPORT GENERATOR</div>
        <h1>Generate Statistical Report</h1>
        <p>Upload your court data Excel file and a blank 9-sheet template to produce a filled statistical report.</p>
    </div>
    """, unsafe_allow_html=True)

    # Upload and config
    col_left, col_right = st.columns([1, 1])

    with col_left:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown("**📁 Step 1: Upload Files**")
        data_file = st.file_uploader("Court Data File (Excel)", type=['xlsx'], key="data_upload",
                                     help="Your raw court book data with case records")
        template_file = st.file_uploader("Blank Report Template (Excel)", type=['xlsx'], key="tpl_upload",
                                         help="The official 9-sheet statistical template, unfilled")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_right:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown("**⚙️ Step 2: Configure Report**")

        mode = st.radio(
            "Data Type",
            ["New Cases (Arraignments)", "Disposed Cases (Concluded)"],
            help="New Cases fills Sheets 1, 3, 8. Disposed Cases fills all 9 sheets."
        )
        is_full_year = st.checkbox("Full Year Report", help="If unchecked, filter by a single month")
        report_year = st.number_input("Year", value=2025, min_value=2000, max_value=2100)
        if not is_full_year:
            report_month = st.selectbox(
                "Month",
                range(1, 13),
                format_func=lambda x: datetime(2025, x, 1).strftime('%B')
            )
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # Process button
    col_btn, col_info = st.columns([1, 2])
    with col_btn:
        process = st.button("🚀 Process & Generate Report")

    if process:
        if not data_file or not template_file:
            st.error("Please upload **both** the data file and the blank template before processing.")
            st.stop()

        with st.spinner("Reading and validating data..."):
            df = smart_read_excel(data_file)
        if df is None:
            st.stop()

        date_col = 'DATE_ARR' if mode.startswith("New") else 'DATE_DISP'
        if date_col not in df.columns:
            date_label = 'Arraignment' if mode.startswith('New') else 'Concluded/Disposal'
            st.error(f"Missing required date column. The data file must contain a '{date_label}' date column.")
            st.stop()

        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

        if is_full_year:
            mask = df[date_col].dt.year == report_year
            period_name = f"Full Year {int(report_year)}"
        else:
            mask = (df[date_col].dt.month == report_month) & (df[date_col].dt.year == report_year)
            period_name = datetime(2025, report_month, 1).strftime('%B') + f" {int(report_year)}"

        df_filtered = df[mask].copy()

        if df_filtered.empty:
            st.warning(f"No records found for **{period_name}**. Try a different period or check your date columns.")
            st.stop()

        st.markdown(f"""
        <div class="alert-success">
            ✅ Found <strong>{len(df_filtered)}</strong> records for <strong>{period_name}</strong>. Filling report template…
        </div>
        """, unsafe_allow_html=True)

        with st.spinner("Filling all 9 report sheets..."):
            try:
                wb_filled = fill_all_sheets(
                    template_file,
                    df_filtered,
                    "New" if mode.startswith("New") else "Disposed"
                )
                out = io.BytesIO()
                wb_filled.save(out)
                out.seek(0)
            except Exception as e:
                st.error(f"Processing error: {e}")
                st.stop()

        st.success("Report generated successfully!")

        fname = f"SanPedro_Stats_9SHEETS_{period_name.replace(' ', '_')}.xlsx"
        st.download_button(
            label="📥 Download Complete 9-Sheet Report",
            data=out,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Summary metrics
        disp_counts = {}
        if 'REMARK' in df_filtered.columns:
            for _, row in df_filtered.iterrows():
                d = parse_disposition(row.get('REMARK', ''))
                disp_counts[d] = disp_counts.get(d, 0) + 1

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Records", len(df_filtered))
        m2.metric("Unique Cases", df_filtered['CASEID'].nunique() if 'CASEID' in df_filtered.columns else "—")
        m3.metric("Convicted", disp_counts.get('CONVICTED', 0))
        m4.metric("Dismissed", disp_counts.get('DISMISSED', 0))

        st.markdown("### Data Preview (first 50 rows)")
        st.dataframe(df_filtered.head(50), use_container_width=True)


# ═══════════════════════════════════════════════════
#  PAGE: HOW IT WORKS
# ═══════════════════════════════════════════════════

elif page == "📋  How It Works":
    st.markdown("""
    <div class="hero">
        <div class="badge">DOCUMENTATION</div>
        <h1>How It Works</h1>
        <p>A step-by-step guide to using the San Pedro Court Reporting System.</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="section-card">
            <h3>📋 Step-by-Step Guide</h3>

            <div class="step-row" style="margin-top:1rem;">
                <div class="step-num">1</div>
                <div class="step-content">
                    <h4>Prepare Your Data File</h4>
                    <p>Export your court register to Excel (.xlsx). The system will auto-detect column headers.
                    Expected columns: <em>Court Book No., Charge/Offence, Complainant/Victim, Arraignment Date,
                    Concluded Date, Age, Sex/Gender, Sentence, Status, Remarks.</em></p>
                </div>
            </div>

            <div class="step-row">
                <div class="step-num">2</div>
                <div class="step-content">
                    <h4>Get the Blank Template</h4>
                    <p>Obtain the official 9-sheet statistical Excel template. This must be the standard template
                    with sheets named Sheet1 through Sheet9.</p>
                </div>
            </div>

            <div class="step-row">
                <div class="step-num">3</div>
                <div class="step-content">
                    <h4>Select Report Mode</h4>
                    <p><strong>New Cases</strong> — use for arraignment data. Fills Sheets 1, 3, and 8.<br>
                    <strong>Disposed Cases</strong> — use for concluded/disposal data. Fills all 9 sheets.</p>
                </div>
            </div>

            <div class="step-row">
                <div class="step-num">4</div>
                <div class="step-content">
                    <h4>Choose Period</h4>
                    <p>Select the year and optionally a specific month. Enable <em>Full Year Report</em>
                    to aggregate all months together.</p>
                </div>
            </div>

            <div class="step-row">
                <div class="step-num">5</div>
                <div class="step-content">
                    <h4>Process & Download</h4>
                    <p>Click <em>Process & Generate Report</em>. When complete, download the filled
                    Excel file with all statistics populated.</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="section-card">
            <h3>🔍 Column Detection Rules</h3>
            <p style="margin-bottom:0.8rem;">The system searches the first 20 rows of your file for a header row
            containing <em>Court Book</em> and <em>Charge/Offence</em>. Columns are then mapped automatically:</p>
            <table class="sheet-table">
                <thead><tr><th>Detected Keyword</th><th>Mapped To</th></tr></thead>
                <tbody>
                    <tr><td>Court Book</td><td>Case ID</td></tr>
                    <tr><td>Charge / Offence</td><td>Charge</td></tr>
                    <tr><td>Complainant / Victim</td><td>Victim</td></tr>
                    <tr><td>Arraignment / Arraingment</td><td>Arraignment Date</td></tr>
                    <tr><td>Concluded / Disposal</td><td>Disposal Date</td></tr>
                    <tr><td>Age</td><td>Age</td></tr>
                    <tr><td>Sex / Gender</td><td>Gender</td></tr>
                    <tr><td>Further / Sentence</td><td>Sentence</td></tr>
                    <tr><td>Status</td><td>Case Status</td></tr>
                    <tr><td>Remark</td><td>Remark</td></tr>
                </tbody>
            </table>
        </div>

        <div class="section-card" style="margin-top:1.2rem;">
            <h3>⚖️ Classification Logic</h3>
            <p>Charges are parsed using keyword matching:</p>
            <div style="margin-top:0.5rem;">
                <span class="sheet-badge">Murder / Manslaughter</span>
                <span class="sheet-badge">Grievous Harm</span>
                <span class="sheet-badge">Wounding</span>
                <span class="sheet-badge">Assault</span>
                <span class="sheet-badge">Rape / Sexual Assault</span>
                <span class="sheet-badge">Robbery</span>
                <span class="sheet-badge">Burglary</span>
                <span class="sheet-badge">Theft</span>
                <span class="sheet-badge">Fraud / Deception</span>
                <span class="sheet-badge">Drug Offences</span>
                <span class="sheet-badge">Firearms</span>
                <span class="sheet-badge">Traffic Offences</span>
                <span class="sheet-badge">Arson</span>
                <span class="sheet-badge">Forgery</span>
                <span class="sheet-badge">Disorderly Conduct</span>
                <span class="sheet-badge">Perjury / Escape</span>
            </div>
            <p style="margin-top:0.8rem;">Disposition is determined from <em>Remarks</em>: <strong>Convicted</strong>,
            <strong>Dismissed</strong>, or <strong>Nolle Prosequi / Withdrawn</strong>.</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div class="alert-info" style="margin-top:0.5rem;">
        ℹ️ <strong>Data Privacy:</strong> All processing happens locally in your browser session.
        No court data is stored or transmitted to external servers.
    </div>
    """, unsafe_allow_html=True)
