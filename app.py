import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import LinearSVC
from sklearn.pipeline import make_pipeline
import io

# --- PAGE CONFIG ---
st.set_page_config(page_title="CSRS System", layout="wide")

# --- AUTHENTICATION & SETUP ---
SCOPE = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

@st.cache_resource
def connect_to_sheets():
    # Load credentials from Streamlit Secrets (NOT from a file)
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
    client = gspread.authorize(creds)
    # CONNECT TO YOUR SPECIFIC GOOGLE SHEET NAME HERE
    sheet = client.open("CSRS_Database") 
    return sheet

# --- MAIN APP LOGIC ---
st.title("⚖️ Automated Court Statistical Reporting System")

# 1. Password Protection
password_guess = st.sidebar.text_input("System Password", type="password")
if password_guess != st.secrets["APP_PASSWORD"]:
    st.warning("Please enter the correct password to access the system.")
    st.stop()

# 2. Connection
try:
    sheet = connect_to_sheets()
    st.sidebar.success("Database Connected")
except Exception as e:
    st.error(f"Connection Error: {e}")
    st.stop()

# 3. Load Data & Train AI
ws_train = sheet.worksheet("TrainingData")
ws_raw = sheet.worksheet("RawData")

df_train = pd.DataFrame(ws_train.get_all_records())
df_raw = pd.DataFrame(ws_raw.get_all_records())

ai_model = None
if not df_train.empty:
    ai_model = make_pipeline(TfidfVectorizer(), LinearSVC())
    ai_model.fit(df_train['Description'].astype(str), df_train['Category'])
    st.sidebar.info(f"AI Trained on {len(df_train)} examples.")

# --- TABS ---
tab1, tab2, tab3 = st.tabs(["📥 Ingestion", "📊 Reports", "🧠 Train AI"])

with tab1:
    st.header("Upload New Case Data")
    uploaded_file = st.file_uploader("Upload Monthly Excel", type=['xlsx'])
    
    if uploaded_file:
        new_data = pd.read_excel(uploaded_file)
        st.write("Preview:", new_data.head())
        
        if st.button("Process & Predict Categories"):
            # AI Prediction
            if ai_model:
                new_data['AI_Category'] = ai_model.predict(new_data['Description'].astype(str))
            else:
                new_data['AI_Category'] = "Uncategorized"
            
            st.dataframe(new_data)
            st.warning("To save this data, we would append it to Google Sheets here (requires write permissions setup).")

with tab2:
    st.header("Statistical Reports")
    if not df_raw.empty:
        # Create Pivot Table (San Pedro Format)
        pivot = pd.pivot_table(
            df_raw, 
            index=['Category'], 
            columns=['Status'], 
            values='CaseID', 
            aggfunc='count', 
            fill_value=0,
            margins=True,
            margins_name='Total'
        )
        st.dataframe(pivot)
        
        # Download Button
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            pivot.to_excel(writer, sheet_name='Stats')
        
        st.download_button("Download Report", buffer, "Monthly_Stats.xlsx")
    else:
        st.info("No data in 'RawData' tab yet.")

with tab3:
    st.header("Teach the AI")
    with st.form("teach_bot"):
        desc = st.text_input("Charge Description (e.g. 'Theft of Motor Vehicle')")
        cat = st.selectbox("Correct Category", ["AGAINST PROPERTY", "AGAINST PERSON", "OTHER"])
        if st.form_submit_button("Add Rule"):
            ws_train.append_row([desc, cat])
            st.success("Rule added! Refresh the app to retrain the AI.")
