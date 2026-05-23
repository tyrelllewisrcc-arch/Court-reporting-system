import os
import io
import json
import time
import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime

# ---------------------------------------------------------------------------
# PAGE CONFIG (must be first Streamlit call)
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="San Pedro Court Reporting Suite",
    page_icon="⚖️",
    layout="wide",
)

# ---------------------------------------------------------------------------
# TOP-LEVEL NAVIGATION TABS
# ---------------------------------------------------------------------------
tab_reports, tab_ads = st.tabs(["⚖️ Court Reports", "📢 Advertising Agent"])

# ============================================================================
# TAB 1 — COURT REPORTS (original functionality, unchanged)
# ============================================================================
with tab_reports:
    st.title("⚖️ San Pedro Court Report: Complete 9-Sheet Auto-Filler")
    st.markdown("""
**System Status:** ✅ Active
**Supported Sheets:**
*   **1 & 3:** Main Crimes (Cases & Persons)
*   **2:** Disposal Breakdown
*   **4 & 5:** Convicted Demographics (Sentence & Age)
*   **6 & 7:** Juvenile Analysis (Offenses & Sentences)
*   **8 & 9:** Statutory Offenses (Drugs, Firearms, Traffic, etc.)
""")

    # --- 1. SMART COLUMN MAPPING ---
    def smart_read_excel(file):
        if not file:
            return None
        df_preview = pd.read_excel(file, header=None, nrows=20)
        header_row = 0
        found = False
        for idx, row in df_preview.iterrows():
            row_text = " ".join([str(x).upper() for x in row.values])
            if "COURT BOOK" in row_text and ("CHARGE" in row_text or "OFFENCE" in row_text):
                header_row = idx
                found = True
                break

        if not found:
            st.error(f"Could not find headers in {file.name}.")
            return None

        df = pd.read_excel(file, header=header_row)
        df.columns = [str(c).strip().upper() for c in df.columns]

        col_map = {}
        for c in df.columns:
            if "COURT BOOK" in c:
                col_map[c] = "CASEID"
            elif "CHARGE" in c or "OFFENCE" in c:
                col_map[c] = "CHARGE"
            elif "COMPLAINANT" in c or "VICTIM" in c:
                col_map[c] = "VICTIM"
            elif "ARRAINGMENT" in c or "ARRAIGNMENT" in c:
                col_map[c] = "DATE_ARR"
            elif "CONCLUDED" in c or "DISPOSAL" in c:
                col_map[c] = "DATE_DISP"
            elif "AGE" in c:
                col_map[c] = "AGE"
            elif "SEX" in c or "GENDER" in c:
                col_map[c] = "GENDER"
            elif "FURTHER" in c:
                col_map[c] = "SENTENCE"
            elif "STATUS" in c:
                col_map[c] = "CASE_STATUS"
            elif "REMARK" in c:
                col_map[c] = "REMARK"

        df = df.rename(columns=col_map)
        return df.loc[:, ~df.columns.duplicated()]

    # --- 2. INTELLIGENT PARSERS ---
    def classify_crime_sheet1(charge, victim):
        charge = str(charge).upper()
        victim = str(victim).upper()

        if any(k in victim for k in ["POLICE", "PC ", "CPL ", "GOB"]) and "MINOR" not in victim:
            return 25

        if "ESCAPE" in charge:
            return 24
        if "PERJURY" in charge:
            return 23
        if any(x in charge for x in ["DISORDERLY", "ABUSIVE", "THREAT"]):
            return 22
        if "RAPE" in charge:
            return 27
        if "SEXUAL ASSAULT" in charge:
            return 28
        if "UNLAWFUL SEXUAL" in charge:
            return 29
        if "UNNATURAL" in charge:
            return 30
        if "ATTEMPT" in charge and "MURDER" in charge:
            return 35
        if "MURDER" in charge:
            return 33
        if "MANSLAUGHTER" in charge:
            return 34
        if "GRIEVOUS" in charge:
            return 36
        if "WOUNDING" in charge:
            return 37
        if "HARM" in charge:
            return 38
        if "AGGRAVATED ASSAULT" in charge:
            return 39
        if "COMMON ASSAULT" in charge:
            return 40
        if "ROBBERY" in charge:
            return 43
        if "BURGLARY" in charge:
            return 44
        if "THEFT" in charge:
            return 45
        if "DECEPTION" in charge or "FRAUD" in charge:
            return 46
        if "HANDLING" in charge:
            return 47
        if "DAMAGE" in charge:
            return 48
        if "ARSON" in charge:
            return 49
        if "FORGERY" in charge:
            return 52
        if "DRUG" in charge or "CANNABIS" in charge:
            return 54
        if "PIPE" in charge:
            return 57
        if "VEHICLE" in charge:
            return 58
        if "TRAFFIC" in charge or "MOTOR" in charge or "LICENSE" in charge:
            return 59
        if "FIREARM" in charge or "AMMUNITION" in charge:
            return 59
        return 59

    def classify_statutory_sheet8(charge):
        c = str(charge).upper()
        if "DRUG" in c or "CANNABIS" in c:
            return 12
        if "FIREARM" in c or "AMMUNITION" in c:
            return 13
        if "LIQUOR" in c:
            return 14
        if "POLICE" in c:
            return 15
        if "GAMBLING" in c:
            return 16
        if "TRAFFIC" in c or "MOTOR" in c or "LICENSE" in c:
            return 17
        return 18

    def parse_disposition(remark):
        r = str(remark).upper()
        if any(x in r for x in ["CONVICTED", "GUILTY", "FINE", "PRISON"]):
            return "CONVICTED"
        if any(x in r for x in ["ACQUITTED", "DISMISSED", "STRUCK", "DISCHARGED"]):
            return "DISMISSED"
        if any(x in r for x in ["WITHDRAWN", "NOLLE"]):
            return "NOLLE"
        return "OTHER"

    def parse_sentence(sentence_text):
        s = str(sentence_text).upper()
        if "FINE" in s or "$" in s:
            return "FINE"
        if any(x in s for x in ["PRISON", "IMPRISONMENT", "CONFINEMENT", "MONTHS", "YEARS"]):
            return "PRISON"
        if "PROBATION" in s or "BOND" in s:
            return "PROBATION"
        if "REFORM" in s or "SCHOOL" in s:
            return "REFORMATORY"
        return "OTHER"

    def is_juvenile(age):
        try:
            return int(age) <= 16
        except Exception:
            return False

    def get_age_col_sheet5(age, gender):
        g = str(gender).upper()
        is_male = "F" not in g
        try:
            a = int(age)
        except Exception:
            return None
        if a <= 16:
            return "B" if is_male else "C"
        if 17 <= a <= 25:
            return "D" if is_male else "E"
        if 26 <= a <= 35:
            return "F" if is_male else "G"
        if 36 <= a <= 45:
            return "H" if is_male else "I"
        if a >= 46:
            return "J" if is_male else "K"
        return None

    # --- 3. TEMPLATE FILLER ---
    def fill_all_sheets(template_file, df, mode):
        wb = openpyxl.load_workbook(template_file)

        seen_cases = set()
        rows_sheet1 = []
        rows_sheet3 = []

        for idx, row in df.iterrows():
            r_num = classify_crime_sheet1(row.get("CHARGE", ""), row.get("VICTIM", ""))
            rows_sheet3.append(r_num)

            case_id = row.get("CASEID", idx)
            if case_id not in seen_cases:
                rows_sheet1.append(r_num)
                seen_cases.add(case_id)

        if "Sheet1" in wb.sheetnames:
            ws = wb["Sheet1"]
            col = "D" if mode == "New" else "J"
            for r in rows_sheet1:
                try:
                    curr = ws[f"{col}{r}"].value or 0
                    ws[f"{col}{r}"] = curr + 1
                except Exception:
                    pass

        if "Sheet3" in wb.sheetnames:
            ws = wb["Sheet3"]
            col = "D" if mode == "New" else "J"
            for r in rows_sheet3:
                try:
                    curr = ws[f"{col}{r}"].value or 0
                    ws[f"{col}{r}"] = curr + 1
                except Exception:
                    pass

        if "Sheet8" in wb.sheetnames:
            ws = wb["Sheet8"]
            for idx, row in df.iterrows():
                stat_row = classify_statutory_sheet8(row.get("CHARGE", ""))
                if mode == "New":
                    try:
                        ws[f"C{stat_row}"] = (ws[f"C{stat_row}"].value or 0) + 1
                    except Exception:
                        pass
                elif mode == "Disposed":
                    disp = parse_disposition(row.get("REMARK", ""))
                    if disp == "CONVICTED":
                        try:
                            ws[f"E{stat_row}"] = (ws[f"E{stat_row}"].value or 0) + 1
                        except Exception:
                            pass
                    elif disp == "DISMISSED":
                        try:
                            ws[f"F{stat_row}"] = (ws[f"F{stat_row}"].value or 0) + 1
                        except Exception:
                            pass

        if mode == "Disposed":
            if "Sheet2" in wb.sheetnames:
                ws = wb["Sheet2"]
                for idx, row in df.iterrows():
                    r_num = classify_crime_sheet1(row.get("CHARGE", ""), row.get("VICTIM", ""))
                    disp = parse_disposition(row.get("REMARK", ""))
                    target_col = {"CONVICTED": "E", "DISMISSED": "C", "NOLLE": "D"}.get(disp)
                    if target_col:
                        try:
                            ws[f"{target_col}{r_num}"] = (ws[f"{target_col}{r_num}"].value or 0) + 1
                        except Exception:
                            pass

            for idx, row in df.iterrows():
                if parse_disposition(row.get("REMARK", "")) != "CONVICTED":
                    continue

                r_num = classify_crime_sheet1(row.get("CHARGE", ""), row.get("VICTIM", ""))
                gender = row.get("GENDER", "M")
                is_male = "F" not in str(gender).upper()
                age = row.get("AGE", 0)
                sent_type = parse_sentence(row.get("SENTENCE", ""))

                if "Sheet4" in wb.sheetnames:
                    ws = wb["Sheet4"]
                    s_col = {
                        "PRISON": "D" if is_male else "E",
                        "PROBATION": "F" if is_male else "G",
                        "FINE": "H" if is_male else "I",
                    }.get(sent_type)
                    if s_col:
                        try:
                            ws[f"{s_col}{r_num}"] = (ws[f"{s_col}{r_num}"].value or 0) + 1
                        except Exception:
                            pass

                if "Sheet5" in wb.sheetnames:
                    ws = wb["Sheet5"]
                    a_col = get_age_col_sheet5(age, gender)
                    if a_col:
                        try:
                            ws[f"{a_col}{r_num - 11}"] = (ws[f"{a_col}{r_num - 11}"].value or 0) + 1
                        except Exception:
                            pass

                if is_juvenile(age):
                    juv_row = r_num - 14
                    if "Sheet6" in wb.sheetnames:
                        try:
                            ws_6 = wb["Sheet6"]
                            ws_6[f"F{juv_row}"] = (ws_6[f"F{juv_row}"].value or 0) + 1
                        except Exception:
                            pass
                    if "Sheet7" in wb.sheetnames:
                        ws = wb["Sheet7"]
                        sent_col = {
                            "PRISON": "B",
                            "PROBATION": "C",
                            "FINE": "D",
                            "REFORMATORY": "E",
                        }.get(sent_type)
                        if sent_col:
                            try:
                                ws[f"{sent_col}{juv_row}"] = (ws[f"{sent_col}{juv_row}"].value or 0) + 1
                            except Exception:
                                pass

                if "Sheet9" in wb.sheetnames:
                    stat_row = classify_statutory_sheet8(row.get("CHARGE", ""))
                    ws = wb["Sheet9"]
                    s_col = {
                        "PRISON": "D" if is_male else "E",
                        "PROBATION": "B" if is_male else "C",
                        "FINE": "F" if is_male else "G",
                    }.get(sent_type)
                    if s_col:
                        try:
                            ws[f"{s_col}{stat_row}"] = (ws[f"{s_col}{stat_row}"].value or 0) + 1
                        except Exception:
                            pass

        return wb

    # --- 4. COURT REPORT SIDEBAR + MAIN UI ---
    st.sidebar.header("1. Uploads")
    data_file = st.sidebar.file_uploader("Data File (Excel)", type=["xlsx"])
    template_file = st.sidebar.file_uploader("Blank Template", type=["xlsx"])

    st.sidebar.header("2. Settings")
    mode = st.sidebar.radio(
        "Data Type", ["New Cases (Arraignments)", "Disposed Cases (Concluded)"]
    )
    is_full_year = st.sidebar.checkbox("Full Year Report")

    if not is_full_year:
        report_month = st.sidebar.selectbox(
            "Month",
            range(1, 13),
            format_func=lambda x: datetime(2025, x, 1).strftime("%B"),
        )
    report_year = st.sidebar.number_input("Year", value=2025)

    if st.button("🚀 Process & Fill Report"):
        if not data_file or not template_file:
            st.error("Upload both files first.")
            st.stop()

        df = smart_read_excel(data_file)
        if df is None:
            st.stop()

        date_col = "DATE_ARR" if mode.startswith("New") else "DATE_DISP"
        if date_col not in df.columns:
            st.error(
                f"Missing Date Column. Need "
                f"'{'Arraignment' if mode.startswith('New') else 'Concluded'}' date."
            )
            st.stop()

        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

        if is_full_year:
            mask = df[date_col].dt.year == report_year
            period_name = f"Full Year {report_year}"
        else:
            mask = (df[date_col].dt.month == report_month) & (
                df[date_col].dt.year == report_year
            )
            period_name = datetime(2025, report_month, 1).strftime("%B %Y")

        df_filtered = df[mask].copy()
        st.success(f"Processing {len(df_filtered)} records for {period_name}")

        try:
            wb_filled = fill_all_sheets(
                template_file,
                df_filtered,
                "New" if mode.startswith("New") else "Disposed",
            )

            out = io.BytesIO()
            wb_filled.save(out)
            out.seek(0)

            st.download_button(
                "📥 Download Complete 9-Sheet Report",
                data=out,
                file_name=f"San_Pedro_Stats_9SHEETS_{period_name.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.write("### Data Preview")
            st.dataframe(df_filtered.head(50))

        except Exception as e:
            st.error(f"Processing Error: {e}")


# ============================================================================
# TAB 2 — ADVERTISING AGENT
# ============================================================================
with tab_ads:
    st.title("📢 Advertising Agent")
    st.markdown(
        """
AI-powered marketing assistant for **any business** — powered by **Claude Opus 4.7**.

Generate professional advertising content, campaign strategies, social media posts,
email campaigns, Google ad copy, and more — tailored to your industry and goals.
"""
    )

    # ── Import the agent (deferred so Tab 1 works even without anthropic) ──
    try:
        from advertising_agent import run_advertising_agent, ADVERTISING_TOOLS
        agent_available = True
    except ImportError:
        agent_available = False
        st.error(
            "The `anthropic` package is not installed. "
            "Run `pip install anthropic` and restart the app."
        )

    if agent_available:
        # ── API key ─────────────────────────────────────────────────────────
        env_key = os.environ.get("ANTHROPIC_API_KEY", "")

        with st.expander("🔑 API Key Settings", expanded=not bool(env_key)):
            if env_key:
                st.success("✅ API key loaded from environment variable `ANTHROPIC_API_KEY`.")
                api_key = env_key
            else:
                api_key = st.text_input(
                    "Anthropic API Key",
                    type="password",
                    placeholder="sk-ant-...",
                    help=(
                        "Your key is used only for this session and never stored. "
                        "You can also set the ANTHROPIC_API_KEY environment variable."
                    ),
                )
                if api_key:
                    st.success("✅ API key entered.")
                else:
                    st.info("Enter your Anthropic API key to use the advertising agent.")

        # ── Business Profile ─────────────────────────────────────────────────
        with st.expander("🏢 Business Profile (optional — improves results)", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                biz_name = st.text_input("Business Name", placeholder="e.g. Bloom Bakery")
                biz_industry = st.text_input("Industry / Business Type", placeholder="e.g. Bakery, SaaS, Fitness Studio")
                biz_location = st.text_input("Location / Market", placeholder="e.g. Austin, TX or Online")
            with col2:
                biz_services = st.text_input(
                    "Products / Services",
                    placeholder="e.g. Custom cakes, pastries, catering",
                )
                biz_target = st.text_input(
                    "Target Customers",
                    placeholder="e.g. Local families, event planners, corporate clients",
                )
                biz_differentiator = st.text_input(
                    "Unique Selling Point",
                    placeholder="e.g. 100% organic ingredients, same-day delivery",
                )

            business_context = {
                "Business Name": biz_name,
                "Industry / Business Type": biz_industry,
                "Location / Market": biz_location,
                "Products / Services": biz_services,
                "Target Customers": biz_target,
                "Unique Selling Point": biz_differentiator,
            }

        # ── Quick-action buttons ─────────────────────────────────────────────
        st.markdown("### ⚡ Quick Actions")
        st.markdown("Click a button to pre-fill the request field, or write your own below.")

        quick_actions = [
            ("📸 Instagram Post", "Write an engaging Instagram post for my business. Include a strong hook, compelling body copy, relevant hashtags, and a clear call-to-action."),
            ("📧 Promotional Email", "Draft a promotional email campaign to send to my existing customer list. Include a subject line, preview text, body copy with a special offer, and a call-to-action button."),
            ("📅 Launch Campaign", "Create a complete 4-week product/service launch campaign plan. Include strategy, weekly content calendar, channel recommendations, budget breakdown, and success KPIs."),
            ("🔍 Google Ads Copy", "Write Google Search Ad copy for my business. Provide 5 headlines (max 30 chars each) and 3 descriptions (max 90 chars each) that drive clicks and conversions."),
            ("🌐 Website Homepage", "Write compelling website homepage copy for my business, including a hero headline, subheadline, key benefits section, social proof section, and a strong call-to-action."),
            ("🛍️ Product Description", "Write persuasive product/service descriptions for my business that highlight key benefits, address customer pain points, and motivate purchases."),
            ("🎯 A/B Test Plan", "Help me design an A/B test for my marketing content. Create 3 variations with different angles and explain how to run the test, measure results, and pick a winner."),
            ("👥 Audience Analysis", "Analyze my target audience and create detailed buyer personas for my business. Include demographics, psychographics, pain points, buying triggers, and the best channels to reach them."),
        ]

        # Display in a grid
        cols_per_row = 4
        for i in range(0, len(quick_actions), cols_per_row):
            row_cols = st.columns(cols_per_row)
            for j, (label, prompt) in enumerate(quick_actions[i : i + cols_per_row]):
                with row_cols[j]:
                    if st.button(label, use_container_width=True, key=f"qa_{i + j}"):
                        st.session_state["ads_request"] = prompt

        # ── Request input ────────────────────────────────────────────────────
        st.markdown("### ✍️ Your Request")
        user_request = st.text_area(
            "Describe what you need:",
            value=st.session_state.get("ads_request", ""),
            height=120,
            placeholder=(
                "Examples:\n"
                "• Write an Instagram caption for my new product launch\n"
                "• Plan a 6-week campaign to grow my email list by 500 subscribers\n"
                "• Create Google ad copy for my local coffee shop\n"
                "• Draft a re-engagement email for customers who haven't bought in 90 days"
            ),
            key="ads_request_input",
        )

        # ── Generate button ──────────────────────────────────────────────────
        col_gen, col_clear = st.columns([3, 1])
        with col_gen:
            generate_clicked = st.button(
                "🚀 Generate Content",
                type="primary",
                use_container_width=True,
                disabled=not api_key or not user_request.strip(),
            )
        with col_clear:
            if st.button("🗑️ Clear History", use_container_width=True):
                st.session_state["ads_history"] = []
                st.rerun()

        if not api_key and not user_request.strip():
            st.info("Enter your API key and describe what you need to get started.")
        elif not api_key:
            st.warning("⚠️ Please enter your Anthropic API key above.")
        elif not user_request.strip():
            st.info("📝 Describe what advertising content or strategy you need.")

        # ── Run the agent ────────────────────────────────────────────────────
        if generate_clicked and api_key and user_request.strip():
            st.markdown("---")
            st.markdown("### 🤖 Agent Response")

            response_placeholder = st.empty()
            full_response = ""
            start_time = time.time()

            try:
                with st.spinner("Agent is working…"):
                    for chunk in run_advertising_agent(
                        api_key=api_key,
                        user_request=user_request,
                        business_context=business_context,
                    ):
                        full_response += chunk
                        response_placeholder.markdown(full_response)

                elapsed = time.time() - start_time
                st.caption(f"⏱️ Generated in {elapsed:.1f} seconds")

                # Save to session history
                if "ads_history" not in st.session_state:
                    st.session_state["ads_history"] = []

                st.session_state["ads_history"].insert(
                    0,
                    {
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "request": user_request,
                        "response": full_response,
                    },
                )

                # Download button for the generated content
                st.download_button(
                    "📥 Download as Markdown",
                    data=f"# Advertising Agent Output\n\n**Request:** {user_request}\n\n---\n\n{full_response}",
                    file_name=f"advertising_content_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
                    mime="text/markdown",
                )

            except Exception as e:
                err_str = str(e)
                if "authentication" in err_str.lower() or "api_key" in err_str.lower() or "401" in err_str:
                    st.error("❌ Invalid API key. Please check your Anthropic API key and try again.")
                elif "rate_limit" in err_str.lower() or "429" in err_str:
                    st.error("❌ Rate limit reached. Please wait a moment and try again.")
                elif "overloaded" in err_str.lower() or "529" in err_str:
                    st.error("❌ The API is temporarily overloaded. Please try again in a few seconds.")
                else:
                    st.error(f"❌ An error occurred: {err_str}")

        # ── Content History ──────────────────────────────────────────────────
        if st.session_state.get("ads_history"):
            st.markdown("---")
            st.markdown("### 📋 Session History")
            st.caption(f"{len(st.session_state['ads_history'])} item(s) generated this session")

            for i, item in enumerate(st.session_state["ads_history"]):
                with st.expander(
                    f"[{item['timestamp']}] {item['request'][:80]}{'…' if len(item['request']) > 80 else ''}",
                    expanded=(i == 0),
                ):
                    st.markdown(item["response"])
                    st.download_button(
                        "📥 Download",
                        data=(
                            f"# Advertising Agent Output\n\n"
                            f"**Request:** {item['request']}\n\n---\n\n{item['response']}"
                        ),
                        file_name=(
                            f"advertising_{item['timestamp'].replace(':', '').replace(' ', '_')}.md"
                        ),
                        mime="text/markdown",
                        key=f"dl_{i}",
                    )

        # ── Tips sidebar panel ───────────────────────────────────────────────
        with st.sidebar:
            st.markdown("---")
            st.markdown("### 📢 Advertising Agent Tips")
            st.markdown(
                """
**For best results:**
- Fill in the Business Profile above — the agent uses this to tailor content to your specific business
- Be specific in your requests (platform, audience, goal, tone)
- Use Quick Actions as a starting point and customize the pre-filled request before generating

**What the agent can do:**
- ✅ Instagram, TikTok, LinkedIn, Facebook posts
- ✅ Email campaigns & promotional sequences
- ✅ Google & Meta ad copy
- ✅ Website & landing page copy
- ✅ Full launch & growth campaign strategies
- ✅ A/B testing plans & frameworks
- ✅ Audience analysis & buyer personas
- ✅ Content calendars & brand voice guides
"""
            )
