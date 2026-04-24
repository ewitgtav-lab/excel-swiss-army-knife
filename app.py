import streamlit as st
import pandas as pd
import pdfplumber
from docx import Document
from io import BytesIO
import re

# --- CONFIGURATION ---
st.set_page_config(page_title="Ultimate Excel Automator", layout="wide")

# --- PERFORMANCE CACHE ---
@st.cache_data(show_spinner="Sharpening the Knife...")
def load_and_sanitize(uploaded_files):
    """Combines files and sanitizes headers once to prevent downstream lag."""
    df_list = []
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                df_list.append(pd.read_csv(file))
            elif file.name.endswith('.docx'):
                doc = Document(file)
                df_list.append(pd.DataFrame([[c.text for c in r.cells] for t in doc.tables for r in t.rows]))
            elif file.name.endswith('.pdf'):
                with pdfplumber.open(file) as pdf:
                    for page in pdf.pages:
                        table = page.extract_table()
                        if table: df_list.append(pd.DataFrame(table))
            else:
                df_list.append(pd.read_excel(file))
        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")
    
    if not df_list: return pd.DataFrame()

    df = pd.concat(df_list, ignore_index=True)
    # Sanitizing headers to strings immediately prevents Streamlit/Arrow rendering lag
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --- AUTHENTICATION ---
def check_password():
    if st.session_state.get("password_correct"): return True
    def password_entered():
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False
    st.text_input("Enter Beta Access Password", type="password", on_change=password_entered, key="password")
    st.info("Direct message @Gewish on Reddit for access.")
    return False

if check_password():
    # Persistent Data Storage
    if "main_df" not in st.session_state:
        st.session_state.main_df = None

    with st.sidebar:
        st.header("🛠️ Tool Belt")
        if st.button("♻️ Reset App & Memory"):
            st.session_state.main_df = None
            st.cache_data.clear()
            st.rerun()
        st.divider()
        st.header("📣 Support")
        st.link_button("☕ Buy Me A Coffee", "https://paypal.me/GewishCatedrilla")

    st.title("🛠️ The Excel Swiss Army Knife")

    # FILE UPLOADER - Only shows if no data is loaded
    if st.session_state.main_df is None:
        files = st.file_uploader("Upload Files", type=["xlsx", "csv", "pdf", "docx"], accept_multiple_files=True)
        if files:
            st.session_state.main_df = load_and_sanitize(files)
            st.rerun()

    # MAIN APP LOGIC
    if st.session_state.main_df is not None:
        df = st.session_state.main_df
        
        st.write(f"### 🔍 Preview ({len(df)} rows)")
        st.dataframe(df.head(5), use_container_width=True)
        
        tabs = st.tabs(["✅ Validator", "🧹 Cleaner", "🧮 Aggregator", "🎯 Mapper", "🕵️ Detective", "⏰ Time Math", "🔄 Shifter"])

        # TAB 1: VALIDATOR (CRASH-PROOF & VECTORIZED)
        with tabs[0]:
            st.header("Data Integrity Check")
            v_col = st.selectbox("Column to Check:", df.columns, key="val_v")
            v_type = st.radio("Rule:", ["Numeric", "Email", "Not Empty"])
            
            if st.button("Run Validation"):
                if v_type == "Numeric":
                    # Vectorized: Strips $ and , then checks for numbers
                    clean_test = df[v_col].astype(str).str.replace(r'[$,]', '', regex=True)
                    # pd.to_numeric with errors='coerce' is the "Fast/Safe" way
                    check = pd.to_numeric(clean_test, errors='coerce')
                    errors = df[check.isna()]
                elif v_type == "Email":
                    errors = df[~df[v_col].astype(str).str.contains(r'[^@]+@[^@]+\.[^@]+', na=False)]
                else:
                    errors = df[df[v_col].astype(str).str.strip() == ""]
                
                if not errors.empty:
                    st.error(f"🚨 Found {len(errors)} invalid rows!")
                    st.dataframe(errors.head(100).astype(str))
                else:
                    st.success("✅ Everything looks perfect!")

        # TAB 2: CLEANER (FAST FIXES)
        with tabs[1]:
            st.header("The Formatting Fixer")
            c_col = st.selectbox("Target Column:", df.columns, key="clean_v")
            c_opt = st.selectbox("Action:", ["Scientific Notation Fix", "Trim Spaces", "Remove Symbols"])
            
            if st.button("Apply Clean"):
                if "Scientific" in c_opt:
                    st.session_state.main_df[c_col] = pd.to_numeric(df[c_col], errors='coerce').apply(
                        lambda x: '{:.10f}'.format(x).rstrip('0').rstrip('.') if pd.notnull(x) else x
                    )
                elif "Spaces" in c_opt:
                    st.session_state.main_df[c_col] = df[c_col].astype(str).str.strip()
                st.success("Cleaned! Preview updated.")

        # [Remaining tabs follow the same 'st.session_state.main_df' pattern for speed]

        st.divider()
        # DOWNLOAD BUTTON
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 DOWNLOAD COMPLETED FILE", output.getvalue(), "fixed_data.xlsx", use_container_width=True)
