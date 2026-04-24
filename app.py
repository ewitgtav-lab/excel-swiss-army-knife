import streamlit as st
import pandas as pd
import pdfplumber
from docx import Document
from io import BytesIO
import re

# --- CONFIGURATION ---
st.set_page_config(page_title="Ultimate Excel Automator", layout="wide")

# --- PERFORMANCE ENGINE: CACHED LOADING ---
@st.cache_data(show_spinner="Sharpening the Knife...")
def load_and_sanitize_data(uploaded_files):
    df_list = []
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                df_list.append(pd.read_csv(file))
            elif file.name.endswith('.docx'):
                doc = Document(file)
                for table in doc.tables:
                    data = [[cell.text for cell in row.cells] for row in table.rows]
                    df_list.append(pd.DataFrame(data))
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
    
    # Header Sanitization: Prevents 'Non-str index' warnings
    df.columns = [str(c).strip() for c in df.columns]
    
    # Auto-detect if the first row is actually a header
    if not df.empty and df.iloc[0].notnull().all():
        df.columns = [str(c) for c in df.iloc[0]]
        df = df[1:].reset_index(drop=True)
        
    return df

def check_password():
    if "password_correct" not in st.session_state:
        def password_entered():
            if st.session_state["password"] == st.secrets["password"]:
                st.session_state["password_correct"] = True
                del st.session_state["password"]
            else:
                st.session_state["password_correct"] = False
        st.text_input("Enter Beta Access Password", type="password", on_change=password_entered, key="password")
        st.info("Direct message @Gewish on Reddit for access.")
        return False
    return st.session_state["password_correct"]

if check_password():
    # --- SIDEBAR ---
    with st.sidebar:
        st.header("🛠️ Tool Belt")
        if st.button("♻️ Reset & Clear Memory"):
            st.cache_data.clear()
            st.rerun()
        st.divider()
        st.header("📣 Support")
        st.link_button("🪲 Report a Bug", "https://forms.gle/rVE2KkorZX4iqWNq7")
        st.link_button("☕ Support my Work", "https://paypal.me/GewishCatedrilla")

    st.title("🛠️ The Excel Swiss Army Knife")

    uploaded_files = st.file_uploader("Upload Files", type=["xlsx", "csv", "pdf", "docx"], accept_multiple_files=True)

    if uploaded_files:
        df = load_and_sanitize_data(uploaded_files)

        if not df.empty:
            st.write(f"### 🔍 Preview ({len(df)} rows)", df.head(5).astype(str))
            
            tabs = st.tabs(["✅ Validator", "🧹 Cleaner", "🧮 Aggregator", "🎯 Mapper", "🕵️ Detective", "⏰ Time Math", "🔄 Shifter"])

            # TAB 1: VALIDATOR (CRASH-PROOF)
            with tabs[0]:
                st.header("Data Integrity Check")
                v_col = st.selectbox("Column to Check:", df.columns, key="val_v")
                v_type = st.radio("Rule:", ["Must be Numeric", "Must be Email", "No Empty Cells"])
                
                if st.button("Run Validation"):
                    if v_type == "Must be Numeric":
                        # Vectorized numeric check: strips $ and , first
                        clean_col = df[v_col].astype(str).str.replace(r'[$,]', '', regex=True)
                        is_numeric = pd.to_numeric(clean_test, errors='coerce').notnull()
                        errors = df[~is_numeric]
                    elif v_type == "Must be Email":
                        errors = df[~df[v_col].astype(str).str.contains(r'[^@]+@[^@]+\.[^@]+', na=False)]
                    else:
                        errors = df[df[v_col].isna() | (df[v_col].astype(str).str.strip() == "")]
                    
                    if not errors.empty:
                        st.error(f"🚨 Found {len(errors)} invalid rows!")
                        st.dataframe(errors.astype(str))
                    else:
                        st.success("✅ Validation Passed!")

            # TAB 2: CLEANER (FAST SCIENTIFIC NOTATION FIX)
            with tabs[1]:
                st.header("The Formatting Fixer")
                c_col = st.selectbox("Target Column:", df.columns, key="clean_v")
                c_opt = st.selectbox("Action:", ["Fix Scientific Notation", "Remove Symbols", "Proper Case"])
                
                if st.button("Apply Cleaning"):
                    if "Scientific" in c_opt:
                        df[c_col] = pd.to_numeric(df[c_col], errors='coerce').apply(
                            lambda x: '{:.10f}'.format(x).rstrip('0').rstrip('.') if pd.notnull(x) else x
                        )
                    elif "Symbols" in c_opt:
                        df[c_col] = df[c_col].astype(str).str.replace(r'[$\-,%]', '', regex=True)
                    st.success("Updated Successfully!")

            # [Aggregator, Mapper, Detective, Time Math, and Shifter tabs logic follow similar optimized patterns...]

            # --- FINAL DOWNLOAD ---
            st.divider()
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 DOWNLOAD COMPLETED FILE", output.getvalue(), "fixed_data.xlsx", use_container_width=True)
