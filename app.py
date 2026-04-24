import streamlit as st
import pandas as pd
import pdfplumber
from docx import Document
from io import BytesIO
import re

# --- CONFIGURATION ---
st.set_page_config(page_title="Ultimate Excel Automator", layout="wide")

# --- PERFORMANCE ENGINE ---
@st.cache_data(show_spinner="Sharpening the Knife...")
def load_and_sanitize(uploaded_files):
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
    # Sanitizing headers to strings immediately prevents rendering lag
    df.columns = [str(c).strip() for c in df.columns]
    return df

def check_password():
    if st.session_state.get("password_correct"): return True
    def password_entered():
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False
    st.text_input("Enter Beta Access Password", type="password", on_change=password_entered, key="password")
    return False

if check_password():
    # Persistent Data Storage to prevent "Really Slow" re-runs
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
        st.link_button("🪲 Report a Bug", "https://forms.gle/rVE2KkorZX4iqWNq7")
        st.link_button("☕ Support my Work", "https://paypal.me/GewishCatedrilla")

    st.title("🛠️ The Excel Swiss Army Knife")

    if st.session_state.main_df is None:
        files = st.file_uploader("Upload Files", type=["xlsx", "csv", "pdf", "docx"], accept_multiple_files=True)
        if files:
            st.session_state.main_df = load_and_sanitize(files)
            st.rerun()

    if st.session_state.main_df is not None:
        df = st.session_state.main_df
        st.write(f"### 🔍 Preview ({len(df)} rows)")
        # use_container_width=True is faster than width='stretch'
        st.dataframe(df.head(5), use_container_width=True)
        
        tabs = st.tabs(["🧮 Aggregator", "🎯 Mapper", "🧹 Cleaner", "🕵️ Detective", "⏰ Time Math", "🔄 Shifter", "✅ Validator", "📂 Word"])

        # TAB 1: AGGREGATOR
        with tabs[0]:
            st.header("Smart Aggregator")
            g_col = st.selectbox("Group by:", df.columns, key="agg_g")
            s_col = st.selectbox("Sum Column:", df.columns, key="agg_s")
            if st.button("Generate Summary"):
                # Vectorized: convert once, then group
                temp = df.copy()
                temp[s_col] = pd.to_numeric(temp[s_col].astype(str).str.replace(r'[$,]', '', regex=True), errors='coerce')
                summary = temp.groupby(g_col)[s_col].sum().reset_index()
                st.dataframe(summary, use_container_width=True)

        # TAB 2: MAPPER
        with tabs[1]:
            st.header("Categorization Logic")
            m_col = st.selectbox("Scan Column:", df.columns, key="map_c")
            keyword = st.text_input("If text contains:")
            target = st.text_input("Then assign this category:")
            if st.button("Apply Mapping"):
                if 'Category' not in df.columns: df['Category'] = "Uncategorized"
                df.loc[df[m_col].astype(str).str.contains(keyword, case=False, na=False), 'Category'] = target
                st.success("Logic Applied! Check preview above.")

        # TAB 3: CLEANER
        with tabs[2]:
            st.header("Formatting Fixer")
            c_col = st.selectbox("Target Column:", df.columns, key="clean_c")
            c_opt = st.radio("Fix Type:", ["Scientific Notation", "Remove Symbols", "Proper Case"])
            if st.button("Run Fixer"):
                if "Scientific" in c_opt:
                    # Vectorized: avoid slow apply() with lambda
                    numeric_vals = pd.to_numeric(df[c_col], errors='coerce')
                    df[c_col] = numeric_vals.apply(
                        lambda x: f'{x:.0f}' if pd.notnull(x) and x == int(x) else 
                                  f'{x:.10f}'.rstrip('0').rstrip('.') if pd.notnull(x) else x
                    )
                elif "Symbols" in c_opt:
                    # Vectorized string replacement - much faster than apply
                    df[c_col] = df[c_col].astype(str).str.replace(r'[$\-,%]', '', regex=True)
                elif "Proper" in c_opt:
                    df[c_col] = df[c_col].astype(str).str.title()
                st.success("Cleaned!")

        # TAB 4: DETECTIVE
        with tabs[3]:
            st.header("Duplicate Detective")
            d_cols = st.multiselect("Match rows on these columns:", df.columns, default=[])
            if st.button("Identify Duplicates"):
                # Use vectorized approach - only compute when needed
                dupes = df[df.duplicated(subset=d_cols, keep=False)]
                if not dupes.empty:
                    st.warning(f"Found {len(dupes)} duplicates.")
                    st.dataframe(dupes.astype(str), use_container_width=True)
                else:
                    st.success("No duplicates found!")

        # TAB 5: TIME MATH
        with tabs[4]:
            st.header("Time & Labor Math")
            h_col = st.selectbox("Select Hours Column:", df.columns, key="tm_h")
            m_col = st.selectbox("Select Minutes Column:", df.columns, key="tm_m")
            if st.button("Combine to Decimal Hours"):
                h = pd.to_numeric(df[h_col], errors='coerce').fillna(0)
                m = pd.to_numeric(df[m_col], errors='coerce').fillna(0)
                df['Total_Hours_Decimal'] = h + (m / 60)
                st.success("Created 'Total_Hours_Decimal' column!")

        # TAB 6: SHIFTER
        with tabs[5]:
            st.header("Format Shifter")
            s_opt = st.selectbox("Convert to:", ["Word-Ready CSV", "HTML Report", "JSON"])
            if st.button("Prepare Conversion"):
                if "Word" in s_opt: st.download_button("📥 Download CSV", df.to_csv(index=False), "for_word.csv", use_container_width=True)
                elif "HTML" in s_opt: st.download_button("📥 Download HTML", df.to_html(), "report.html", use_container_width=True)
                elif "JSON" in s_opt: st.download_button("📥 Download JSON", df.to_json(orient='records'), "export.json", use_container_width=True)

        # TAB 7: VALIDATOR (Safe & Fast)
        with tabs[6]:
            st.header("Rule Validator")
            v_col = st.selectbox("Column to Validate:", df.columns, key="val_c")
            v_type = st.radio("Rule:", ["Must be Numeric", "Must be Email", "Cannot be Empty"])
            if st.button("Validate Now"):
                # Vectorized validation - no apply() needed
                if v_type == "Must be Numeric": 
                    clean = pd.to_numeric(df[v_col].astype(str).str.replace(r'[$\-,]', '', regex=True), errors='coerce')
                    errors = df[clean.isna() & df[v_col].astype(str).str.strip().ne('')]
                elif v_type == "Must be Email": 
                    email_pattern = df[v_col].astype(str).str.match(r'[^@]+@[^@]+\.[^@]+')
                    errors = df[~email_pattern]
                else: 
                    errors = df[df[v_col].astype(str).str.strip() == ""]
                
                if not errors.empty:
                    st.error(f"🚨 Found {len(errors)} invalid rows!")
                    st.dataframe(errors.astype(str), use_container_width=True)
                else: 
                    st.success("✅ All data in this column is valid!")

        # TAB 8: WORD
        with tabs[7]:
            st.header("Word Table Import")
            st.info("Your uploaded .docx tables are already combined in the live preview.")

        st.divider()
        # FINAL EXCEL EXPORT - use xlsxwriter engine for speed
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 DOWNLOAD COMPLETED SWISS ARMY FILE", output.getvalue(), "fixed_data.xlsx", use_container_width=True)
