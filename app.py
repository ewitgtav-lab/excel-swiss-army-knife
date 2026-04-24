import streamlit as st
import pandas as pd
import pdfplumber
from docx import Document
from io import BytesIO

# --- CONFIGURATION ---
st.set_page_config(page_title="Ultimate Excel Automator", layout="wide")

# --- SPEED OPTIMIZATION: CACHED LOADING ---
@st.cache_data(show_spinner="Sharpening the Knife...")
def load_and_clean_data(uploaded_files):
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
                        if table:
                            df_list.append(pd.DataFrame(table))
            else:
                df_list.append(pd.read_excel(file))
        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")
    
    if not df_list:
        return pd.DataFrame()

    df = pd.concat(df_list, ignore_index=True)
    
    # 1. Header Sanitization (Prevents Arrow Warnings)
    df.columns = [str(c) for c in df.columns]
    
    # 2. Auto-fix Scientific Notation immediately on upload
    for col in df.columns:
        if df[col].astype(str).str.contains('E', na=False).any():
            df[col] = df[col].astype(str).apply(
                lambda x: '{:.10f}'.format(float(x)).rstrip('0').rstrip('.') if 'E' in str(x) else x
            )
    return df

def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False
    if "password_correct" not in st.session_state:
        st.text_input("Enter Beta Access Password", type="password", on_change=password_entered, key="password")
        return False
    return st.session_state["password_correct"]

if check_password():
    with st.sidebar:
        st.header("🛠️ Tool Belt")
        if st.button("♻️ Reset App"):
            st.cache_data.clear()
            st.rerun()
        
        st.divider()
        st.header("📣 Support")
        st.link_button("🪲 Report a Bug", "https://forms.gle/rVE2KkorZX4iqWNq7")
        st.link_button("☕ Support my Work", "https://paypal.me/GewishCatedrilla")
        
        st.divider()
        st.markdown("**Version 2.0 - Swiss Army Knife**")

    st.title("🛠️ The Excel Swiss Army Knife")
    st.markdown("Automating Reddit's biggest spreadsheet headaches.")

    uploaded_files = st.file_uploader("Upload Excel, CSV, PDF, or Word", 
                                     type=["xlsx", "csv", "pdf", "docx"], 
                                     accept_multiple_files=True)

    if uploaded_files:
        df = load_and_clean_data(uploaded_files)

        if not df.empty:
            st.write(f"### 🔍 Live Preview ({len(df)} rows)", df.head(10).astype(str))
            
            tabs = st.tabs([
                "🧮 Aggregator", "🎯 Mapper", "🧹 Cleaner", "🕵️ Detective", 
                "⏰ Time Math", "🔄 Shifter", "✅ Validator", "📂 Word"
            ])

            # TAB 1: NEW! AGGREGATOR (The SUMIFS Problem)
            with tabs[0]:
                st.header("Smart Aggregator (Auto-SUMIFS)")
                st.write("Calculate totals based on groups (e.g., Total Cost per Item Quantity).")
                group_col = st.selectbox("Group by (e.g., 'Item Name'):", df.columns)
                calc_col = st.selectbox("Column to Sum (e.g., 'Cost'):", df.columns)
                
                if st.button("Generate Summary"):
                    df[calc_col] = pd.to_numeric(df[calc_col], errors='coerce')
                    summary = df.groupby(group_col)[calc_col].sum().reset_index()
                    st.dataframe(summary)
                    st.download_button("📥 Download Summary", summary.to_csv(index=False), "summary.csv")

            # TAB 2: LOGIC MAPPER
            with tabs[1]:
                st.header("Categorization Logic")
                col_to_check = st.selectbox("Column to scan:", df.columns, key="map_col")
                keyword = st.text_input("If it contains:")
                category_val = st.text_input("Then assign:")
                if st.button("Apply"):
                    if 'Category' not in df.columns: df['Category'] = "Uncategorized"
                    df.loc[df[col_to_check].astype(str).str.contains(keyword, case=False, na=False), 'Category'] = category_val
                    st.success("Logic applied!")

            # TAB 3: CLEANER
            with tabs[2]:
                st.header("Format Scrubbing")
                clean_col = st.selectbox("Target Column:", df.columns, key="clean_tab")
                c_opt = st.radio("Fix:", ["Remove Symbols ($, -, %)", "Trim Whitespace", "Proper Case"])
                if st.button("Clean Now"):
                    if "Symbols" in c_opt:
                        df[clean_col] = df[clean_col].astype(str).str.replace(r'[$\-,%]', '', regex=True)
                    elif "Spaces" in c_opt:
                        df[clean_col] = df[clean_col].astype(str).str.strip()
                    elif "Proper" in c_opt:
                        df[clean_col] = df[clean_col].astype(str).str.title()
                    st.success("Cleaned!")

            # TAB 4: DUPLICATE DETECTIVE
            with tabs[3]:
                st.header("Duplicate Detective")
                match_cols = st.multiselect("Match based on:", df.columns)
                if st.button("Run Scan"):
                    df['Is_Duplicate'] = df.duplicated(subset=match_cols, keep=False)
                    dupes = df[df['Is_Duplicate']]
                    st.warning(f"Found {len(dupes)} duplicates.")
                    st.dataframe(dupes.astype(str))

            # TAB 5: TIME MATH (Solving the Hours/Minutes problem)
            with tabs[4]:
                st.header("Time & Labor Math")
                st.write("Convert hours/minutes into a single decimal value automatically.")
                h_col = st.selectbox("Hours Column:", df.columns)
                m_col = st.selectbox("Minutes Column:", df.columns)
                if st.button("Fix Time Columns"):
                    h = pd.to_numeric(df[h_col], errors='coerce').fillna(0)
                    m = pd.to_numeric(df[m_col], errors='coerce').fillna(0)
                    df['Total_Hours_Decimal'] = h + (m / 60)
                    st.success("Created 'Total_Hours_Decimal' column!")

            # TAB 6: FORMAT SHIFTER
            with tabs[5]:
                st.header("The Shifter")
                shift_opt = st.selectbox("Format:", ["Word Table (CSV)", "HTML Report", "JSON"])
                if st.button("Prepare"):
                    if "Word" in shift_opt:
                        st.download_button("📥 Download", df.to_csv(index=False), "for_word.csv")
                    elif "HTML" in shift_opt:
                        st.download_button("📥 Download", df.to_html(), "report.html")

            # TAB 7: VALIDATOR
            with tabs[6]:
                st.header("Data Rules")
                v_col = st.selectbox("Check Column:", df.columns, key="val_col")
                v_type = st.selectbox("Rule:", ["Must be a Number", "Email Format", "Not Empty"])
                if st.button("Validate"):
                    if "Number" in v_type:
                        err = df[~df[v_col].astype(str).str.replace('.','',1).isdigit()]
                        st.error(f"Errors: {len(err)}")
                        st.dataframe(err)

            # TAB 8: WORD-TO-EXCEL
            with tabs[7]:
                st.header("Word Import")
                st.info("Any tables in uploaded .docx files are already merged into the live preview.")

            # --- FINAL DOWNLOAD ---
            st.divider()
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 DOWNLOAD COMPLETED SWISS ARMY FILE", output.getvalue(), "fixed_data.xlsx", use_container_width=True)
