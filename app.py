import streamlit as st
import pandas as pd
import pdfplumber
from docx import Document # Requires python-docx in requirements.txt
from io import BytesIO

# --- CONFIGURATION ---
st.set_page_config(page_title="Ultimate Excel Automator", layout="wide")

def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Enter Beta Access Password", type="password", on_change=password_entered, key="password")
        st.info("Direct message @Gewish on Reddit for access.")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Enter Beta Access Password", type="password", on_change=password_entered, key="password")
        st.error("😕 Password incorrect")
        return False
    else:
        return True

if check_password():
    # --- SIDEBAR: COMMUNITY & SUPPORT ---
    with st.sidebar:
        st.header("🛠️ Tool Belt")
        if st.button("♻️ Reset App & Clear Data"):
            st.rerun()
        
        st.divider()
        st.header("📣 Support & Feedback")
        st.link_button("🪲 Report a Bug", "https://forms.gle/rVE2KkorZX4iqWNq7")
        st.link_button("☕ Buy Me A Coffee", "https://paypal.me/GewishCatedrilla")
        
        st.divider()
        st.markdown("**Built for the r/excel community.**")

    st.title("🛠️ The Excel Swiss Army Knife")
    st.markdown("Automating the headaches you'd normally need 100 formulas to fix.")

    uploaded_files = st.file_uploader("Upload Excel, CSV, PDF, or Word", 
                                     type=["xlsx", "csv", "pdf", "docx"], 
                                     accept_multiple_files=True)

    df = pd.DataFrame()

    if uploaded_files:
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
        
        if df_list:
            df = pd.concat(df_list, ignore_index=True)
            
            # --- CRITICAL FIX: Header Sanitization ---
            # This converts all column names to strings to stop the Arrow warning
            df.columns = [str(c) for c in df.columns]
            
            # Auto-detect if row 0 is actually the header
            if not df.empty and any(df.iloc[0].isna()) == False:
                df.columns = [str(c) for c in df.iloc[0]]
                df = df[1:].reset_index(drop=True)

        if not df.empty:
            st.write("### 🔍 Live Data Preview", df.head(10).astype(str))
            
            tabs = st.tabs([
                "🎯 Mapper", "🧹 Cleaner", "🕵️ Detective", "⏰ Time Math", 
                "📊 Health", "🔄 Shifter", "✅ Validator", "📂 Word-to-Excel"
            ])

            # TAB 1: LOGIC MAPPER
            with tabs[0]:
                st.header("Categorization & Logic")
                col_to_check = st.selectbox("Column to scan:", df.columns, key="map_col")
                keyword = st.text_input("Find this keyword:")
                category_val = st.text_input("Assign this value:")
                if st.button("Apply Logic"):
                    if 'New_Category' not in df.columns: df['New_Category'] = "Uncategorized"
                    df.loc[df[col_to_check].astype(str).str.contains(keyword, case=False, na=False), 'New_Category'] = category_val
                    st.success("Logic applied!")

            # TAB 2: CLEANER (Scientific Notation Fix)
            with tabs[1]:
                st.header("The Formatting Fixer")
                clean_col = st.selectbox("Select Column:", df.columns, key="clean_tab")
                c_opt = st.radio("Fix Type:", ["Fix Scientific Notation (0.000077)", "Remove Symbols ($, -, %)", "Trim Spaces"])
                if st.button("Run Cleaner"):
                    if "Scientific" in c_opt:
                        df[clean_col] = df[clean_col].astype(str).apply(lambda x: '{:.10f}'.format(float(x)).rstrip('0').rstrip('.') if 'E' in str(x) else x)
                    elif "Symbols" in c_opt:
                        df[clean_col] = df[clean_col].astype(str).str.replace(r'[$\-,%]', '', regex=True)
                    elif "Spaces" in c_opt:
                        df[clean_col] = df[clean_col].astype(str).str.strip()
                    st.success("Column cleaned!")

            # TAB 3: DUPLICATE DETECTIVE
            with tabs[2]:
                st.header("Duplicate Detective")
                match_cols = st.multiselect("Check for matches in these columns:", df.columns)
                if st.button("Find Duplicates"):
                    df['Is_Duplicate'] = df.duplicated(subset=match_cols, keep=False)
                    dupes = df[df['Is_Duplicate']]
                    if not dupes.empty:
                        st.warning(f"🚨 Found {len(dupes)} matching rows!")
                        st.dataframe(dupes.astype(str))
                    else:
                        st.success("No duplicates found!")

            # TAB 4: TIME CALCULATOR
            with tabs[3]:
                st.header("Time & Duration Math")
                t1 = st.selectbox("Start Time:", df.columns, key="t1")
                t2 = st.selectbox("End Time:", df.columns, key="t2")
                if st.button("Calculate Minutes"):
                    df['Total_Minutes'] = (pd.to_datetime(df[t2], errors='coerce') - pd.to_datetime(df[t1], errors='coerce')).dt.total_seconds() / 60
                    st.success("Calculated durations!")

            # TAB 5: DATA HEALTH
            with tabs[4]:
                st.header("Data Health Audit")
                st.write(f"**Total Records:** {len(df)}")
                st.write("**Null/Empty Values:**")
                st.write(df.isnull().sum())

            # TAB 6: SHIFTER
            with tabs[5]:
                st.header("Format Shifter")
                shift_opt = st.selectbox("Export to:", ["Word Table Style (CSV)", "HTML Report", "JSON"])
                if st.button("Process Shifter"):
                    if "Word" in shift_opt:
                        st.download_button("📥 Download for Word", df.to_csv(index=False), "word_data.csv")
                    elif "HTML" in shift_opt:
                        st.download_button("📥 Download HTML", df.to_html(), "report.html")

            # TAB 7: VALIDATOR
            with tabs[6]:
                st.header("Rule Validator")
                v_col = st.selectbox("Column to Check:", df.columns, key="val_col")
                v_type = st.selectbox("Validation Rule:", ["Must be a Number", "Must be an Email", "Cannot be Empty"])
                if st.button("Validate Now"):
                    if "Number" in v_type:
                        # Clean up strings that look like numbers before checking
                        test_col = df[v_col].astype(str).str.replace(',','').str.replace('$','')
                        errors = df[~test_col.str.replace('.','',1).isdigit()]
                        if not errors.empty:
                            st.error(f"Found {len(errors)} invalid numbers!")
                            st.dataframe(errors)
                        else:
                            st.success("All numbers are valid!")

            # TAB 8: WORD-TO-EXCEL
            with tabs[7]:
                st.header("Word Table Import")
                st.info("Any tables found in Word docs are already in the Live Preview above.")

            # --- FINAL DOWNLOAD ---
            st.divider()
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 DOWNLOAD CLEAN EXCEL FILE", output.getvalue(), "fixed_data.xlsx", use_container_width=True)
