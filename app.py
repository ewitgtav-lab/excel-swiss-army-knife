import streamlit as st
import pandas as pd
import pdfplumber
from docx import Document # Add python-docx to requirements.txt
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
    with st.sidebar:
        st.header("🛠️ Tool Belt")
        if st.button("♻️ Reset App & Clear Data"):
            st.rerun()
        st.divider()
        st.markdown("""
        **How to use:**
        1. **Upload** Excel, CSV, PDF, or Word.
        2. **Choose a tab** for your specific task.
        3. **Process & Download** the fixed result.
        """)
        st.divider()
        st.write("Built for the r/excel community by Gewish.")

    st.title("🛠️ The Excel Swiss Army Knife")
    st.markdown("One-click solutions for the most annoying spreadsheet tasks on Reddit.")

    uploaded_files = st.file_uploader("Upload your files (Excel, CSV, PDF, or Word)", 
                                     type=["xlsx", "csv", "pdf", "docx"], 
                                     accept_multiple_files=True)

    df = pd.DataFrame()

    if uploaded_files:
        # --- MULTI-FORMAT LOADING ENGINE ---
        df_list = []
        for file in uploaded_files:
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
        
        if df_list:
            df = pd.concat(df_list, ignore_index=True)
            # Cleanup headers if they were imported as data
            if any(df.iloc[0].isna()) == False:
                df.columns = df.iloc[0]
                df = df[1:].reset_index(drop=True)

        if not df.empty:
            st.write("### 🔍 Current Data Preview", df.head(10).astype(str))
            
            # --- THE 8-TAB TOOLSET ---
            tabs = st.tabs([
                "🎯 Mapper", "🧹 Cleaner", "🕵️ Detective", "⏰ Time Math", 
                "📊 Summary", "🔄 Shifter", "✅ Validator", "📂 Word-to-Excel"
            ])

            # TAB 1: LOGIC MAPPER (For the Fire Loss Inventory User)
            with tabs[0]:
                st.header("Logic & Category Mapper")
                st.write("Automatically categorize items based on keywords (e.g., if Item contains 'TV', Category is 'Electronics').")
                col_to_check = st.selectbox("Column to scan:", df.columns, key="map_col")
                keyword = st.text_input("Keyword to find:")
                category_val = st.text_input("Category/Value to assign:")
                if st.button("Apply Categorization"):
                    if 'Category' not in df.columns: df['Category'] = "Uncategorized"
                    df.loc[df[col_to_check].astype(str).str.contains(keyword, case=False, na=False), 'Category'] = category_val
                    st.success(f"Tagged items containing '{keyword}' as '{category_val}'!")

            # TAB 2: TEXT CLEANER (For the Scientific Notation User)
            with tabs[1]:
                st.header("The Formatting Fixer")
                clean_col = st.selectbox("Select Column:", df.columns, key="clean_tab")
                c_opt = st.radio("Fix Type:", ["Remove Symbols ($, -, %)", "Force Plain Text (Fix Scientific Notation)", "Trim Extra Spaces"])
                if st.button("Run Cleaner"):
                    if "Symbols" in c_opt:
                        df[clean_col] = df[clean_col].astype(str).str.replace(r'[$\-,%]', '', regex=True)
                    elif "Plain Text" in c_opt:
                        df[clean_col] = df[clean_col].astype(str).apply(lambda x: '{:.10f}'.format(float(x)).rstrip('0').rstrip('.') if 'E' in str(x) else x)
                    elif "Spaces" in c_opt:
                        df[clean_col] = df[clean_col].astype(str).str.strip()
                    st.success("Column cleaned!")

            # TAB 3: DUPLICATE DETECTIVE (For the Insurance Auditor)
            with tabs[2]:
                st.header("Duplicate Detective")
                match_cols = st.multiselect("Match rows based on these columns:", df.columns)
                if st.button("Scan for Duplicates"):
                    df['Is_Duplicate'] = df.duplicated(subset=match_cols, keep=False)
                    dupes = df[df['Is_Duplicate']]
                    st.warning(f"Found {len(dupes)} duplicate rows.")
                    st.dataframe(dupes.astype(str))

            # TAB 4: TIME CALCULATOR (For the Payroll/Hours User)
            with tabs[3]:
                st.header("Time & Duration Math")
                t1 = st.selectbox("Start Time Column:", df.columns)
                t2 = st.selectbox("End Time Column:", df.columns)
                if st.button("Calculate Total Minutes"):
                    df['Total_Minutes'] = (pd.to_datetime(df[t2], errors='coerce') - pd.to_datetime(df[t1], errors='coerce')).dt.total_seconds() / 60
                    st.success("Calculation complete! See 'Total_Minutes' column in download.")

            # TAB 5: MULTI-FILE SUMMARY
            with tabs[4]:
                st.header("Data Health Report")
                st.write(f"**Total Records:** {len(df)}")
                st.write("**Missing Values Per Column:**")
                st.write(df.isnull().sum())

            # TAB 6: FORMAT SHIFTER
            with tabs[5]:
                st.header("The Shifter")
                shift_opt = st.selectbox("Export to:", ["Word Table (CSV)", "HTML Report (Save as PDF)", "JSON Data"])
                if st.button("Prepare Export"):
                    if "Word" in shift_opt:
                        st.download_button("📥 Download for Word", df.to_csv(index=False), "for_word.csv")
                    elif "HTML" in shift_opt:
                        st.download_button("📥 Download HTML", df.to_html(), "report.html")

            # TAB 7: DATA VALIDATOR (New Feature!)
            with tabs[6]:
                st.header("Data Validator")
                st.write("Check for common errors like broken emails or negative prices.")
                v_col = st.selectbox("Column to Validate:", df.columns)
                v_type = st.selectbox("Rule:", ["Should be a Number", "Should be an Email", "Cannot be Empty"])
                if st.button("Check Rules"):
                    if "Number" in v_type:
                        errors = df[~df[v_col].astype(str).str.replace('.','',1).isdigit()]
                        st.error(f"Found {len(errors)} non-number entries!")
                    st.dataframe(errors)

            # TAB 8: WORD-TO-EXCEL (Human Guidance)
            with tabs[7]:
                st.header("Word Table Extractor")
                st.info("Your uploaded Word files were automatically scanned for tables. They are already merged into the preview above!")
                st.write("Tip: If the Word table didn't import correctly, save the Word doc as a PDF and re-upload it.")

            # --- THE FINAL DOWNLOAD ---
            st.divider()
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Ready to go?")
                st.write("All changes from the tabs above are included in this file.")
            with col2:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False)
                st.download_button("📥 DOWNLOAD FIXED EXCEL FILE", output.getvalue(), "fixed_by_swiss_army_knife.xlsx", use_container_width=True)
