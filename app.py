import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO

# --- CONFIGURATION ---
st.set_page_config(page_title="Ultimate Excel Automator", layout="wide")

def check_password():
    """Returns True if the user had the correct password."""
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
        st.header("How to use:")
        st.markdown("""
        1. **Upload** your messy Excel, CSV, or PDF.
        2. **Choose a tab** based on your problem.
        3. **Configure** your rules.
        4. **Download** the fixed file!
        """)
        st.divider()
        st.header("Feedback Loop")
        st.link_button("Report a Bug", "https://forms.gle/your_google_form_link") 
        st.divider()
        st.write("Built for the r/excel community.")

    st.title("🛠️ The Excel Swiss Army Knife")
    st.markdown("Automating the most common spreadsheet headaches from Reddit.")

    uploaded_files = st.file_uploader("Upload Excel, CSV, or PDF file(s)", type=["xlsx", "csv", "pdf"], accept_multiple_files=True)

    df = pd.DataFrame()

    if uploaded_files:
        if len(uploaded_files) == 1:
            file = uploaded_files[0]
            if file.name.endswith('.csv'):
                df = pd.read_csv(file)
            elif file.name.endswith('.pdf'):
                try:
                    with pdfplumber.open(file) as pdf:
                        all_rows = []
                        for page in pdf.pages:
                            table = page.extract_table()
                            if table:
                                valid_rows = [row for row in table if any(c is not None and str(c).strip() != "" for c in row)]
                                all_rows.extend(valid_rows)
                        if all_rows:
                            max_cols = max(len(r) for r in all_rows)
                            padded_rows = [r + [None] * (max_cols - len(r)) for r in all_rows]
                            df = pd.DataFrame(padded_rows)
                            raw_headers = df.iloc[0]
                            clean_headers = []
                            for i, val in enumerate(raw_headers):
                                h_name = str(val).strip() if val else f"Column_{i}"
                                if h_name in clean_headers: h_name = f"{h_name}_{i}"
                                clean_headers.append(h_name)
                            df.columns = clean_headers
                            df = df[1:].reset_index(drop=True)
                except Exception as e:
                    st.error(f"PDF Error: {e}")
            else:
                df = pd.read_excel(file)
        else:
            df_list = [pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f) for f in uploaded_files]
            df = pd.concat(df_list, ignore_index=True)

        if not df.empty:
            # FIX: Convert preview to string to prevent Arrow Serialization errors
            st.write("### Data Preview", df.head(5).astype(str))
            st.divider()

            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "🎯 Logic Mapper", "📄 PDF Extractor", "🧹 Text Cleaner", 
                "⏰ Time Calculator", "📊 Data Merger", "🕵️ Duplicate Detective"
            ])

            with tab1:
                st.header("Conditional Data Population")
                source = st.selectbox("Trigger Column", df.columns, key="logic_src")
                target = st.text_input("New Column Name", "Result_Column")
                mapping = {val: st.text_input(f"If '{val}':", key=f"map_{val}") for val in df[source].unique()}
                if st.button("Apply Logic"):
                    df[target] = df[source].map(mapping)
                    st.success("Logic Applied!")

            with tab2:
                st.header("PDF to Table Extraction")
                st.dataframe(df.astype(str))

            with tab3:
                st.header("Quick String Scrubbing")
                clean_col = st.selectbox("Select Column to Clean", df.columns, key="clean_src")
                char_to_rem = st.text_input("Character(s) to remove", "-")
                if st.button("Clean Column"):
                    df[clean_col] = df[clean_col].astype(str).str.replace(char_to_rem, "", regex=False)
                    st.success(f"Cleaned {clean_col}!")

            # --- TAB 4 FIX (Time Calculator) ---
with tab4:
    st.header("Time & Date Math")
    t_col1 = st.selectbox("Start Time", df.columns, key="time_src1")
    t_col2 = st.selectbox("End Time", df.columns, key="time_src2")
    if st.button("Calculate Minutes"):
        # errors='coerce' prevents the "SHP-101" crash by turning bad dates into "NaT" instead of an error
        df['Duration_Mins'] = (pd.to_datetime(df[t_col2], errors='coerce') - 
                               pd.to_datetime(df[t_col1], errors='coerce')).dt.total_seconds() / 60
        st.success("Calculated durations! Invalid dates were skipped.")

# --- TAB 5 UPDATE (Added Column Labels) ---
with tab5:
    st.header("📊 Multi-File Summary")
    st.write(f"**Total Rows:** {len(df)}")
    st.write(f"**Total Columns:** {len(df.columns)}")
    st.write("**Column Labels:**")
    st.info(", ".join(df.columns)) # This lists all your headers clearly

# --- TAB 6 FIX (Robust Duplicate Detective) ---
with tab6:
    st.header("🕵️ Duplicate Detective")
    match_cols = st.multiselect("Select columns that should be unique:", df.columns)
    
    if match_cols:
        if st.button("Run Duplicate Scan"):
            # 1. Create a clean temporary version for scanning
            # We remove rows that are entirely empty in the match columns
            scan_df = df.dropna(subset=match_cols).copy()
            
            # 2. Force to string to prevent Arrow crashes with mixed IDs (e.g., TT-991)
            for col in match_cols:
                scan_df[col] = scan_df[col].astype(str)
            
            # 3. Mark duplicates
            df['Is_Duplicate'] = False
            df.loc[scan_df.index, 'Is_Duplicate'] = scan_df.duplicated(subset=match_cols, keep=False)
            
            dupes_only = df[df['Is_Duplicate'] == True].sort_values(by=match_cols)
            
            if not dupes_only.empty:
                st.warning(f"🚨 Found {len(dupes_only)} duplicate rows!")
                def highlight_dupes(x):
                    return ['background-color: #4b2525' if x.Is_Duplicate else '' for _ in x]
                # astype(str) here ensures the UI doesn't crash on "SHP-103" types
                st.dataframe(dupes_only.astype(str).style.apply(highlight_dupes, axis=1))
            else:
                st.success("✅ No duplicates found in your selected columns!")
            st.divider()
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 Download Processed File", output.getvalue(), "automated_results.xlsx")
