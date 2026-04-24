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
        # --- LOADING LOGIC ---
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
            st.write("### Data Preview", df.head(5).astype(str))
            st.divider()

            # --- TABS (Fixed Indentation) ---
            tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
                "🎯 Logic Mapper", "📄 PDF Extractor", "🧹 Text Cleaner", 
                "⏰ Time Calculator", "📊 Data Merger", "🕵️ Duplicate Detective", "🔄 Format Shifter"
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
                st.write("Data extracted from PDF is shown below. You can now use other tabs to clean it.")
                st.dataframe(df.astype(str))

            with tab3:
                st.header("Quick String Scrubbing")
                clean_col = st.selectbox("Select Column to Clean", df.columns, key="clean_src")
                char_to_rem = st.text_input("Character(s) to remove", "-")
                if st.button("Clean Column"):
                    df[clean_col] = df[clean_col].astype(str).str.replace(char_to_rem, "", regex=False)
                    st.success(f"Cleaned {clean_col}!")

            with tab4:
                st.header("Time & Date Math")
                t_col1 = st.selectbox("Start Time", df.columns, key="time_src1")
                t_col2 = st.selectbox("End Time", df.columns, key="time_src2")
                if st.button("Calculate Minutes"):
                    df['Duration_Mins'] = (pd.to_datetime(df[t_col2], errors='coerce') - 
                                           pd.to_datetime(df[t_col1], errors='coerce')).dt.total_seconds() / 60
                    st.success("Calculated durations!")

            with tab5:
                st.header("📊 Multi-File Summary")
                st.write(f"**Total Rows:** {len(df)}")
                st.write(f"**Total Columns:** {len(df.columns)}")
                st.info(", ".join(df.columns)) 

            with tab6:
                st.header("🕵️ Duplicate Detective")
                match_cols = st.multiselect("Select unique-identifying columns:", df.columns)
                if match_cols:
                    if st.button("Run Duplicate Scan"):
                        scan_df = df.dropna(subset=match_cols).copy()
                        for col in match_cols:
                            scan_df[col] = scan_df[col].astype(str)
                        df['Is_Duplicate'] = False
                        df.loc[scan_df.index, 'Is_Duplicate'] = scan_df.duplicated(subset=match_cols, keep=False)
                        dupes_only = df[df['Is_Duplicate'] == True].sort_values(by=match_cols)
                        if not dupes_only.empty:
                            st.warning(f"🚨 Found {len(dupes_only)} duplicate rows!")
                            def highlight_dupes(x):
                                return ['background-color: #4b2525' if x.Is_Duplicate else '' for _ in x]
                            st.dataframe(dupes_only.astype(str).style.apply(highlight_dupes, axis=1))
                        else:
                            st.success("✅ No duplicates found!")

            with tab7:
                st.header("🔄 The Format Shifter")
                st.write("Convert your data into report-ready formats.")
                convert_option = st.selectbox("Export format:", [
                    "Excel to PDF (Clean Report)",
                    "Excel to Word (Ready for Tables)",
                    "Data to JSON (For Developers)"
                ])

                if st.button("Process Conversion"):
                    try:
                        if "PDF" in convert_option:
                            html = df.to_html()
                            st.download_button("📥 Download HTML Report", data=html, file_name="report.html")
                            st.info("Browser Tip: Open the HTML file and press Ctrl+P to 'Save as PDF'.")
                        elif "Word" in convert_option:
                            csv_buffer = BytesIO()
                            df.to_csv(csv_buffer, index=False)
                            st.download_button("📥 Download Word-Ready CSV", data=csv_buffer.getvalue(), file_name="word_table.csv")
                            st.info("Word Tip: Insert > Table > Convert Text to Table in Microsoft Word.")
                        elif "JSON" in convert_option:
                            json_data = df.to_json(orient="records", indent=4)
                            st.download_button("📥 Download JSON", data=json_data, file_name="export.json")
                        st.success("Conversion Ready!")
                    except Exception as e:
                        st.error(f"Error: {e}")

            # --- FINAL DOWNLOAD (Must stay inside 'if not df.empty') ---
            st.divider()
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 Download Main Processed Excel", output.getvalue(), "automated_results.xlsx")
