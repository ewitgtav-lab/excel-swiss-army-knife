import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO

st.set_page_config(page_title="Ultimate Excel Automator", layout="wide")

# --- SIDEBAR (Usage Instructions & Feedback) ---
with st.sidebar:
    st.header("How to use:")
    st.markdown("""
    1. **Upload** your messy Excel, CSV, or PDF.
    2. **Choose a tab** based on your problem (e.g., Logic Mapper).
    3. **Configure** your rules or cleaning settings.
    4. **Download** the fixed file at the bottom!
    """)
    
    st.divider()
    
    st.header("Feedback Loop")
    st.info("Found a bug or have a request?")
    # Replacement for "Report a Bug" link
    st.link_button("Report a Bug", "https://forms.gle/ohLGcn4jUv9B4ppH8") 
    
    st.divider()
    st.write("Built for the r/excel community.")

st.title("🛠️ The Excel Swiss Army Knife")
st.markdown("Automating the most common spreadsheet headaches from Reddit.")

# --- FILE UPLOAD ---
uploaded_files = st.file_uploader("Upload Excel, CSV, or PDF file(s)", 
                                  type=["xlsx", "csv", "pdf"], 
                                  accept_multiple_files=True)

df = pd.DataFrame() 

if uploaded_files:
    if len(uploaded_files) == 1:
        file = uploaded_files[0]
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        elif file.name.endswith('.pdf'):
            try:
                # Using pdfplumber for better table extraction than PyPDF2
                with pdfplumber.open(file) as pdf:
                    all_rows = []
                    for page in pdf.pages:
                        table = page.extract_table()
                        if table:
                            valid_rows = [row for row in table if any(c is not None and str(c).strip() != "" for c in row)]
                            all_rows.extend(valid_rows)
                    
                    if all_rows:
                        # Pad rows to ensure consistent column counts for multi-page PDFs
                        max_cols = max(len(r) for r in all_rows)
                        padded_rows = [r + [None] * (max_cols - len(r)) for r in all_rows]
                        
                        df = pd.DataFrame(padded_rows)
                        
                        # Sanitize headers to ensure they are unique strings
                        raw_headers = df.iloc[0]
                        clean_headers = []
                        for i, val in enumerate(raw_headers):
                            h_name = str(val).strip() if val else f"Column_{i}"
                            if h_name in clean_headers:
                                h_name = f"{h_name}_{i}"
                            clean_headers.append(h_name)
                            
                        df.columns = clean_headers
                        df = df[1:].reset_index(drop=True)
                    else:
                        st.warning("No tables found in this PDF.")
                        df = pd.DataFrame()
            except Exception as e:
                st.error(f"PDF Error: {e}")
                df = pd.DataFrame()
        else:
            df = pd.read_excel(file)
    else:
        # Multi-file merger
        df_list = []
        for f in uploaded_files:
            if f.name.endswith('.csv'): 
                df_list.append(pd.read_csv(f))
            elif f.name.endswith(('.xlsx', '.xls')): 
                df_list.append(pd.read_excel(f))
        if df_list:
            df = pd.concat(df_list, ignore_index=True)

    if not df.empty:
        st.write("### Data Preview", df.head(5))
        st.divider()

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "🎯 Logic Mapper", 
            "📄 PDF Extractor",
            "🧹 Text Cleaner", 
            "⏰ Time Calculator", 
            "📊 Data Merger"
        ])

        # TAB 1: LOGIC MAPPER (Solves "populate based on entry")
        with tab1:
            st.header("Conditional Data Population")
            col_a, col_b = st.columns(2)
            with col_a:
                source = st.selectbox("Trigger Column", df.columns, key="logic_src")
            with col_b:
                target = st.text_input("New Column Name", "Result_Column")

            mapping = {val: st.text_input(f"If '{val}':", key=f"map_{val}") for val in df[source].unique()}
            
            if st.button("Apply Logic"):
                df[target] = df[source].map(mapping)
                st.success("Logic Applied!")
                st.dataframe(df.head())

        # TAB 2: PDF EXTRACTOR (Solves "organized PDF extraction")
        with tab2:
            st.header("PDF to Table Extraction")
            st.info("Raw data extracted from your PDF document.")
            st.dataframe(df)

        # TAB 3: TEXT CLEANER (Solves "remove dash from numbers")
        with tab3:
            st.header("Quick String Scrubbing")
            clean_col = st.selectbox("Select Column to Clean", df.columns, key="clean_src")
            char_to_rem = st.text_input("Character(s) to remove", "-")
            if st.button("Clean Column"):
                df[clean_col] = df[clean_col].astype(str).str.replace(char_to_rem, "", regex=False)
                st.success(f"Removed '{char_to_rem}' from {clean_col}!")
                st.dataframe(df.head())

        # TAB 4: TIME CALCULATOR (Solves "difference in time by minutes")
        with tab4:
            st.header("Time & Date Math")
            t_col1 = st.selectbox("Start Time", df.columns, key="time_src1")
            t_col2 = st.selectbox("End Time", df.columns, key="time_src2")
            if st.button("Calculate Minutes"):
                try:
                    df['Duration_Mins'] = (pd.to_datetime(df[t_col2]) - pd.to_datetime(df[t_col1])).dt.total_seconds() / 60
                    st.success("Calculated durations!")
                    st.dataframe(df.head())
                except:
                    st.error("Check date formats! Ensure columns contain valid time data.")

        # TAB 5: DATA MERGER
        with tab5:
            st.header("Multi-File Summary")
            st.write(f"Total Rows: {len(df)}")
            st.write(f"Total Columns: {len(df.columns)}")

        # --- DOWNLOAD ---
        st.divider()
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 Download Processed File", output.getvalue(), "automated_results.xlsx")

# PRO-TIP: To seed your community, check the Reddit usernames from your screenshots:
# - u/mrklmngbta (Removing dashes)
# - u/Escaping_Peter_Pan (MT 940 Statement ETL)
# - u/ThenLandscape2108 (Organized PDF extraction)
