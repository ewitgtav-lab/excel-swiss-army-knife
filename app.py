import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO

st.set_page_config(page_title="Ultimate Excel Automator", layout="wide")

st.title("🛠️ The Excel Swiss Army Knife")
st.markdown("Automating the most common spreadsheet headaches from Reddit.")

# --- FILE UPLOAD ---
# Accept PDFs now as well
uploaded_files = st.file_uploader("Upload Excel, CSV, or PDF file(s)", 
                                  type=["xlsx", "csv", "pdf"], 
                                  accept_multiple_files=True)

df = pd.DataFrame() # Initialize empty dataframe

if uploaded_files:
    # Logic to handle file types
    if len(uploaded_files) == 1:
        file = uploaded_files[0]
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        elif file.name.endswith('.pdf'):
            # Specific PDF extraction logic for the "Main" dataframe preview
            with pdfplumber.open(file) as pdf:
                all_rows = []
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        all_rows.extend(table)
                if all_rows:
                    df = pd.DataFrame(all_rows[1:], columns=all_rows[0])
        else:
            df = pd.read_excel(file)
    else:
        # Multi-file merger
        df_list = []
        for f in uploaded_files:
            if f.name.endswith('.csv'): df_list.append(pd.read_csv(f))
            elif f.name.endswith(('.xlsx', '.xls')): df_list.append(pd.read_excel(f))
        if df_list:
            df = pd.concat(df_list, ignore_index=True)

    if not df.empty:
        st.write("### Data Preview", df.head(5))
        st.divider()

        # --- TABBED INTERFACE ---
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "🎯 Logic Mapper", 
            "📄 PDF Extractor",
            "🧹 Text Cleaner", 
            "⏰ Time Calculator", 
            "📊 Data Merger"
        ])

        # TAB 1: LOGIC MAPPER (Dependent Cells)
        with tab1:
            st.header("Conditional Data Population")
            col_a, col_b = st.columns(2)
            with col_a:
                source = st.selectbox("Trigger Column", df.columns)
            with col_b:
                target = st.text_input("New Column Name", "Result_Column")

            mapping = {val: st.text_input(f"If '{val}':", key=f"map_{val}") 
                       for val in df[source].unique()}
            
            if st.button("Apply Logic"):
                df[target] = df[source].map(mapping)
                st.success("Logic Applied!")

        # TAB 2: PDF EXTRACTOR (The New Addition)
        with tab5: # Keeping the merger info here
            st.header("Multi-File Summary")
            st.write(f"Total Rows: {len(df)}")

        with tab2:
            st.header("PDF to Table Extraction")
            st.info("This tab shows the raw data extracted from your PDF.")
            if any(f.name.endswith('.pdf') for f in uploaded_files):
                st.dataframe(df)
            else:
                st.warning("Please upload a PDF file to use this feature.")

        # TAB 3: TEXT CLEANER
        with tab3:
            st.header("Quick String Scrubbing")
            clean_col = st.selectbox("Select Column to Clean", df.columns)
            char_to_rem = st.text_input("Character(s) to remove (e.g., - )", "-")
            if st.button("Clean Column"):
                df[clean_col] = df[clean_col].astype(str).str.replace(char_to_rem, "", regex=False)
                st.success("Cleaned!")

        # TAB 4: TIME CALCULATOR
        with tab4:
            st.header("Time & Date Math")
            t_col1 = st.selectbox("Start Time", df.columns)
            t_col2 = st.selectbox("End Time", df.columns)
            if st.button("Calculate Minutes"):
                try:
                    df['Duration_Mins'] = (pd.to_datetime(df[t_col2]) - pd.to_datetime(df[t_col1])).dt.total_seconds() / 60
                    st.dataframe(df.head())
                except:
                    st.error("Check date formats!")

        # --- DOWNLOAD ---
        st.divider()
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 Download Processed File", output.getvalue(), "results.xlsx")
