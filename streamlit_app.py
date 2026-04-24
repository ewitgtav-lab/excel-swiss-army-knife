import re
from io import BytesIO

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Ultimate Excel Automator", layout="wide")


def _freeze_uploads(files) -> list[tuple[str, bytes]]:
    return [(f.name, f.getvalue()) for f in files]


@st.cache_data(show_spinner="Sharpening the Knife...", ttl=3600)
def load_and_sanitize(frozen_uploads: list[tuple[str, bytes]]) -> pd.DataFrame:
    # Lazy imports: speeds up app startup.
    import pdfplumber
    from docx import Document

    df_list: list[pd.DataFrame] = []
    for name, data in frozen_uploads:
        lower = name.lower()
        bio = BytesIO(data)
        try:
            if lower.endswith(".csv"):
                df_list.append(pd.read_csv(bio))
            elif lower.endswith(".docx"):
                doc = Document(bio)
                df_list.append(
                    pd.DataFrame(
                        [[c.text for c in r.cells] for t in doc.tables for r in t.rows]
                    )
                )
            elif lower.endswith(".pdf"):
                with pdfplumber.open(bio) as pdf:
                    for page in pdf.pages:
                        table = page.extract_table()
                        if table:
                            df_list.append(pd.DataFrame(table))
            else:
                df_list.append(pd.read_excel(bio))
        except Exception as e:
            # Surface errors without breaking the whole batch.
            st.error(f"Error reading {name}: {e}")

    if not df_list:
        return pd.DataFrame()

    df = pd.concat(df_list, ignore_index=True)
    df.columns = df.columns.map(lambda c: str(c).strip())
    return df


def _bump_df_version():
    st.session_state.df_version = int(st.session_state.get("df_version", 0)) + 1
    st.session_state.pop("xlsx_bytes", None)


def _ensure_state():
    st.session_state.setdefault("main_df", None)
    st.session_state.setdefault("df_version", 0)


def _build_xlsx_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


_ensure_state()

with st.sidebar:
    st.header("🛠️ Tool Belt")
    if st.button("♻️ Reset App & Memory"):
        st.session_state.main_df = None
        st.session_state.pop("xlsx_bytes", None)
        st.session_state.df_version = 0
        st.cache_data.clear()
        st.rerun()

    st.divider()
    st.header("📣 Support")
    st.link_button("🪲 Report a Bug", "https://forms.gle/rVE2KkorZX4iqWNq7")
    st.link_button("☕ Support my Work", "https://paypal.me/GewishCatedrilla")


st.title("🛠️ The Excel Swiss Army Knife")

if st.session_state.main_df is None:
    files = st.file_uploader(
        "Upload Files",
        type=["xlsx", "csv", "pdf", "docx"],
        accept_multiple_files=True,
    )
    if files:
        frozen = _freeze_uploads(files)
        st.session_state.main_df = load_and_sanitize(frozen)
        _bump_df_version()
        st.rerun()


if st.session_state.main_df is None:
    st.info("Upload one or more files to begin.")
    st.stop()


df: pd.DataFrame = st.session_state.main_df
st.write(f"### 🔍 Preview ({len(df)} rows)")
st.dataframe(df.head(5), width="stretch")

tabs = st.tabs(
    [
        "🧮 Aggregator",
        "🎯 Mapper",
        "🧹 Cleaner",
        "🕵️ Detective",
        "⏰ Time Math",
        "🔄 Shifter",
        "✅ Validator",
        "📂 Word",
    ]
)


with tabs[0]:
    st.header("Smart Aggregator")
    with st.form("agg_form"):
        g_col = st.selectbox("Group by:", df.columns, key="agg_g")
        s_col = st.selectbox("Sum Column:", df.columns, key="agg_s")
        go = st.form_submit_button("Generate Summary")
    if go:
        s = pd.to_numeric(
            df[s_col].astype(str).str.replace(r"[$,]", "", regex=True),
            errors="coerce",
        )
        value_name = s_col if g_col != s_col else f"{s_col}_sum"
        summary = (
            df.assign(_s=s)
            .groupby(g_col, sort=False)["_s"]
            .sum()
            .reset_index(name=value_name)
        )
        st.dataframe(summary, width="stretch")


with tabs[1]:
    st.header("Categorization Logic")
    with st.form("map_form"):
        m_col = st.selectbox("Scan Column:", df.columns, key="map_c")
        keyword = st.text_input("If text contains:")
        target = st.text_input("Then assign this category:")
        go = st.form_submit_button("Apply Mapping")
    if go:
        if not keyword:
            st.warning("Enter a keyword to match.")
        else:
            if "Category" not in df.columns:
                df["Category"] = "Uncategorized"
            mask = df[m_col].astype(str).str.contains(re.escape(keyword), case=False, na=False)
            df.loc[mask, "Category"] = target if target else "Uncategorized"
            st.session_state.main_df = df
            _bump_df_version()
            st.success("Logic Applied! Check preview above.")


with tabs[2]:
    st.header("Formatting Fixer")
    with st.form("clean_form"):
        c_col = st.selectbox("Target Column:", df.columns, key="clean_c")
        c_opt = st.radio(
            "Fix Type:",
            ["Scientific Notation", "Remove Symbols", "Proper Case"],
            horizontal=True,
        )
        go = st.form_submit_button("Run Fixer")
    if go:
        if c_opt == "Remove Symbols":
            df[c_col] = df[c_col].astype(str).str.replace(r"[$\-,%]", "", regex=True)
        elif c_opt == "Proper Case":
            df[c_col] = df[c_col].astype(str).str.title()
        else:
            s = pd.to_numeric(df[c_col], errors="coerce")
            mask = s.notna()
            out = df[c_col].astype(object)
            # Fast-enough formatting for typical Excel columns; only formats numeric values.
            out.loc[mask] = s.loc[mask].map(
                lambda x: f"{x:.0f}" if float(x).is_integer() else f"{x:.10f}".rstrip("0").rstrip(".")
            )
            df[c_col] = out
        st.session_state.main_df = df
        _bump_df_version()
        st.success("Cleaned!")


with tabs[3]:
    st.header("Duplicate Detective")
    with st.form("dupe_form"):
        d_cols = st.multiselect("Match rows on these columns:", df.columns, default=[])
        go = st.form_submit_button("Identify Duplicates")
    if go:
        if not d_cols:
            st.warning("Pick at least one column to match on.")
        else:
            dupes = df[df.duplicated(subset=d_cols, keep=False)]
            if not dupes.empty:
                st.warning(f"Found {len(dupes)} duplicates.")
                st.dataframe(dupes.astype(str), width="stretch")
            else:
                st.success("No duplicates found!")


with tabs[4]:
    st.header("Time & Labor Math")
    with st.form("time_form"):
        h_col = st.selectbox("Select Hours Column:", df.columns, key="tm_h")
        m_col = st.selectbox("Select Minutes Column:", df.columns, key="tm_m")
        go = st.form_submit_button("Combine to Decimal Hours")
    if go:
        h = pd.to_numeric(df[h_col], errors="coerce").fillna(0)
        m = pd.to_numeric(df[m_col], errors="coerce").fillna(0)
        df["Total_Hours_Decimal"] = h + (m / 60)
        st.session_state.main_df = df
        _bump_df_version()
        st.success("Created 'Total_Hours_Decimal' column!")


with tabs[5]:
    st.header("Format Shifter")
    with st.form("shift_form"):
        s_opt = st.selectbox("Convert to:", ["Word-Ready CSV", "HTML Report", "JSON"])
        go = st.form_submit_button("Prepare Conversion")
    if go:
        if s_opt == "Word-Ready CSV":
            st.download_button(
                "📥 Download CSV",
                df.to_csv(index=False).encode("utf-8"),
                "for_word.csv",
                width="stretch",
            )
        elif s_opt == "HTML Report":
            st.download_button(
                "📥 Download HTML",
                df.to_html(index=False).encode("utf-8"),
                "report.html",
                width="stretch",
            )
        else:
            st.download_button(
                "📥 Download JSON",
                df.to_json(orient="records").encode("utf-8"),
                "export.json",
                width="stretch",
            )


with tabs[6]:
    st.header("Rule Validator")
    with st.form("val_form"):
        v_col = st.selectbox("Column to Validate:", df.columns, key="val_c")
        v_type = st.radio("Rule:", ["Must be Numeric", "Must be Email", "Cannot be Empty"], horizontal=True)
        go = st.form_submit_button("Validate Now")
    if go:
        if v_type == "Must be Numeric":
            clean = pd.to_numeric(
                df[v_col].astype(str).str.replace(r"[$\-,]", "", regex=True),
                errors="coerce",
            )
            errors = df[clean.isna() & df[v_col].astype(str).str.strip().ne("")]
        elif v_type == "Must be Email":
            # Simple, fast heuristic (not RFC-perfect).
            email_ok = df[v_col].astype(str).str.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", na=False)
            errors = df[~email_ok]
        else:
            errors = df[df[v_col].astype(str).str.strip() == ""]

        if not errors.empty:
            st.error(f"🚨 Found {len(errors)} invalid rows!")
            st.dataframe(errors.astype(str), width="stretch")
        else:
            st.success("✅ All data in this column is valid!")


with tabs[7]:
    st.header("Word Table Import")
    st.info("Your uploaded .docx tables are already combined in the live preview.")


st.divider()

col1, col2 = st.columns([1, 2], vertical_alignment="center")
with col1:
    if st.button("⚡ Prepare Excel Download", width="stretch"):
        st.session_state.xslx_bytes_error = None
        try:
            st.session_state.xlsx_bytes = _build_xlsx_bytes(df)
        except Exception as e:
            st.session_state.xslx_bytes_error = str(e)

with col2:
    err = st.session_state.get("xslx_bytes_error")
    if err:
        st.error(f"Excel export failed: {err}")
    xbytes = st.session_state.get("xlsx_bytes")
    st.download_button(
        "📥 DOWNLOAD COMPLETED SWISS ARMY FILE",
        xbytes if xbytes else b"",
        "fixed_data.xlsx",
        disabled=not bool(xbytes),
        width="stretch",
    )

