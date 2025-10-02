
import re
import io
import math
import json
import pandas as pd
import streamlit as st

# Optional PDF parsing; app still runs without pdfplumber installed
try:
    import pdfplumber  # type: ignore
    HAS_PDFPLUMBER = True
except Exception:
    HAS_PDFPLUMBER = False

st.set_page_config(page_title="Reserves Tie-Out Checker", layout="wide")
st.title("📊 Reserves Tie-Out Checker")
st.caption("Upload your standard-format outputs (PDF and/or XLSX). This version includes a parser for the Schaper Energy PDF layout (Table 1.1, Cash Flows, and One‑line Summary).")

# ------------------------------------------------------------------------------------------
# Sidebar options
# ------------------------------------------------------------------------------------------
with st.sidebar:
    st.header("Options")
    abs_tol = st.number_input("Absolute tolerance", min_value=0.0, value=0.5, step=0.1, help="Allowed absolute difference in units (Mboe or $MM).")
    rel_tol_pct = st.number_input("Relative tolerance (%)", min_value=0.0, value=0.1, step=0.05, help="Allowed percent difference.")
    strict = st.checkbox("Strict: every source must have every field", value=False)
    case_name = st.text_input("Case/Project name", "")
    st.markdown("---")
    st.caption("If PDFs fail to parse, ensure `pdfplumber` is installed.")

# ------------------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------------------
def _to_float(s):
    if s is None:
        return math.nan
    s = str(s)
    s = s.replace(",", "").replace("$","").strip()
    try:
        return float(s)
    except Exception:
        return math.nan

def within_tolerance(vals, abs_tol, rel_tol_pct):
    s = pd.Series(vals, dtype="float64").dropna()
    if s.empty: 
        return False
    if len(s)==1:
        return True
    mn, mx = float(s.min()), float(s.max())
    if abs(mx-mn) <= abs_tol:
        return True
    denom = max(abs(mx), abs(mn), 1e-12)
    return (abs(mx-mn)/denom*100.0) <= rel_tol_pct

# ------------------------------------------------------------------------------------------
# Schaper PDF specific parsing
# ------------------------------------------------------------------------------------------
CAT_KEYS = ["1PDP", "4PUD", "5PROB", "6POSS", "TOTAL PROVED"]

TABLE11_ROW_PAT = re.compile(
    r"(?i)(Total\s+Proved\s+Reserves|Proved\s+Developed\s+Producing\s+\(1PDP\)|Proved\s+Undeveloped\s+\(4PUD\)|Total\s+Probable\s+Reserves\s+\(5PROB\)|Total\s+Possible\s+Reserves\s+\(6POSS\)).*?"
    r"([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+\$\s*([0-9,]+)\s+\$\s*([0-9,]+)"
)

CASHFLOWS_TOTAL_PAT = re.compile(
    r"(?is)SE_RSV_CAT\s*=\s*(1PDP|4PUD|5PROB|6POSS).*?TOTAL\s+[0-9\.\s]+\s+([0-9\.,]+)\s+([0-9\.,]+)\s*$"
)

ONELINE_TOTAL_PAT = re.compile(
    r"(?im)^TOTAL\s+(1PDP|4PUD|5PROB)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9,]+)$"
)

def parse_pdf_schaper(file_obj):
    if not HAS_PDFPLUMBER:
        return None, "pdfplumber is not installed"
    # Extract full text
    text_pages = []
    try:
        with pdfplumber.open(file_obj) as pdf:
            for p in pdf.pages:
                t = p.extract_text() or ""
                text_pages.append(t)
    except Exception as e:
        return None, f"Failed to read PDF: {e}"
    blob = "\n".join(text_pages)

    # 1) Table 1.1 rows
    table11 = []
    for m in TABLE11_ROW_PAT.finditer(blob):
        label = m.group(1)
        gas, ngl, oil, boe, undisc, pv10 = [_to_float(m.group(i)) for i in range(2,8)]
        key = None
        if "Developed" in label: key = "1PDP"
        elif "Undeveloped" in label: key = "4PUD"
        elif "Probable" in label: key = "5PROB"
        elif "Possible" in label: key = "6POSS"
        elif "Total Proved Reserves" in label: key = "TOTAL PROVED"
        if key:
            table11.append({"Source":"Table1.1", "Category": key, "Net BOE (Mboe)": boe, "PV10 ($MM)": pv10})
    # 2) Cash Flows - use final TOTAL line: TOTAL <...> <Cum Free Cashflow> <CUM. DISC. FCF> 
    cash_totals = []
    for m in CASHFLOWS_TOTAL_PAT.finditer(blob):
        key = m.group(1)
        undisc = _to_float(m.group(2))
        pv10 = _to_float(m.group(3))
        cash_totals.append({"Source":"Cash Flows", "Category": key, "Net BOE (Mboe)": math.nan, "PV10 ($MM)": pv10})

    # 3) One-line Summary - TOTAL rows (values in dollars)
    one_line = []
    for m in ONELINE_TOTAL_PAT.finditer(blob):
        key = m.group(1)
        oil = _to_float(m.group(2))
        gas = _to_float(m.group(3))
        ngl = _to_float(m.group(4))
        boe = _to_float(m.group(5))
        pv10_dollars = _to_float(m.group(6))
        one_line.append({"Source":"One-line", "Category": key, "Net BOE (Mboe)": boe, "PV10 ($MM)": pv10_dollars/1_000.0})

    rows = table11 + cash_totals + one_line
    if not rows:
        return None, "No recognizable sections found. Is this a Schaper-format PDF?"
    df = pd.DataFrame(rows)
    return df, None

def check_consistency(df):
    # For each Category x Metric, check tie across sources
    out = []
    for cat in sorted(df["Category"].unique()):
        for metric in ["Net BOE (Mboe)", "PV10 ($MM)"]:
            vals = df.loc[df["Category"]==cat, metric].dropna().tolist()
            has_vals = len(vals) >= 1
            consistent = within_tolerance(vals, abs_tol, rel_tol_pct) if has_vals else False
            out.append({"Category": cat, "Metric": metric, "Sources": int(df.loc[(df['Category']==cat) & df[metric].notna()].shape[0]), "Min": pd.Series(vals).min() if vals else math.nan, "Max": pd.Series(vals).max() if vals else math.nan, "Consistent?": bool(consistent)})
    return pd.DataFrame(out)

# ------------------------------------------------------------------------------------------
# UI
# ------------------------------------------------------------------------------------------
pdf_files = st.file_uploader("Upload Reserves PDF(s)", type=["pdf"], accept_multiple_files=True)

if pdf_files:
    all_frames = []
    with st.spinner("Parsing..."):
        for f in pdf_files:
            df, err = parse_pdf_schaper(f)
            if err:
                st.error(f"{f.name}: {err}")
                continue
            df.insert(0, "File", f.name)
            all_frames.append(df)

    if all_frames:
        merged = pd.concat(all_frames, ignore_index=True)
        st.subheader("Extracted figures")
        st.dataframe(merged)

        st.subheader("Consistency checks (by category & metric)")
        results = merged.groupby(["File"]).apply(lambda g: check_consistency(g)).reset_index(level=0).rename(columns={"level_0":"File"})
        st.dataframe(results)

        # Overall pass/fail per file
        st.subheader("Overall")
        overall_rows = []
        for file, group in results.groupby("File"):
            overall_rows.append({"File": file, "Pass?": bool(group["Consistent?"].all())})
        st.dataframe(pd.DataFrame(overall_rows))

        # Download
        st.download_button(
            "Download detailed CSV",
            data=merged.to_csv(index=False).encode("utf-8"),
            file_name=f"schaper_tieout_{(case_name or 'report').replace(' ', '_')}.csv",
            mime="text/csv"
        )
else:
    st.info("Upload at least one PDF to begin.")

st.markdown("---")
st.caption("Detector keys: Table 1.1 (Net BOE & PV10), Cash Flow (PV10), One‑line Summary (BOE & PV10). Dollar PV10s are normalized to $MM.")
