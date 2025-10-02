
import re
import io
import math
import pandas as pd
import streamlit as st

# Optional, but required for PDF parsing
try:
    import pdfplumber  # type: ignore
    HAS_PDFPLUMBER = True
except Exception:
    HAS_PDFPLUMBER = False

st.set_page_config(page_title="Reserves Tie-Out Checker", layout="wide")
st.title("📊 Reserves Tie-Out Checker")
st.caption("Schaper-format parser: Table 1.1, Cash-Flow pages (detected by header), and One-line Summary totals.")

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("Options")
    abs_tol = st.number_input("Absolute tolerance", min_value=0.0, value=0.5, step=0.1,
                              help="Allowed absolute difference (same units as metric; e.g., Mboe or $MM).")
    rel_tol_pct = st.number_input("Relative tolerance (%)", min_value=0.0, value=0.1, step=0.05,
                                  help="Allowed percent difference across sources.")
    strict = st.checkbox("Strict: every source must have every field", value=False)
    case_name = st.text_input("Case/Project name", "")

# ---------------- Helpers ----------------
def _to_f(val):
    """String -> float with commas/$ stripped."""
    if val is None:
        return math.nan
    s = str(val).replace(",", "").replace("$", "").strip()
    try:
        return float(s)
    except Exception:
        return math.nan

def within_tolerance(vals, abs_tol, rel_tol_pct):
    s = pd.Series(vals, dtype="float64").dropna()
    if s.empty:
        return False
    if len(s) == 1:
        return True
    mn, mx = float(s.min()), float(s.max())
    if abs(mx - mn) <= abs_tol:
        return True
    denom = max(abs(mx), abs(mn), 1e-12)
    return (abs(mx - mn) / denom * 100.0) <= rel_tol_pct

# ---------------- Schaper PDF parsing ----------------
# Cash-flow page: category is indicated in the header (top-left), like "SE_RSV_CAT = 1PDP"
RSV_CAT_PAT = re.compile(r"(?i)SE[_\s]*RSV[_\s]*CAT\s*[:=]\s*(1PDP|4PUD|5PROB|6POSS)")

# Table 1.1 rows: capture Net BOE and PV10 ($MM) for each category
TABLE11_ROW_PAT = re.compile(
    r"(?i)(Total\s+Proved\s+Reserves|Proved\s+Developed\s+Producing\s+\(1PDP\)|Proved\s+Undeveloped\s+\(4PUD\)|Total\s+Probable\s+Reserves\s+\(5PROB\)|Total\s+Possible\s+Reserves\s+\(6POSS\)).*?"
    r"([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+\$\s*([0-9,]+)\s+\$\s*([0-9,]+)"
)

# Cash-flow TOTAL line: capture the last two numeric tokens (Undisc, PV10) to be robust to spacing
TOTAL_LINE_PAT = re.compile(r"(?mi)^\s*TOTAL\b[^\n]*?(-?\d[\d,]*\.?\d*)\s+(-?\d[\d,]*\.?\d*)\s*$")

# One-line Summary TOTAL rows (PV10 in dollars; convert to $MM). Often for 1PDP/4PUD/5PROB.
ONELINE_TOTAL_PAT = re.compile(
    r"(?im)^\s*TOTAL\s+(1PDP|4PUD|5PROB|6POSS)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9,]+)\s*$"
)

def parse_pdf_schaper(file_obj):
    """Parse a Schaper-format reserves PDF into rows of [Source, Category, Net BOE (Mboe), PV10 ($MM)]."""
    if not HAS_PDFPLUMBER:
        return None, "pdfplumber is not installed"
    rows = []
    try:
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""

                # ---- Table 1.1 matches on any page ----
                for m in TABLE11_ROW_PAT.finditer(text):
                    label = m.group(1)
                    gas, ngl, oil, boe, undisc, pv10 = [_to_f(m.group(i)) for i in range(2, 8)]
                    key = None
                    if "Developed" in label:
                        key = "1PDP"
                    elif "Undeveloped" in label:
                        key = "4PUD"
                    elif "Probable" in label:
                        key = "5PROB"
                    elif "Possible" in label:
                        key = "6POSS"
                    elif "Total Proved Reserves" in label:
                        key = "TOTAL PROVED"
                    if key:
                        rows.append(
                            {"Source": "Table1.1", "Category": key, "Net BOE (Mboe)": boe, "PV10 ($MM)": pv10}
                        )

                # ---- Cash-flow page: detect category from header, then read last TOTAL line ----
                mcat = RSV_CAT_PAT.search(text)
                if mcat:
                    cat = mcat.group(1)
                    last_tot = None
                    for m in TOTAL_LINE_PAT.finditer(text):
                        last_tot = m  # keep the last TOTAL occurrence on the page
                    if last_tot:
                        undisc = _to_f(last_tot.group(1))
                        pv10 = _to_f(last_tot.group(2))
                        rows.append(
                            {"Source": "Cash Flows", "Category": cat, "Net BOE (Mboe)": math.nan, "PV10 ($MM)": pv10}
                        )

                # ---- One-line Summary totals ----
                for m in ONELINE_TOTAL_PAT.finditer(text):
                    key = m.group(1)
                    boe = _to_f(m.group(5))
                    pv10_dollars = _to_f(m.group(6))
                    rows.append(
                        {"Source": "One-line", "Category": key, "Net BOE (Mboe)": boe, "PV10 ($MM)": pv10_dollars / 1000.0}
                    )

    except Exception as e:
        return None, f"Failed to read PDF: {e}"

    if not rows:
        return None, "No recognizable sections found."
    return pd.DataFrame(rows), None

def check_consistency(df):
    """For each Category × Metric, check if sources tie within tolerance."""
    out = []
    for cat in sorted(df["Category"].unique()):
        for metric in ["Net BOE (Mboe)", "PV10 ($MM)"]:
            vals = df.loc[df["Category"] == cat, metric].dropna().tolist()
            if strict and len(vals) < df["Source"].nunique():
                ok = False
            else:
                ok = within_tolerance(vals, abs_tol, rel_tol_pct) if vals else False
            out.append(
                {
                    "Category": cat,
                    "Metric": metric,
                    "Sources": int(df.loc[(df["Category"] == cat) & df[metric].notna()].shape[0]),
                    "Min": pd.Series(vals).min() if vals else math.nan,
                    "Max": pd.Series(vals).max() if vals else math.nan,
                    "Consistent?": bool(ok),
                }
            )
    return pd.DataFrame(out)

# ---------------- UI ----------------
pdf_files = st.file_uploader("Upload Reserves PDF(s)", type=["pdf"], accept_multiple_files=True)

if pdf_files:
    frames = []
    with st.spinner("Parsing..."):
        for f in pdf_files:
            df, err = parse_pdf_schaper(f)
            if err:
                st.error(f"{f.name}: {err}")
            else:
                df.insert(0, "File", f.name)
                frames.append(df)

    if frames:
        merged = pd.concat(frames, ignore_index=True)

        st.subheader("Extracted figures")
        st.dataframe(merged)

        st.subheader("Consistency checks (by file)")
        results = (
            merged.groupby(["File"])
            .apply(lambda g: check_consistency(g))
            .reset_index(level=0)
            .rename(columns={"level_0": "File"})
        )
        st.dataframe(results)

        st.subheader("Overall")
        overall = results.groupby("File")["Consistent?"].all().reset_index().rename(columns={"Consistent?": "Pass?"})
        st.dataframe(overall)

        st.download_button(
            "Download detailed CSV",
            data=merged.to_csv(index=False).encode("utf-8"),
            file_name=f"schaper_tieout_{(case_name or 'report').replace(' ', '_')}.csv",
            mime="text/csv",
        )
else:
    st.info("Upload at least one PDF to begin.")
