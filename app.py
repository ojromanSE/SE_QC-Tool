import re
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
st.caption(
    "Schaper-format parser: Table 1.1, Cash-Flow pages (detected by header), and One-line Summary totals. "
    "Now ties out Oil, Gas, NGL, BOE, and PV10 across sections."
)

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("Options")
    abs_tol = st.number_input(
        "Absolute tolerance", min_value=0.0, value=0.5, step=0.1,
        help="Allowed absolute difference (units: Oil/NGL in Mbbl, Gas in MMcf, BOE in Mboe, PV in $MM)."
    )
    rel_tol_pct = st.number_input(
        "Relative tolerance (%)", min_value=0.0, value=0.1, step=0.05,
        help="Allowed percent difference across sources."
    )
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

# Table 1.1 rows: capture Gas (MMcf), NGL (Mbbls), Oil (Mbbls), Equivalent (Mboe), Undisc ($MM), PV10 ($MM)
TABLE11_ROW_PAT = re.compile(
    r"(?i)(Total\s+Proved\s+Reserves|Proved\s+Developed\s+Producing\s+\(1PDP\)|Proved\s+Undeveloped\s+\(4PUD\)|Total\s+Probable\s+Reserves\s+\(5PROB\)|Total\s+Possible\s+Reserves\s+\(6POSS\)).*?"
    r"([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+\$\s*([0-9,]+)\s+\$\s*([0-9,]+)"
)

# Cash-flow TOTAL line: we’ll slice numeric columns by header positions; still capture the last two numeric tokens as (Undisc, PV10)
TOTAL_LINE_LAST_TWO = re.compile(r"(?mi)^\s*TOTAL\b[^\n]*?(-?\d[\d,]*\.?\d*)\s+(-?\d[\d,]*\.?\d*)\s*$")

# One-line Summary TOTAL rows: Oil, Gas, NGL, BOE, PV10($); convert PV10 to $MM
ONELINE_TOTAL_PAT = re.compile(
    r"(?im)^\s*TOTAL\s+(1PDP|4PUD|5PROB|6POSS)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9,]+)\s*$"
)

METRICS = [
    "Oil (Mbbl)",
    "Gas (MMcf)",
    "NGL (Mbbl)",
    "Net BOE (Mboe)",
    "PV10 ($MM)",
]

def _slice_last_number(text_segment: str):
    m = re.findall(r"-?\d[\d,]*\.?\d*", text_segment)
    return _to_f(m[-1]) if m else math.nan

def _extract_cashflow_totals_by_columns(lines):
    """
    From a cash-flow page, find the header line (with NET OIL/GAS/NGL/EQUIV),
    detect start indices for those columns, then read the last TOTAL row and slice the numbers.
    Returns dict with Oil (Mbbl), Gas (MMcf), NGL (Mbbl), Net BOE (Mboe).
    """
    header_idx = None
    header_line = ""
    for i, ln in enumerate(lines):
        if ("NET OIL" in ln and "NET GAS" in ln and "NET NGL" in ln and "NET EQUIV" in ln):
            header_idx = i
            header_line = ln
            break
    if header_idx is None:
        return {}

    # Column start indices (robust to variable spacing/monospace layout)
    pos_oil = header_line.find("NET OIL")
    pos_gas = header_line.find("NET GAS")
    pos_ngl = header_line.find("NET NGL")
    pos_boe = header_line.find("NET EQUIV")
    if min(pos_oil, pos_gas, pos_ngl, pos_boe) < 0:
        return {}

    # Find the last TOTAL line on this page
    total_idx = None
    for i, ln in enumerate(lines):
        if ln.strip().startswith("TOTAL"):
            total_idx = i  # keep last
    if total_idx is None:
        return {}

    total_line = lines[total_idx]

    # Slice segments based on header starts
    seg_oil = total_line[pos_oil:pos_gas]
    seg_gas = total_line[pos_gas:pos_ngl]
    seg_ngl = total_line[pos_ngl:pos_boe]
    seg_boe = total_line[pos_boe:]  # until end

    oil = _slice_last_number(seg_oil)
    gas = _slice_last_number(seg_gas)
    ngl = _slice_last_number(seg_ngl)
    boe = _slice_last_number(seg_boe)

    return {
        "Oil (Mbbl)": oil,
        "Gas (MMcf)": gas,
        "NGL (Mbbl)": ngl,
        "Net BOE (Mboe)": boe,
    }

def parse_pdf_schaper(file_obj):
    """Parse a Schaper-format reserves PDF into rows of [Source, Category, metrics...]."""
    if not HAS_PDFPLUMBER:
        return None, "pdfplumber is not installed"

    rows = []
    try:
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                lines = text.splitlines()

                # ---- Table 1.1 matches on any page ----
                for m in TABLE11_ROW_PAT.finditer(text):
                    label = m.group(1)
                    gas = _to_f(m.group(2))     # MMcf
                    ngl = _to_f(m.group(3))     # Mbbl
                    oil = _to_f(m.group(4))     # Mbbl
                    boe = _to_f(m.group(5))     # Mboe
                    undisc = _to_f(m.group(6))  # $MM (unused here)
                    pv10 = _to_f(m.group(7))    # $MM

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
                            {
                                "Source": "Table1.1",
                                "Category": key,
                                "Oil (Mbbl)": oil,
                                "Gas (MMcf)": gas,
                                "NGL (Mbbl)": ngl,
                                "Net BOE (Mboe)": boe,
                                "PV10 ($MM)": pv10,
                            }
                        )

                # ---- Cash-flow page: detect category in header ----
                mcat = RSV_CAT_PAT.search(text)
                if mcat:
                    cat = mcat.group(1)

                    # Pull PV10 (last two numeric tokens on TOTAL line -> second is PV10)
                    last_two = None
                    for m in TOTAL_LINE_LAST_TWO.finditer(text):
                        last_two = m  # keep last TOTAL on page
                    pv10_cf = _to_f(last_two.group(2)) if last_two else math.nan

                    # Pull Oil/Gas/NGL/BOE totals using header column positions
                    col_totals = _extract_cashflow_totals_by_columns(lines)

                    rows.append(
                        {
                            "Source": "Cash Flows",
                            "Category": cat,
                            "Oil (Mbbl)": col_totals.get("Oil (Mbbl)", math.nan),
                            "Gas (MMcf)": col_totals.get("Gas (MMcf)", math.nan),
                            "NGL (Mbbl)": col_totals.get("NGL (Mbbl)", math.nan),
                            "Net BOE (Mboe)": col_totals.get("Net BOE (Mboe)", math.nan),
                            "PV10 ($MM)": pv10_cf,
                        }
                    )

                # ---- One-line Summary totals (OIL, GAS, NGL, NET BOE, PV10-$) ----
                for m in ONELINE_TOTAL_PAT.finditer(text):
                    key = m.group(1)
                    oil = _to_f(m.group(2))
                    gas = _to_f(m.group(3))
                    ngl = _to_f(m.group(4))
                    boe = _to_f(m.group(5))
                    pv10_dollars = _to_f(m.group(6))
                    rows.append(
                        {
                            "Source": "One-line",
                            "Category": key,
                            "Oil (Mbbl)": oil,
                            "Gas (MMcf)": gas,
                            "NGL (Mbbl)": ngl,
                            "Net BOE (Mboe)": boe,
                            "PV10 ($MM)": pv10_dollars / 1000.0,
                        }
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
        for metric in METRICS:
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
