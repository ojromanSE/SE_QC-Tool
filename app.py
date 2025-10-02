import re
import math
import pandas as pd
import streamlit as st

# PDF parsing
try:
    import pdfplumber  # type: ignore
    HAS_PDFPLUMBER = True
except Exception:
    HAS_PDFPLUMBER = False

st.set_page_config(page_title="Reserves Tie-Out Checker", layout="wide")
st.title("📊 Reserves Tie-Out Checker")
st.caption(
    "Schaper-format parser: Table 1.1, Cash-Flow pages (NET OIL/GAS/NGL PROD headers), and One-line Summary totals. "
    "Green ✅ / Red ❌ indicators show consistency."
)

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("Options")
    abs_tol = st.number_input(
        "Absolute tolerance", min_value=0.0, value=0.5, step=0.1,
        help="Oil/NGL: Mbbl | Gas: MMcf | BOE: Mboe | PV: $MM"
    )
    rel_tol_pct = st.number_input(
        "Relative tolerance (%)", min_value=0.0, value=0.1, step=0.05,
        help="Percent difference allowed across sources."
    )
    strict = st.checkbox("Strict: every source must have every field", value=False)
    case_name = st.text_input("Case/Project name", "")

# ---------------- Helpers ----------------
NUM_RE = re.compile(r"^-?\d[\d,]*\.?\d*$")

def _to_f(val):
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

# ---------------- Regex patterns ----------------
RSV_CAT_PAT = re.compile(r"(?i)SE[_\s]*RSV[_\s]*CAT\s*[:=]\s*(1PDP|4PUD|5PROB|6POSS)")

TABLE11_ROW_PAT = re.compile(
    r"(?i)(Total\s+Proved\s+Reserves|Proved\s+Developed\s+Producing\s+\(1PDP\)|Proved\s+Undeveloped\s+\(4PUD\)|Total\s+Probable\s+Reserves\s+\(5PROB\)|Total\s+Possible\s+Reserves\s+\(6POSS\)).*?"
    r"([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+\$\s*([0-9,]+)\s+\$\s*([0-9,]+)"
)

ONELINE_TOTAL_PAT = re.compile(
    r"(?im)^\s*TOTAL\s+(1PDP|4PUD|5PROB|6POSS)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9,]+)\s*$"
)

TOTAL_LAST_TWO = re.compile(r"(?mi)^\s*TOTAL\b[^\n]*?(-?\d[\d,]*\.?\d*)\s+(-?\d[\d,]*\.?\d*)\s*$")

METRICS = [
    "Oil (Mbbl)",
    "Gas (MMcf)",
    "NGL (Mbbl)",
    "Net BOE (Mboe)",
    "PV10 ($MM)",
]

# ---------------- Cashflow extraction ----------------
def _nearest_by_x(target_x, numeric_words):
    best, best_d = None, 1e9
    for w in numeric_words:
        xcen = (w["x0"] + w["x1"]) / 2.0
        d = abs(target_x - xcen)
        if d < best_d:
            best, best_d = w, d
    return best

def _extract_cashflow_totals_from_page(page):
    """Map TOTAL row values to NET OIL PROD / NET GAS PROD / NET NGL PROD headers."""
    words = page.extract_words(
        use_text_flow=True, keep_blank_chars=False, extra_attrs=["x0", "x1", "top", "bottom"]
    )
    if not words:
        return {}

    # Locate TOTAL row
    total_words = [w for w in words if w["text"].strip().upper() == "TOTAL"]
    if not total_words:
        return {}
    tot_word = max(total_words, key=lambda w: w["top"])
    tot_y = tot_word["top"]

    headers = [w for w in words if w["top"] < tot_y - 5]

    def find_header_x(label):
        lbl = label.upper()
        matches = [w for w in headers if w["text"].strip().upper() == lbl]
        if not matches:
            return None
        h = min(matches, key=lambda w: w["top"])
        return (h["x0"] + h["x1"]) / 2.0

    x_oil = find_header_x("NET OIL PROD")
    x_gas = find_header_x("NET GAS PROD")
    x_ngl = find_header_x("NET NGL PROD")
    x_boe = find_header_x("NET EQUIV")  # optional

    numeric_on_total = [
        w for w in words if abs(w["top"] - tot_y) < 3 and NUM_RE.match(w["text"].strip())
    ]
    out = {}
    def nearest_val(x_target):
        if x_target is None or not numeric_on_total:
            return math.nan
        w = _nearest_by_x(x_target, numeric_on_total)
        return _to_f(w["text"]) if w else math.nan

    out["Oil (Mbbl)"] = nearest_val(x_oil)
    out["Gas (MMcf)"] = nearest_val(x_gas)
    out["NGL (Mbbl)"] = nearest_val(x_ngl)
    if x_boe is not None:
        out["Net BOE (Mboe)"] = nearest_val(x_boe)
    return out

# ---------------- Parsing ----------------
def parse_pdf_schaper(file_obj):
    if not HAS_PDFPLUMBER:
        return None, "pdfplumber is not installed"

    rows = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            # Table 1.1
            for m in TABLE11_ROW_PAT.finditer(text):
                label = m.group(1)
                gas, ngl, oil, boe, pv10 = _to_f(m.group(2)), _to_f(m.group(3)), _to_f(m.group(4)), _to_f(m.group(5)), _to_f(m.group(7))
                key = None
                if "Developed" in label: key = "1PDP"
                elif "Undeveloped" in label: key = "4PUD"
                elif "Probable" in label: key = "5PROB"
                elif "Possible" in label: key = "6POSS"
                elif "Total Proved Reserves" in label: key = "TOTAL PROVED"
                if key:
                    rows.append({"Source":"Table1.1","Category":key,"Oil (Mbbl)":oil,"Gas (MMcf)":gas,"NGL (Mbbl)":ngl,"Net BOE (Mboe)":boe,"PV10 ($MM)":pv10})

            # Cash-flows
            mcat = RSV_CAT_PAT.search(text)
            if mcat:
                cat = mcat.group(1)
                pv10_cf = math.nan
                for m in TOTAL_LAST_TWO.finditer(text):
                    last_two = m
                if 'last_two' in locals():
                    pv10_cf = _to_f(last_two.group(2))
                cf_vals = _extract_cashflow_totals_from_page(page)
                rows.append({"Source":"Cash Flows","Category":cat,
                             "Oil (Mbbl)":cf_vals.get("Oil (Mbbl)",math.nan),
                             "Gas (MMcf)":cf_vals.get("Gas (MMcf)",math.nan),
                             "NGL (Mbbl)":cf_vals.get("NGL (Mbbl)",math.nan),
                             "Net BOE (Mboe)":cf_vals.get("Net BOE (Mboe)",math.nan),
                             "PV10 ($MM)":pv10_cf})

            # One-line
            for m in ONELINE_TOTAL_PAT.finditer(text):
                key = m.group(1)
                oil, gas, ngl, boe, pv10_dollars = _to_f(m.group(2)), _to_f(m.group(3)), _to_f(m.group(4)), _to_f(m.group(5)), _to_f(m.group(6))
                rows.append({"Source":"One-line","Category":key,
                             "Oil (Mbbl)":oil,"Gas (MMcf)":gas,"NGL (Mbbl)":ngl,
                             "Net BOE (Mboe)":boe,"PV10 ($MM)":pv10_dollars/1000.0})
    if not rows:
        return None, "No recognizable sections found."
    return pd.DataFrame(rows), None

# ---------------- Consistency ----------------
def check_consistency(df):
    out = []
    for cat in sorted(df["Category"].unique()):
        for metric in METRICS:
            vals = df.loc[df["Category"] == cat, metric].dropna().tolist()
            if strict and len(vals) < df["Source"].nunique():
                ok = False
            else:
                ok = within_tolerance(vals, abs_tol, rel_tol_pct) if vals else False
            status = "✅" if ok else "❌"
            out.append({"Category":cat,"Metric":metric,
                        "Sources":int(df.loc[(df["Category"]==cat)&df[metric].notna()].shape[0]),
                        "Min":pd.Series(vals).min() if vals else math.nan,
                        "Max":pd.Series(vals).max() if vals else math.nan,
                        "Consistent?":status})
    return pd.DataFrame(out)

# ---------------- UI ----------------
pdf_files = st.file_uploader("Upload Reserves PDF(s)", type=["pdf"], accept_multiple_files=True)

if pdf_files:
    frames=[]
    with st.spinner("Parsing..."):
        for f in pdf_files:
            df, err = parse_pdf_schaper(f)
            if err: st.error(f"{f.name}: {err}")
            else:
                df.insert(0,"File",f.name)
                frames.append(df)
    if frames:
        merged=pd.concat(frames,ignore_index=True)
        st.subheader("Extracted figures")
        st.dataframe(merged)

        st.subheader("Consistency checks (by file)")
        results=merged.groupby(["File"]).apply(lambda g:check_consistency(g)).reset_index(level=0).rename(columns={"level_0":"File"})
        st.dataframe(results)

        st.subheader("Overall")
        overall=[]
        for file,group in results.groupby("File"):
            ok = group["Consistent?"].eq("✅").all()
            overall.append({"File":file,"Pass?":"✅" if ok else "❌"})
        st.dataframe(pd.DataFrame(overall))

        st.download_button("Download detailed CSV",
                           data=merged.to_csv(index=False).encode("utf-8"),
                           file_name=f"schaper_tieout_{(case_name or 'report').replace(' ', '_')}.csv",
                           mime="text/csv")
else:
    st.info("Upload at least one PDF to begin.")
