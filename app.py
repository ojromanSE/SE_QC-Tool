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
    "Schaper-format parser: Table 1.1, Cash-Flow pages (by header), and One-line Summary totals. "
    "Cash-flows: read NET OIL PROD / NET GAS PROD / NET NGL PROD and PV10; BOE only if a NET EQUIV header is present."
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

# ---------------- Schaper PDF parsing ----------------
# Cash-flow page: category header like: "SE_RSV_CAT = 1PDP"
RSV_CAT_PAT = re.compile(r"(?i)SE[_\s]*RSV[_\s]*CAT\s*[:=]\s*(1PDP|4PUD|5PROB|6POSS)")

# Table 1.1 rows: Gas (MMcf), NGL (Mbbl), Oil (Mbbl), BOE (Mboe), Undisc ($MM), PV10 ($MM)
TABLE11_ROW_PAT = re.compile(
    r"(?i)(Total\s+Proved\s+Reserves|Proved\s+Developed\s+Producing\s+\(1PDP\)|Proved\s+Undeveloped\s+\(4PUD\)|Total\s+Probable\s+Reserves\s+\(5PROB\)|Total\s+Possible\s+Reserves\s+\(6POSS\)).*?"
    r"([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+\$\s*([0-9,]+)\s+\$\s*([0-9,]+)"
)

# One-line Summary TOTAL rows: OIL, GAS, NGL, BOE, PV10 ($); convert PV10 to $MM
ONELINE_TOTAL_PAT = re.compile(
    r"(?im)^\s*TOTAL\s+(1PDP|4PUD|5PROB|6POSS)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9,]+)\s*$"
)

# Last two numbers on TOTAL line -> (Undisc, PV10) for cash-flow pages
TOTAL_LAST_TWO = re.compile(r"(?mi)^\s*TOTAL\b[^\n]*?(-?\d[\d,]*\.?\d*)\s+(-?\d[\d,]*\.?\d*)\s*$")

METRICS = [
    "Oil (Mbbl)",
    "Gas (MMcf)",
    "NGL (Mbbl)",
    "Net BOE (Mboe)",
    "PV10 ($MM)",
]

def _nearest_by_x(target_x, numeric_words):
    """Pick numeric word whose x-center is nearest to target_x."""
    best = None
    best_d = 1e9
    for w in numeric_words:
        xcen = (w["x0"] + w["x1"]) / 2.0
        d = abs(target_x - xcen)
        if d < best_d:
            best_d = d
            best = w
    return best, best_d

def _extract_cashflow_totals_from_page(page):
    """
    Use pdfplumber word positions to map the TOTAL row values to the headers
    'NET OIL PROD', 'NET GAS PROD', 'NET NGL PROD' (and optionally 'NET EQUIV').
    """
    words = page.extract_words(
        use_text_flow=True,
        keep_blank_chars=False,
        extra_attrs=["x0", "x1", "top", "bottom"]
    )
    if not words:
        return {}

    # Find last TOTAL token -> y position of totals row
    total_words = [w for w in words if w["text"].strip().upper() == "TOTAL"]
    if not total_words:
        return {}
    tot_word = max(total_words, key=lambda w: w["top"])
    tot_y = tot_word["top"]

    # Headers above the TOTAL row
    headers = [w for w in words if w["top"] < tot_y - 5]

    def find_header_x(label):
        """Find the x-center of a token matching exactly the label."""
        lbl = label.upper()
        matches = [w for w in headers if w["text"].strip().upper() == lbl]
        if not matches:
            return None
        h = min(matches, key=lambda w: w["top"])  # choose the topmost
        return (h["x0"] + h["x1"]) / 2.0

    x_oil = find_header_x("NET OIL PROD")
    x_gas = find_header_x("NET GAS PROD")
    x_ngl = find_header_x("NET NGL PROD")
    x_boe = find_header_x("NET EQUIV")  # optional; only present in some layouts

    # Numeric words on TOTAL row
    numeric_on_total = [
        w for w in words
        if abs(w["top"] - tot_y) < 3 and NUM_RE.match(w["text"].strip())
    ]
    if not numeric_on_total:
        return {}

    def nearest_val(x_target):
        if x_target is None or not numeric_on_total:
            return math.nan
        best = min(numeric_on_total, key=lambda w: abs(((w["x0"]+w["x1"])/2) - x_target))
        return _to_f(best["text"])

    out = {
        "Oil (Mbbl)": nearest_val(x_oil),
        "Gas (MMcf)": nearest_val(x_gas),
        "NGL (Mbbl)": nearest_val(x_ngl),
    }
    if x_boe is not None:
        out["Net BOE (Mboe)"] = nearest_val(x_boe)  # only if header exists
    return out

def parse_pdf_schaper(file_obj):
    """Parse a Schaper-format reserves PDF into rows of [Source, Category, metrics...]."""
    if not HAS_PDFPLUMBER:
        return None, "pdfplumber is not installed"

    rows = []
    try:
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""

                # ---- Table 1.1 on any page ----
                for m in TABLE11_ROW_PAT.finditer(text):
                    label = m.group(1)
                    gas = _to_f(m.group(2))     # MMcf
                    ngl = _to_f(m.group(3))     # Mbbl
                    oil = _to_f(m.group(4))     # Mbbl
                    boe = _to_f(m.group(5))     # Mboe
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

                # ---- Cash-flow page: detect category, pull volumes & PV10 ----
                mcat = RSV_CAT_PAT.search(text)
                if mcat:
                    cat = mcat.group(1)

                    # PV10 from last two numbers on TOTAL line (second is PV10)
                    pv10_cf = math.nan
                    last_two = None
                    for m in TOTAL_LAST_TWO.finditer(text):
                        last_two = m
                    if last_two:
                        pv10_cf = _to_f(last_two.group(2))

                    # Volumes from header-aligned TOTAL row (no BOE unless NET EQUIV header exists)
                    cf_vals = _extract_cashflow_totals_from_page(page)
                    rows.append(
                        {
                            "Source": "Cash Flows",
                            "Category": cat,
                            "Oil (Mbbl)": cf_vals.get("Oil (Mbbl)", math.nan),
                            "Gas (MMcf)": cf_vals.get("Gas (MMcf)", math.nan),
                            "NGL (Mbbl)": cf_vals.get("NGL (Mbbl)", math.nan),
                            "Net BOE (Mboe)": cf_vals.get("Net BOE (Mboe)", math.nan),
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
    out = []
    for cat in sorted(df["Category"].unique()):
        for metric in METRICS:
            vals = df.loc[df["Category"] == cat, metric].dropna().tolist()
            if strict and len(vals) < df["Source"].nunique():
                ok = False
            else:
                ok = within_tolerance(vals, abs_tol, rel_tol_pct) if vals else False

            status = "✅" if ok else "❌"   # green check / red cross

            out.append(
                {
                    "Category": cat,
                    "Metric": metric,
                    "Sources": int(df.loc[(df["Category"] == cat) & df[metric].notna()].shape[0]),
                    "Min": pd.Series(vals).min() if vals else math.nan,
                    "Max": pd.Series(vals).max() if vals else math.nan,
                    "Consistent?": status,
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
