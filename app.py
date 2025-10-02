import re
import math
import io
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
    "Schaper-format parser with cross-checks across PDF (Table 1.1 / Cash Flows / One-line) "
    "and two Excel files (Oneline + Monthly by SE_RSV_CAT). Green ✅ / Red ❌ indicate consistency."
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

def _norm_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _find_col(df, patterns):
    """Find first column whose name matches any regex in patterns (case-insensitive)."""
    cols = list(df.columns)
    for pat in patterns:
        cre = re.compile(pat, re.I)
        for c in cols:
            if cre.search(str(c)):
                return c
    return None

def _sum_numeric(series):
    return pd.to_numeric(series, errors="coerce").sum(skipna=True)

# ---------------- Regex patterns (PDF) ----------------
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

# ---------------- Cashflow extraction (PDF) ----------------
def _nearest_by_x(target_x, numeric_words):
    best, best_d = None, 1e9
    for w in numeric_words:
        xcen = (w["x0"] + w["x1"]) / 2.0
        d = abs(target_x - xcen)
        if d < best_d:
            best, best_d = w, d
    return best

def _extract_cashflow_totals_from_page(page):
    """Map TOTAL row values to NET OIL PROD / NET GAS PROD / NET NGL PROD headers. BOE only if NET EQUIV exists."""
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
    if not numeric_on_total:
        return {}

    def nearest_val(x_target):
        if x_target is None or not numeric_on_total:
            return math.nan
        w = _nearest_by_x(x_target, numeric_on_total)
        return _to_f(w["text"]) if w else math.nan

    out = {
        "Oil (Mbbl)": nearest_val(x_oil),
        "Gas (MMcf)": nearest_val(x_gas),
        "NGL (Mbbl)": nearest_val(x_ngl),
    }
    if x_boe is not None:
        out["Net BOE (Mboe)"] = nearest_val(x_boe)  # only if header exists
    return out

# ---------------- PDF parser ----------------
def parse_pdf_schaper(file_obj):
    """Parse PDF into rows of [Source, Category, metrics...]."""
    if not HAS_PDFPLUMBER:
        return None, "pdfplumber is not installed"

    rows = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            # Table 1.1 on any page
            for m in TABLE11_ROW_PAT.finditer(text):
                label = m.group(1)
                gas = _to_f(m.group(2))     # MMcf
                ngl = _to_f(m.group(3))     # Mbbl
                oil = _to_f(m.group(4))     # Mbbl
                boe = _to_f(m.group(5))     # Mboe
                pv10 = _to_f(m.group(7))    # $MM

                key = None
                if "Developed" in label: key = "1PDP"
                elif "Undeveloped" in label: key = "4PUD"
                elif "Probable" in label: key = "5PROB"
                elif "Possible" in label: key = "6POSS"
                elif "Total Proved Reserves" in label: key = "TOTAL PROVED"
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

            # Cash-flow page
            mcat = RSV_CAT_PAT.search(text)
            if mcat:
                cat = mcat.group(1)
                pv10_cf = math.nan
                last_two = None
                for m in TOTAL_LAST_TWO.finditer(text):
                    last_two = m
                if last_two:
                    pv10_cf = _to_f(last_two.group(2))
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

            # One-line summary totals
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

    if not rows:
        return None, "No recognizable sections found."
    return pd.DataFrame(rows), None

# ---------------- Excel: Oneline parser ----------------
def parse_oneline_xlsx(file):
    """
    Expect a column like SE_RSV_CAT (reserve category).
    Oil/Gas/NGL/BOE/PV columns may be named a variety of ways; we detect by regex.
    Returns rows per Category with totals.
    """
    xls = pd.ExcelFile(file)
    frames = []
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        if df.empty:
            continue
        df = _norm_columns(df)

        cat_col = _find_col(df, [r"\bSE[_\s-]*RSV[_\s-]*CAT\b", r"\bRESERVE\s*CAT", r"\bCATEGORY\b"])
        if not cat_col:
            continue

        oil_col = _find_col(df, [r"\bNET\s*OIL\b", r"\bOIL\b"])
        gas_col = _find_col(df, [r"\bNET\s*GAS\b", r"\bGAS\b"])
        ngl_col = _find_col(df, [r"\bNET\s*NGL\b", r"\bNGL\b"])
        boe_col = _find_col(df, [r"\bNET\s*BOE\b", r"\bEQUIV\b"])
        # PV could be $ or $MM; we normalize to $MM below
        pv_col = _find_col(df, [r"\bPV\s*10\b", r"\bPRESENT\s*VALUE", r"\bPV\b"])

        cols = { "Category": cat_col, "Oil (Mbbl)": oil_col, "Gas (MMcf)": gas_col, "NGL (Mbbl)": ngl_col,
                 "Net BOE (Mboe)": boe_col, "PV_raw": pv_col }
        # keep only existing
        use = [c for c in cols.values() if c]
        if not use:
            continue
        slim = df[list(set(use))].copy()
        # coerce
        for c in [oil_col, gas_col, ngl_col, boe_col, pv_col]:
            if c and c in slim:
                slim[c] = pd.to_numeric(slim[c], errors="coerce")
        # aggregate
        grouped = slim.groupby(df[cat_col]).agg(_sum_numeric).reset_index()
        grouped.rename(columns={cat_col: "Category"}, inplace=True)
        # normalize PV to $MM if very large (assume dollars)
        if pv_col and pv_col in grouped:
            pv_vals = grouped[pv_col]
            if pv_vals.max(skipna=True) and pv_vals.max(skipna=True) > 1_000_000:
                grouped["PV10 ($MM)"] = grouped[pv_col] / 1_000_000.0
            else:
                grouped["PV10 ($MM)"] = grouped[pv_col]
        # rename volume cols
        if oil_col and oil_col in grouped: grouped["Oil (Mbbl)"] = grouped[oil_col]
        if gas_col and gas_col in grouped: grouped["Gas (MMcf)"] = grouped[gas_col]
        if ngl_col and ngl_col in grouped: grouped["NGL (Mbbl)"] = grouped[ngl_col]
        if boe_col and boe_col in grouped: grouped["Net BOE (Mboe)"] = grouped[boe_col]
        # build rows
        for _, r in grouped.iterrows():
            rows = {
                "Source": "Oneline XLS",
                "Category": str(r.get("Category")).strip(),
                "Oil (Mbbl)": r.get("Oil (Mbbl)", math.nan),
                "Gas (MMcf)": r.get("Gas (MMcf)", math.nan),
                "NGL (Mbbl)": r.get("NGL (Mbbl)", math.nan),
                "Net BOE (Mboe)": r.get("Net BOE (Mboe)", math.nan),
                "PV10 ($MM)": r.get("PV10 ($MM)", math.nan),
            }
            frames.append(rows)
    return pd.DataFrame(frames) if frames else pd.DataFrame(columns=["Source","Category"]+METRICS)

# ---------------- Excel: Monthly parser ----------------
def parse_monthly_xlsx(file):
    """
    Expect monthly rows with a category column SE_RSV_CAT and columns:
    'NET OIL PROD', 'NET GAS PROD', 'NET NGL PROD' (case-insensitive).
    We group by category and sum volumes; PV10 is not expected here.
    """
    xls = pd.ExcelFile(file)
    frames = []
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        if df.empty:
            continue
        df = _norm_columns(df)

        cat_col = _find_col(df, [r"\bSE[_\s-]*RSV[_\s-]*CAT\b", r"\bRESERVE\s*CAT", r"\bCATEGORY\b"])
        if not cat_col:
            continue

        oil_col = _find_col(df, [r"\bNET\s*OIL\s*PROD\b"])
        gas_col = _find_col(df, [r"\bNET\s*GAS\s*PROD\b"])
        ngl_col = _find_col(df, [r"\bNET\s*NGL\s*PROD\b"])
        boe_col = _find_col(df, [r"\bNET\s*(BOE|EQUIV)\b"])  # optional, may not exist

        cols = [c for c in [cat_col, oil_col, gas_col, ngl_col, boe_col] if c]
        if not cols:
            continue

        slim = df[cols].copy()
        for c in [oil_col, gas_col, ngl_col, boe_col]:
            if c and c in slim:
                slim[c] = pd.to_numeric(slim[c], errors="coerce")

        grouped = slim.groupby(df[cat_col]).agg(_sum_numeric).reset_index()
        grouped.rename(columns={cat_col: "Category"}, inplace=True)

        for _, r in grouped.iterrows():
            frames.append({
                "Source": "Monthly XLS",
                "Category": str(r.get("Category")).strip(),
                "Oil (Mbbl)": r.get(oil_col, math.nan) if oil_col else math.nan,
                "Gas (MMcf)": r.get(gas_col, math.nan) if gas_col else math.nan,
                "NGL (Mbbl)": r.get(ngl_col, math.nan) if ngl_col else math.nan,
                "Net BOE (Mboe)": r.get(boe_col, math.nan) if boe_col else math.nan,
                "PV10 ($MM)": math.nan,  # monthly usually doesn't have PV
            })

    return pd.DataFrame(frames) if frames else pd.DataFrame(columns=["Source","Category"]+METRICS)

# ---------------- Consistency table ----------------
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
left, right = st.columns(2)
with left:
    pdf_files = st.file_uploader("Upload **PDF** report(s)", type=["pdf"], accept_multiple_files=True)
with right:
    oneline_xls = st.file_uploader("Upload **Oneline XLS**", type=["xls", "xlsx"], accept_multiple_files=False)

monthly_xls = st.file_uploader("Upload **Monthly XLS**", type=["xls", "xlsx"], accept_multiple_files=False, key="monthly")

frames = []

# Parse PDF(s)
if pdf_files:
    with st.spinner("Parsing PDF(s)..."):
        for f in pdf_files:
            df, err = parse_pdf_schaper(f)
            if err:
                st.error(f"{f.name}: {err}")
            else:
                df.insert(0, "File", f.name)
                frames.append(df)

# Parse Oneline XLS
if oneline_xls:
    with st.spinner("Parsing Oneline XLS..."):
        try:
            df = parse_oneline_xlsx(oneline_xls)
            if not df.empty:
                df.insert(0, "File", getattr(oneline_xls, "name", "Oneline.xlsx"))
                frames.append(df)
            else:
                st.warning("Oneline XLS: no recognizable sheets/columns found.")
        except Exception as e:
            st.error(f"Oneline XLS parse error: {e}")

# Parse Monthly XLS
if monthly_xls:
    with st.spinner("Parsing Monthly XLS..."):
        try:
            df = parse_monthly_xlsx(monthly_xls)
            if not df.empty:
                df.insert(0, "File", getattr(monthly_xls, "name", "Monthly.xlsx"))
                frames.append(df)
            else:
                st.warning("Monthly XLS: no recognizable sheets/columns found.")
        except Exception as e:
            st.error(f"Monthly XLS parse error: {e}")

if frames:
    merged = pd.concat(frames, ignore_index=True)

    st.subheader("Extracted figures (all sources)")
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
    overall = []
    for file, group in results.groupby("File"):
        ok = group["Consistent?"].eq("✅").all()
        overall.append({"File": file, "Pass?": "✅" if ok else "❌"})
    st.dataframe(pd.DataFrame(overall))

    st.download_button(
        "Download detailed CSV",
        data=merged.to_csv(index=False).encode("utf-8"),
        file_name=f"schaper_tieout_{(case_name or 'report').replace(' ', '_')}.csv",
        mime="text/csv",
    )
else:
    st.info("Upload at least one PDF and/or XLS to begin.")
