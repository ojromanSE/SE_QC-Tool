import io
import re
import math
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd
import streamlit as st

# -------------------- PDF deps --------------------
try:
    import pdfplumber  # type: ignore
    HAS_PDFPLUMBER = True
except Exception:
    HAS_PDFPLUMBER = False

st.set_page_config(page_title="Reserves Cross-Check", layout="wide")
st.title("📊 Reserves Cross‑Check")
st.caption(
    "Parse and cross‑check PDF (Table 1.1 / Cash Flows / Oneline) + Excel (Oneline/Monthly). "
    "Green ✅ / Red ❌ indicate consistency. PV shown as PV9 / PV10 (M$ = thousands of dollars)."
)

# -------------------- Sidebar --------------------
with st.sidebar:
    st.header("Options")
    abs_tol = st.number_input(
        "Absolute tolerance", min_value=0.0, value=0.5, step=0.1,
        help="Oil/NGL in Mbbl, Gas in MMcf, BOE in Mboe, PV in M$ (thousands of dollars)"
    )
    rel_tol_pct = st.number_input(
        "Relative tolerance (%)", min_value=0.0, value=0.10, step=0.05,
        help="Percent difference allowed across sources"
    )
    strict = st.checkbox("Strict: each metric must appear in every source for a pass", value=False)
    case_name = st.text_input("Case/Project name (for CSVs)", "")

# -------------------- Helpers --------------------
GREEN_CHECK = "✅"
RED_X = "❌"
CAT_ORDER = ["TOTAL PROVED", "1PDP", "3PDNP", "4PUD", "5PROB", "6POSS"]

def check_mark(ok: bool) -> str:
    return GREEN_CHECK if ok else RED_X

def numberize(x):
    """Convert strings like '$ 1,234.56' or '1,234' to float; leave NaN if not parseable."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x)
    s = s.replace(",", "")
    s = re.sub(r"^\$?\s*", "", s)
    s = s.strip()
    try:
        return float(s)
    except Exception:
        return np.nan

def within_tolerance(vals: List[float], abs_tol: float, rel_tol_pct: float) -> bool:
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

def normalize_header(h: str) -> str:
    """
    Lowercase, strip units/punct/extra spaces so we can match flexible headers.
    'Net Res Gas (MMcf)' -> 'net res gas', 'NPV at 10%' -> 'npv at 10'
    """
    s = h or ""
    s = s.replace("\xa0", " ")
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"[%$]", "", s)
    s = re.sub(r"[_\-./]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

# Canonical metric names we use everywhere
METRICS = [
    "Oil (Mbbl)",
    "Gas (MMcf)",
    "NGL (Mbbl)",
    "Net BOE (Mboe)",
    "PV9 (M$)",
    "PV10 (M$)",
]

# -------------------- XLS header mapping --------------------
XLS_PATTERNS: Dict[str, List[str]] = {
    "rsv_cat": [r"\bse\s*_?\s*rsv\s*_?\s*cat\b", r"\breserve[s]?\s*category\b", r"\bcategory\b"],
    "oil":     [r"\bnet\s*res\s*oil\b", r"\bnet\s*oil\b", r"\bnet\s*oil\s*prod(uction)?\b"],
    "gas":     [r"\bnet\s*res\s*gas\b", r"\bnet\s*gas\b", r"\bnet\s*gas\s*prod(uction)?\b"],
    "ngl":     [r"\bnet\s*res\s*ngl\b", r"\bnet\s*ngl\b", r"\bnet\s*ngl\s*prod(uction)?\b"],
    "boe":     [r"\bnet\s*res\s*(mboe|boe)\b", r"\bnet\s*boe\b"],
    "npv9":    [r"\bnpv\s*at\s*9\b", r"\bnpv9\b", r"\bpv9\b"],
    "npv10":   [r"\bnpv\s*at\s*10\b", r"\bnpv10\b", r"\bpv10\b"],
}

def map_columns(df: pd.DataFrame) -> Dict[str, str]:
    """Return a map {canonical -> actual_column_name_in_df} using tolerant patterns."""
    norm_cols = {c: normalize_header(str(c)) for c in df.columns}
    mapping = {}
    for want, patterns in XLS_PATTERNS.items():
        for c, norm in norm_cols.items():
            if any(re.search(p, norm) for p in patterns):
                mapping[want] = c
                break
    return mapping

def read_all_sheets(uploaded_file) -> pd.DataFrame:
    """Read all sheets into one DataFrame (stacked)."""
    content = uploaded_file.read()
    bio = io.BytesIO(content)
    # Try default engine, then openpyxl
    for engine in [None, "openpyxl"]:
        try:
            dfs = pd.read_excel(bio, sheet_name=None, engine=engine)
            frames = []
            for _, sheet in (dfs or {}).items():
                if sheet is not None and not sheet.empty:
                    frames.append(sheet)
            if frames:
                return pd.concat(frames, ignore_index=True)
        except Exception:
            bio.seek(0)
    return pd.DataFrame()

def aggregate_xls_by_category(df: pd.DataFrame) -> pd.DataFrame:
    """
    Map columns, clean numerics, group by SE_RSV_CAT, sum desired metrics.
    Returns standardized columns aligned with METRICS (PV in M$).
    """
    if df.empty:
        return pd.DataFrame(columns=["Category"] + METRICS)

    colmap = map_columns(df)
    if "rsv_cat" not in colmap or "oil" not in colmap or "gas" not in colmap or "ngl" not in colmap:
        return pd.DataFrame(columns=["Category"] + METRICS)

    work = df.rename(columns={v: k for k, v in colmap.items()})

    for c in ["oil", "gas", "ngl", "boe", "npv9", "npv10"]:
        if c in work:
            work[c] = work[c].map(numberize)

    work["rsv_cat"] = work["rsv_cat"].astype(str).str.strip().str.upper()

    agg = (
        work.groupby("rsv_cat", dropna=True)[[c for c in ["oil", "gas", "ngl", "boe", "npv9", "npv10"] if c in work]]
        .sum(numeric_only=True)
        .reset_index()
        .rename(columns={"rsv_cat": "Category"})
    )

    # Normalize PV to M$ (input NPV columns are in $)
    if "npv9" in agg.columns:
        agg["PV9 (M$)"] = agg["npv9"] / 1_000.0
    else:
        agg["PV9 (M$)"] = np.nan
    if "npv10" in agg.columns:
        agg["PV10 (M$)"] = agg["npv10"] / 1_000.0
    else:
        agg["PV10 (M$)"] = np.nan

    # Standard metric names
    rename = {}
    if "oil" in agg.columns: rename["oil"] = "Oil (Mbbl)"
    if "gas" in agg.columns: rename["gas"] = "Gas (MMcf)"
    if "ngl" in agg.columns: rename["ngl"] = "NGL (Mbbl)"
    if "boe" in agg.columns: rename["boe"] = "Net BOE (Mboe)"
    agg = agg.rename(columns=rename)

    # Ensure all metric columns exist
    for m in METRICS:
        if m not in agg.columns:
            agg[m] = np.nan

    return agg[["Category"] + METRICS]

# -------------------- PDF extraction --------------------
# Table 1.1 rows (Gas, NGL, Oil, Mboe, $Undisc, PV10 M$)
TABLE11_ROW_PAT = re.compile(
    r"(?i)(Total\s+Proved\s+Reserves|Proved\s+Developed\s+Producing\s+\(1PDP\)|Proved\s+Developed\s+Non-?Producing\s+\(3PDNP\)|Proved\s+Undeveloped\s+\(4PUD\)|Total\s+Probable\s+Reserves\s+\(5PROB\)|Total\s+Possible\s+Reserves\s+\(6POSS\)).*?"
    r"([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+\$\s*([0-9,]+)\s+\$\s*([0-9,]+)"
)

# Cash flow page category tag
RSV_CAT_PAT = re.compile(r"(?i)SE[_\s]*RSV[_\s]*CAT\s*=\s*(1PDP|3PDNP|4PUD|5PROB|6POSS)")
# "P.W., M$" box where PV rates/values live
PV_BOX_HEADER = re.compile(r"(?i)P\.W\.,\s*M\$")

def _nearest_by_x(target_x, numeric_words):
    best, best_d = None, 1e9
    for w in numeric_words:
        xcen = (w["x0"] + w["x1"]) / 2.0
        d = abs(target_x - xcen)
        if d < best_d:
            best, best_d = w, d
    return best

def _find_header_center(words, vol_tot_y, token2):
    """
    Return the x-center of the header column for NET <token2> (token2 in {'OIL','GAS','NGL'}).
    Works whether headers are split as NET / OIL / PROD or joined as 'NET OIL'.
    """
    U = lambda s: str(s).strip().upper()
    headers = [w for w in words if w["top"] < vol_tot_y - 5]

    # Single token 'NET OIL'
    single = [w for w in headers if U(w["text"]) in (f"NET {token2}", f"NET\u00A0{token2}")]
    if single:
        w = min(single, key=lambda t: t["top"])
        return (w["x0"] + w["x1"]) / 2.0

    # Adjacent tokens 'NET' then '<token2>'
    nets = [w for w in headers if U(w["text"]) == "NET"]
    for n in nets:
        same = [w for w in headers if abs(w["top"] - n["top"]) < 2.5 and w["x0"] > n["x0"]]
        t2 = [w for w in same if U(w["text"]) == token2]
        if t2:
            w = t2[0]
            return (min(n["x0"], w["x0"]) + max(n["x1"], w["x1"])) / 2.0

    tokens = [w for w in headers if U(w["text"]) == token2]
    if tokens:
        w = min(tokens, key=lambda t: t["top"])
        return (w["x0"] + w["x1"]) / 2.0
    return None

def _extract_cashflow_totals_from_page(page) -> Dict[str, float]:
    """
    From a CF page:
      • NET OIL / NET GAS / NET NGL totals from the *upper* table's TOTAL row
      • PV9 / PV10 (M$) from the P.W., M$ box
    """
    words = page.extract_words(
        use_text_flow=True, keep_blank_chars=False, extra_attrs=["x0", "x1", "top", "bottom"]
    )
    if not words:
        return {}

    # Find econ header line ("CUM. DISC. FCF")
    by_line = {}
    for w in words:
        y = round(w["top"], 1)
        by_line.setdefault(y, []).append(w)
    lines = []
    for y, ws in by_line.items():
        ws_sorted = sorted(ws, key=lambda t: t["x0"])
        text = " ".join([t["text"] for t in ws_sorted]).upper()
        lines.append((y, ws_sorted, text))
    lines.sort(key=lambda t: t[0])

    econ_hdr_y = None
    for y, _, txt in lines:
        if "CUM" in txt and "DISC" in txt and "FCF" in txt:
            econ_hdr_y = y
            break

    # Volumes TOTAL: last TOTAL above economics header
    totals = [w for w in words if w["text"].strip().upper() == "TOTAL"]
    if not totals:
        return {}
    vol_tot_y = None
    if econ_hdr_y is not None:
        above = [w for w in totals if w["top"] < econ_hdr_y - 1]
        if above:
            vol_tot_y = max(above, key=lambda w: w["top"])["top"]
    if vol_tot_y is None:
        vol_tot_y = min(totals, key=lambda w: w["top"])["top"]

    # Header centers (NET OIL / NET GAS / NET NGL)
    x_oil = _find_header_center(words, vol_tot_y, "OIL")
    x_gas = _find_header_center(words, vol_tot_y, "GAS")
    x_ngl = _find_header_center(words, vol_tot_y, "NGL")

    # Numeric words on the volumes TOTAL line
    nums_on_total = [
        w for w in words if abs(w["top"] - vol_tot_y) < 2.5 and re.match(r"^-?\d[\d,]*\.?\d*$", w["text"].strip())
    ]

    def at_col(xc):
        if xc is None or not nums_on_total:
            return np.nan
        w = _nearest_by_x(xc, nums_on_total)
        return numberize(w["text"]) if w else np.nan

    oil = at_col(x_oil)
    gas = at_col(x_gas)
    ngl = at_col(x_ngl)

    # PV9 & PV10 from 'P.W., M$' box
    pv9 = np.nan
    pv10 = np.nan
    text = page.extract_text() or ""
    m = PV_BOX_HEADER.search(text)
    if m:
        tail = text[m.start():]
        # find rate/value pairs like '9.00 479.064' or '10.00 2199.233'
        for rate_str, val_str in re.findall(r"(\d{1,2}(?:\.\d+)?)\s+([0-9][\d,]*\.?\d+)", tail):
            r = numberize(rate_str)
            v = numberize(val_str)
            if abs(r - 9.0) < 0.05:
                pv9 = v
            elif abs(r - 10.0) < 0.05:
                pv10 = v

    return {"Oil (Mbbl)": oil, "Gas (MMcf)": gas, "NGL (Mbbl)": ngl, "PV9 (M$)": pv9, "PV10 (M$)": pv10}

def parse_pdf_to_rows(file_bytes: bytes) -> pd.DataFrame:
    """Return rows from Table 1.1, Cash Flows, Oneline PDF for a PDF file."""
    rows = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            # Table 1.1
            for m in TABLE11_ROW_PAT.finditer(text):
                label = m.group(1)
                gas = numberize(m.group(2))
                ngl = numberize(m.group(3))
                oil = numberize(m.group(4))
                mboe = numberize(m.group(5))
                pv10_m = numberize(m.group(7))  # already M$

                key = None
                if "Developed Producing" in label: key = "1PDP"
                elif "Non-Prod" in label or "NonProducing" in label or "3PDNP" in label: key = "3PDNP"
                elif "Undeveloped" in label: key = "4PUD"
                elif "Total Proved Reserves" in label: key = "TOTAL PROVED"
                elif "Probable" in label: key = "5PROB"
                elif "Possible" in label: key = "6POSS"
                if key:
                    rows.append({
                        "Source": "Table1.1",
                        "Category": key,
                        "Oil (Mbbl)": oil,
                        "Gas (MMcf)": gas,
                        "NGL (Mbbl)": ngl,
                        "Net BOE (Mboe)": mboe,
                        "PV9 (M$)": np.nan,
                        "PV10 (M$)": pv10_m,
                    })

            # Cash Flows by SE_RSV_CAT
            mcat = RSV_CAT_PAT.search(text)
            if mcat:
                cat = mcat.group(1)
                cf = _extract_cashflow_totals_from_page(page)
                rows.append({
                    "Source": "Cash Flows",
                    "Category": cat,
                    "Oil (Mbbl)": cf.get("Oil (Mbbl)", np.nan),
                    "Gas (MMcf)": cf.get("Gas (MMcf)", np.nan),
                    "NGL (Mbbl)": cf.get("NGL (Mbbl)", np.nan),
                    "Net BOE (Mboe)": np.nan,
                    "PV9 (M$)": cf.get("PV9 (M$)", np.nan),
                    "PV10 (M$)": cf.get("PV10 (M$)", np.nan),
                })

            # Oneline PDF (grey category totals)
            for line in text.splitlines():
                s = re.sub(r"\s+", " ", line.strip())
                m = re.match(
                    r"^(1PDP|3PDNP|4PUD|5PROB|6POSS)\s+([-\d,]+)\s+([-\d,]+)\s+([-\d,]+)\s+([-\d,]+).*?([-\d,]+)\s*$",
                    s
                )
                if m:
                    cat = m.group(1).upper()
                    oil = numberize(m.group(2)); gas = numberize(m.group(3)); ngl = numberize(m.group(4))
                    mboe = numberize(m.group(5)); pv10_dollars = numberize(m.group(6))
                    rows.append({
                        "Source": "Oneline PDF",
                        "Category": cat,
                        "Oil (Mbbl)": oil,
                        "Gas (MMcf)": gas,
                        "NGL (Mbbl)": ngl,
                        "Net BOE (Mboe)": mboe,
                        "PV9 (M$)": np.nan,
                        "PV10 (M$)": pv10_dollars / 1_000.0,  # $ -> M$
                    })

    return pd.DataFrame(rows)

# -------------------- Pivots & Views --------------------
def build_category_pivot(df: pd.DataFrame, index_cols=("File", "Source")) -> pd.DataFrame:
    """
    Wide table with Categories as columns and (Metric, Category) as MultiIndex columns.
    """
    blocks = {}
    for m in METRICS:
        p = df.pivot_table(index=list(index_cols), columns="Category", values=m, aggfunc="first")
        if not p.empty:
            cols = list(p.columns)
            ordered = [c for c in CAT_ORDER if c in cols] + [c for c in cols if c not in CAT_ORDER]
            p = p.reindex(columns=ordered)
        blocks[m] = p
    wide = pd.concat(blocks, axis=1)
    return wide.round(3)

def build_slice_matrix(df: pd.DataFrame, category: str, metric: str,
                       row_dim: str = "Source", col_dim: str = "File",
                       abs_tol_val: float = 0.5, rel_tol_val: float = 0.1) -> pd.DataFrame:
    """
    One metric for one category, shown side-by-side by document (columns).
    Adds Min/Max/Δ and ✅/❌ per row.
    """
    use = df[df["Category"] == category].copy()
    if use.empty:
        return pd.DataFrame()

    pivot = use.pivot_table(index=[row_dim], columns=[col_dim], values=metric, aggfunc="first")
    # Sort columns alphabetically by file name for visual consistency
    pivot = pivot.reindex(sorted(pivot.columns, key=lambda t: str(t)), axis=1)

    # Min / Max / Delta across columns
    result = pivot.copy()
    result["Min"] = pivot.min(axis=1, skipna=True)
    result["Max"] = pivot.max(axis=1, skipna=True)
    result["Δ"] = (result["Max"] - result["Min"]).abs()

    def ok_row(row):
        vals = row[pivot.columns].tolist()
        return within_tolerance(vals, abs_tol_val, rel_tol_val)

    result["Consistent?"] = result.apply(ok_row, axis=1).map(lambda b: GREEN_CHECK if b else RED_X)
    return result.round(3)

# -------------------- Upload UI --------------------
st.subheader("Upload files")
pdf_files = st.file_uploader("PDF report(s)", type=["pdf"], accept_multiple_files=True)
oneline_files = st.file_uploader("Oneline Report XLS/XLSX (you can add several)", type=["xls", "xlsx"], accept_multiple_files=True, key="one")
monthly_files = st.file_uploader("Monthly Report XLS/XLSX (you can add several)", type=["xls", "xlsx"], accept_multiple_files=True, key="mon")

# -------------------- Parse all files --------------------
all_rows: List[pd.DataFrame] = []

# PDFs
if pdf_files:
    with st.spinner("Parsing PDF(s)…"):
        for f in pdf_files:
            try:
                rows = parse_pdf_to_rows(f.read())
                if not rows.empty:
                    rows.insert(0, "File", f.name)
                    all_rows.append(rows)
            except Exception as e:
                st.error(f"{f.name}: PDF parse error — {e}")

# Oneline XLS
if oneline_files:
    with st.spinner("Parsing Oneline XLS…"):
        for f in oneline_files:
            try:
                df = read_all_sheets(f)
                agg = aggregate_xls_by_category(df)
                if not agg.empty:
                    agg.insert(0, "Source", "Oneline XLS")
                    agg.insert(0, "File", f.name)
                    all_rows.append(agg)
                else:
                    st.warning(f"{f.name}: Oneline XLS — columns not recognized.")
            except Exception as e:
                st.error(f"{f.name}: Oneline XLS parse error — {e}")

# Monthly XLS
if monthly_files:
    with st.spinner("Parsing Monthly XLS…"):
        for f in monthly_files:
            try:
                df = read_all_sheets(f)
                agg = aggregate_xls_by_category(df)  # reuses same detector (looks for Net Oil/Gas/NGL)
                if not agg.empty:
                    agg.insert(0, "Source", "Monthly XLS")
                    agg.insert(0, "File", f.name)
                    all_rows.append(agg)
                else:
                    st.warning(f"{f.name}: Monthly XLS — columns not recognized.")
            except Exception as e:
                st.error(f"{f.name}: Monthly XLS parse error — {e}")

if not all_rows:
    st.info("Upload at least one PDF and/or XLS to begin.")
    st.stop()

merged = pd.concat(all_rows, ignore_index=True)

# Ensure all metric columns exist
for m in METRICS:
    if m not in merged.columns:
        merged[m] = np.nan

# -------------------- Views --------------------
st.subheader("Extracted figures (all sources)")
st.dataframe(merged, use_container_width=True, height=420)

# ---- Categories-as-columns pivot
st.subheader("Pivot view (Categories as columns)")
row_choice = st.radio(
    "Rows in pivot",
    options=["File & Source", "Source only", "File only"],
    horizontal=True,
    key="pivot_rows_choice"
)
if row_choice == "File & Source":
    idx = ("File", "Source")
elif row_choice == "Source only":
    idx = ("Source",)
else:
    idx = ("File",)

pivot_df = build_category_pivot(merged, index_cols=idx)
st.dataframe(pivot_df, use_container_width=True, height=420)
st.download_button(
    "Download pivot CSV",
    data=pivot_df.to_csv().encode("utf-8"),
    file_name=f"pivot_by_category_{(case_name or 'report').replace(' ', '_')}.csv",
    mime="text/csv",
)

# ---- Slice & Compare (your new side‑by‑side view)
st.subheader("Slice & Compare (one Category × one Metric across documents)")

# Category and metric pickers
available_categories = [
    c for c in CAT_ORDER if c in merged["Category"].unique().tolist()
] + [c for c in merged["Category"].unique() if c not in CAT_ORDER]
category_sel = st.selectbox("Category", options=available_categories, index=available_categories.index("1PDP") if "1PDP" in available_categories else 0)

metric_map = {
    "Net Oil": "Oil (Mbbl)",
    "Net Gas": "Gas (MMcf)",
    "Net NGL": "NGL (Mbbl)",
    "Net BOE": "Net BOE (Mboe)",
    "PV9 (M$)": "PV9 (M$)",
    "PV10 (M$)": "PV10 (M$)",
}
metric_label = st.selectbox("Metric", options=list(metric_map.keys()), index=0)
metric_sel = metric_map[metric_label]

# Build the matrix: rows=Source, cols=File
slice_df = build_slice_matrix(
    merged, category=category_sel, metric=metric_sel,
    row_dim="Source", col_dim="File",
    abs_tol_val=abs_tol, rel_tol_val=rel_tol_pct
)

if slice_df.empty:
    st.info("No data for that (Category, Metric) slice yet.")
else:
    st.dataframe(slice_df, use_container_width=True)
    st.download_button(
        "Download slice CSV",
        data=slice_df.to_csv().encode("utf-8"),
        file_name=f"slice_{category_sel}_{metric_label.replace(' ', '_')}.csv",
        mime="text/csv",
    )

# ---- Consistency checks (per file)
st.subheader("Consistency checks (by file)")
def _check_consistency(df: pd.DataFrame) -> pd.DataFrame:
    out = []
    for cat in sorted(df["Category"].dropna().unique()):
        for m in METRICS:
            vals = df.loc[df["Category"] == cat, m].dropna().tolist()
            if strict and len(vals) < df["Source"].nunique():
                ok = False
            else:
                ok = within_tolerance(vals, abs_tol, rel_tol_pct) if vals else False
            out.append({
                "Category": cat, "Metric": m,
                "Sources": int(df.loc[(df["Category"] == cat) & df[m].notna()].shape[0]),
                "Min": pd.Series(vals).min() if vals else np.nan,
                "Max": pd.Series(vals).max() if vals else np.nan,
                "Consistent?": check_mark(ok),
            })
    return pd.DataFrame(out)

cons_rows = (
    merged.groupby(["File"])
    .apply(lambda g: _check_consistency(g))
    .reset_index(level=0)
    .rename(columns={"level_0": "File"})
)
st.dataframe(cons_rows, use_container_width=True)

st.subheader("Overall (per file)")
overall = []
for file, g in cons_rows.groupby("File"):
    ok = g["Consistent?"].eq(GREEN_CHECK).all()
    overall.append({"File": file, "Pass?": check_mark(ok)})
st.dataframe(pd.DataFrame(overall), use_container_width=True)

# ---- Downloads
st.download_button(
    "Download extracted data (CSV)",
    data=merged.to_csv(index=False).encode("utf-8"),
    file_name=f"extracted_all_{(case_name or 'report').replace(' ', '_')}.csv",
    mime="text/csv",
)
