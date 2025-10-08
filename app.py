# app.py
import io
import re
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# PDF parsing
import pdfplumber

st.set_page_config(page_title="Reserves Cross-Check", layout="wide")

# ---------- UI helpers ----------
GREEN_CHECK = "✅"
RED_X = "❌"

def check_mark(ok: bool) -> str:
    return GREEN_CHECK if ok else RED_X

def numberize(x):
    """Convert strings like '$ 1,234.56' or '1,234' to float; leave None if not parseable."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
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

def normalize_header(h: str) -> str:
    """
    Lowercase, remove units, punctuation & extra spaces so we can match flexible headers.
    Examples:
      'Net Res Gas (MMcf)' -> 'net res gas'
      'NPV at 10%' -> 'npv at 10'
    """
    s = h or ""
    s = re.sub(r"\(.*?\)", "", s, flags=re.I)  # remove units in ( )
    s = re.sub(r"[%$]", "", s)                 # remove % and $
    s = re.sub(r"[_\-./]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

# Canonical columns we want from XLS workbooks
XLS_PATTERNS: Dict[str, List[str]] = {
    "rsv_cat": [
        r"\bse\s*_?\s*rsv\s*_?\s*cat\b",
        r"\breserve[s]?\s*category\b",
        r"\bcategory\b",
    ],
    "oil": [
        r"\bnet\s*res\s*oil\b",
        r"\bnet\s*oil\b",
        r"\bnet\s*oil\s*prod(uction)?\b",
    ],
    "gas": [
        r"\bnet\s*res\s*gas\b",
        r"\bnet\s*gas\b",
        r"\bnet\s*gas\s*prod(uction)?\b",
    ],
    "ngl": [
        r"\bnet\s*res\s*ngl\b",
        r"\bnet\s*ngl\b",
        r"\bnet\s*ngl\s*prod(uction)?\b",
    ],
    "boe": [
        r"\bnet\s*res\s*(mboe|boe)\b",
        r"\bnet\s*boe\b",
    ],
    "npv9": [
        r"\bnpv\s*at\s*9\b",
        r"\bnpv9\b",
        r"\bpv9\b",
    ],
    "npv10": [
        r"\bnpv\s*at\s*10\b",
        r"\bnpv10\b",
        r"\bpv10\b",
    ],
    "lease": [
        r"\blease\b",
        r"\bwell\b",
        r"\bname\b",
    ],
}

def map_columns(df: pd.DataFrame) -> Dict[str, str]:
    """Return a map {canonical -> actual_column_name_in_df} using fuzzy patterns."""
    norm_cols = {c: normalize_header(str(c)) for c in df.columns}
    mapping = {}
    for want, patterns in XLS_PATTERNS.items():
        for c, norm in norm_cols.items():
            if any(re.search(p, norm) for p in patterns):
                mapping[want] = c
                break
    return mapping

def read_any_excel(uploaded_file) -> pd.DataFrame:
    """
    Read all sheets from an Excel file into a single DataFrame (stacked),
    preserving headers and trying both default and openpyxl engines.
    """
    content = uploaded_file.read()
    bio = io.BytesIO(content)

    tried = []
    for engine in [None, "openpyxl"]:
        try:
            dfs = pd.read_excel(bio, sheet_name=None, engine=engine)
            frames = []
            for name, sheet in dfs.items():
                # Drop fully-empty columns/rows
                sheet = sheet.copy()
                sheet = sheet.dropna(how="all")
                sheet = sheet.loc[:, ~sheet.columns.to_series().astype(str).str.fullmatch(r"\s*nan\s*", case=False, na=False)]
                if not sheet.empty:
                    frames.append(sheet)
            if not frames:
                return pd.DataFrame()
            out = pd.concat(frames, ignore_index=True)
            return out
        except Exception as e:
            tried.append(str(e))
            bio.seek(0)
            continue

    # If we get here, all engines failed.
    raise RuntimeError("Could not read Excel; tried engines -> " + " | ".join(tried))

def aggregate_xls_by_category(df: pd.DataFrame, file_label: str) -> Tuple[pd.DataFrame, List[str]]:
    """
    Map columns, clean numerics, group by SE_RSV_CAT, sum desired metrics.
    Returns (agg_df, warnings)
    """
    warnings = []
    if df.empty:
        return pd.DataFrame(), [f"{file_label}: no data"]

    colmap = map_columns(df)
    required_any = ["rsv_cat", "oil", "gas", "ngl", "npv10"]  # npv9/boe optional
    missing = [k for k in required_any if k not in colmap]
    if missing:
        warnings.append(f"{file_label}: missing columns {missing} (header mapping is case/spacing tolerant).")

    # Build a working frame with what we have
    use_cols = {k: colmap[k] for k in colmap.keys()}
    work = df.rename(columns={v: k for k, v in use_cols.items()})

    # Clean numbers
    for c in ["oil", "gas", "ngl", "boe", "npv9", "npv10"]:
        if c in work:
            work[c] = work[c].map(numberize)

    # Normalize category labels (strip/upper/no spaces)
    if "rsv_cat" in work:
        work["rsv_cat"] = work["rsv_cat"].astype(str).str.strip().str.upper()

    # Keep just needed columns
    keep = ["rsv_cat", "oil", "gas", "ngl", "boe", "npv9", "npv10"]
    work = work[[c for c in keep if c in work]]

    # Group & sum
    agg = work.groupby("rsv_cat", dropna=True).sum(numeric_only=True).reset_index()
    # Standardize category display (1PDP, 3PDNP, 4PUD, 5PROB, 6POSS, TOTAL PROVED, etc.)
    agg = agg.rename(columns={"rsv_cat": "Category"})
    return agg, warnings

# ---------- PDF extraction ----------
TABLE11_CATEGORY_ROW = re.compile(
    r"^(Total Proved Reserves|Proved Developed Producing\s*\(1PDP\)|Proved Developed Non-Producing\s*\(3PDNP\)|Proved Undeveloped\s*\(4PUD\))",
    re.I
)

def extract_table11(pdf: pdfplumber.PDF) -> pd.DataFrame:
    """
    Extract Table 1.1 rows with Gas (MMcf), NGL (Mbbls), Oil (Mbbls), Mboe, Total Undisc $M, PV10 $M.
    Strategy: find page containing 'Table 1.1' and parse the text lines around it.
    """
    records = []
    for page in pdf.pages:
        txt = page.extract_text() or ""
        if "Table 1.1" in txt and "Summary of Reserves" in txt:
            # Parse line-by-line
            for line in txt.splitlines():
                if not TABLE11_CATEGORY_ROW.search(line.strip()):
                    continue
                # Collapse runs of spaces to single
                s = re.sub(r"\s+", " ", line.strip())
                # Extract columns by greedy tail numbers
                # Pattern: <Label> <gas> <ngl> <oil> <mboe> <undisc> <pv10>
                m = re.search(
                    r"^(.*?)([-\d,]+)\s+([-\d,]+)\s+([-\d,]+)\s+([-\d,]+)\s+\$?\s*([-\d,]+)\s+\$?\s*([-\d,]+)\s*$",
                    s
                )
                if m:
                    label = m.group(1).strip()
                    gas = numberize(m.group(2))
                    ngl = numberize(m.group(3))
                    oil = numberize(m.group(4))
                    mboe = numberize(m.group(5))
                    undisc = numberize(m.group(6))
                    pv10 = numberize(m.group(7))
                    # Normalize label to category token
                    category = (
                        "1PDP" if "1PDP" in label.upper()
                        else "3PDNP" if "3PDNP" in label.upper()
                        else "4PUD" if "4PUD" in label.upper()
                        else "TOTAL PROVED"
                    )
                    records.append(
                        {
                            "Source": "Table1.1",
                            "Category": category,
                            "Oil (Mbbl)": oil,
                            "Gas (MMcf)": gas,
                            "NGL (Mbbl)": ngl,
                            "Net BOE (Mboe)": mboe,
                            "PV10 ($M)": pv10,
                            "PV9 ($M)": np.nan,  # Table 1.1 usually shows PV10 only
                        }
                    )
    return pd.DataFrame(records)

def extract_cashflow_totals(pdf: pdfplumber.PDF) -> pd.DataFrame:
    """
    For each Cash Flow section (page header 'SE_RSV_CAT = X'), read the TOTAL line:
      - Net Oil Prod, Net Gas Prod, Net NGL Prod
      - 'TOTAL' block with 'CUM. DISC. FCF.' (PV10) and sometimes '% P.W., M$' 9% (PV9)
    We rely on parsed text and robust regex.
    """
    rows = []
    for page in pdf.pages:
        txt = page.extract_text() or ""
        head = re.search(r"SE_RSV_CAT\s*=\s*([A-Z0-9]+)", txt)
        if not head:
            continue
        cat = head.group(1).strip().upper()

        # Net production totals (look for 'TOTAL' section line containing three numbers just before prices)
        # We’ll capture the line that begins with TOTAL and has many numeric columns,
        # then specifically pick the three numbers right after the headers 'NET OIL PROD', 'NET GAS PROD', 'NET NGL PROD' totals.
        # A simpler robust approach: look for the line 'TOTAL' in the production block and parse the three small totals
        prod_total = None
        for line in txt.splitlines():
            if re.search(r"^\s*TOTAL\s", line) and re.search(r"NET\s+OIL", txt):
                # Scan the last occurrence of a 'TOTAL' in the production half. The first TOTAL in production block usually contains 3 small numbers.
                prod_total = line

        oil_m = gas_m = ngl_m = np.nan
        if prod_total:
            # Extract triplet like: 'TOTAL .... 55.6 84.0 15.0 ...'
            trip = re.findall(r"\s([0-9]{1,3}(?:\.[0-9]+)?)\s([0-9]{1,4}(?:\.[0-9]+)?)\s([0-9]{1,3}(?:\.[0-9]+)?)\s", prod_total)
            if trip:
                oil_m, gas_m, ngl_m = [numberize(v) for v in trip[-1]]

        # PV10 from 'CUM. DISC. FCF.' at the very end TOTAL line
        pv10 = np.nan
        pv9 = np.nan
        tail_total = None
        for line in txt.splitlines():
            if re.search(r"^\s*TOTAL\s", line) and re.search(r"CUM\.\s*DISC\.\s*FCF", txt):
                tail_total = line
        if tail_total:
            # The last numeric on this line is PV10 (million $)
            nums = [numberize(n) for n in re.findall(r"[-]?\d{1,3}(?:,\d{3})*(?:\.\d+)?", tail_total)]
            if nums:
                pv10 = nums[-1]

        rows.append(
            {
                "Source": "Cash Flows",
                "Category": cat,
                "Oil (Mbbl)": oil_m,  # these totals are in Mbbl/MMcf/Mbbl (already net)
                "Gas (MMcf)": gas_m,
                "NGL (Mbbl)": ngl_m,
                "Net BOE (Mboe)": np.nan,
                "PV10 ($M)": pv10,
                "PV9 ($M)": pv9,
            }
        )
    return pd.DataFrame(rows)

def extract_oneline_pdf_totals(pdf: pdfplumber.PDF) -> pd.DataFrame:
    """
    From the Oneline Summary section (gray sub-totals), capture the bold subtotal
    rows per category (e.g., '1PDP 56 84 15 85 ... 2,199,233').
    """
    rows = []
    for page in pdf.pages:
        txt = page.extract_text() or ""
        if "Oneline Summary" not in txt and "LEASE" not in txt:
            continue
        for line in txt.splitlines():
            s = re.sub(r"\s+", " ", line.strip())
            # Match '1PDP 56 84 15 85 ... 2,199,233'
            m = re.match(
                r"^(1PDP|3PDNP|4PUD|5PROB|6POSS)\s+([-\d,]+)\s+([-\d,]+)\s+([-\d,]+)\s+([-\d,]+).*?([-\d,]+)\s*$",
                s
            )
            if m:
                cat = m.group(1).strip().upper()
                oil = numberize(m.group(2))
                gas = numberize(m.group(3))
                ngl = numberize(m.group(4))
                mboe = numberize(m.group(5))
                pv10 = numberize(m.group(6))
                rows.append(
                    {
                        "Source": "Oneline PDF",
                        "Category": cat,
                        "Oil (Mbbl)": oil,
                        "Gas (MMcf)": gas,
                        "NGL (Mbbl)": ngl,
                        "Net BOE (Mboe)": mboe,
                        "PV10 ($M)": pv10,
                        "PV9 ($M)": np.nan,
                    }
                )
    return pd.DataFrame(rows)

def combine_pdf_sources(pdf_bytes: bytes) -> pd.DataFrame:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        t11 = extract_table11(pdf)
        cf = extract_cashflow_totals(pdf)
        ol = extract_oneline_pdf_totals(pdf)
    all_df = pd.concat([t11, cf, ol], ignore_index=True)
    # fill category names more uniformly
    all_df["Category"] = all_df["Category"].replace(
        {
            "TOTAL PROVED": "TOTAL PROVED",
            "1PDP": "1PDP",
            "3PDNP": "3PDNP",
            "4PUD": "4PUD",
        }
    )
    return all_df

def compare_frames(pdf_df: pd.DataFrame, xls_df: pd.DataFrame, label: str, tol=0.51) -> pd.DataFrame:
    """
    Merge PDF vs XLS by Category and show deltas and a 'Consistent?' flag
    (tolerance = absolute difference <= tol).
    """
    want_cols = ["Oil (Mbbl)", "Gas (MMcf)", "NGL (Mbbl)", "PV9 ($M)", "PV10 ($M)"]
    pdf_slim = (
        pdf_df.groupby("Category", dropna=False)[want_cols]
        .sum(numeric_only=True)
        .reset_index()
    )

    # XLS may not have PV9; fill if missing
    for c in want_cols:
        if c not in xls_df.columns:
            xls_df[c] = np.nan
    xls_slim = xls_df.groupby("Category", dropna=False)[want_cols].sum(numeric_only=True).reset_index()

    merged = pd.merge(pdf_slim, xls_slim, on="Category", how="outer", suffixes=(" PDF", f" {label}"))
    for c in ["Oil (Mbbl)", "Gas (MMcf)", "NGL (Mbbl)", "PV9 ($M)", "PV10 ($M)"]:
        merged[f"{c} Δ"] = (merged[f"{c} PDF"] - merged[f"{c} {label}"]).abs()

    merged["Consistent?"] = merged.apply(
        lambda r: all(
            (np.isnan(r[f"{c} Δ"]) or r[f"{c} Δ"] <= tol)
            for c in ["Oil (Mbbl)", "Gas (MMcf)", "NGL (Mbbl)", "PV10 ($M)"]
        ),
        axis=1,
    )
    return merged

def pivot_by_category(df: pd.DataFrame, value_cols: List[str]) -> pd.DataFrame:
    g = df.groupby("Category")[value_cols].sum(numeric_only=True)
    # Categories as columns, metrics as rows
    return g.T

# ---------- App ----------
st.title("Reserves Reconciliation")

pdf_file = st.file_uploader("Upload PDF (Reserves Cover Letter & Cash Flows)", type=["pdf"])
col1, col2 = st.columns(2)
with col1:
    oneline_xls = st.file_uploader("Upload Oneline Report (XLS/XLSX)", type=["xls", "xlsx"], key="one")
with col2:
    monthly_xls = st.file_uploader("Upload Monthly Report (XLS/XLSX)", type=["xls", "xlsx"], key="mon")

pdf_df = pd.DataFrame()
if pdf_file is not None:
    pdf_bytes = pdf_file.read()
    with st.spinner("Parsing PDF…"):
        pdf_df = combine_pdf_sources(pdf_bytes)

    st.subheader("Extracted figures (all PDF sources)")
    st.dataframe(pdf_df, use_container_width=True, height=420)

    # Flipped (categories as columns)
    st.subheader("PDF — categories as columns")
    st.dataframe(
        pivot_by_category(pdf_df, ["Oil (Mbbl)", "Gas (MMcf)", "NGL (Mbbl)", "Net BOE (Mboe)", "PV9 ($M)", "PV10 ($M)"]),
        use_container_width=True,
    )

# ----- XLS parsing & aggregation -----
xls_agg_frames = []
one_warns: List[str] = []
mon_warns: List[str] = []

if oneline_xls is not None:
    try:
        df_one = read_any_excel(oneline_xls)
        one_agg, one_warns = aggregate_xls_by_category(df_one, "Oneline")
        if not one_agg.empty:
            one_agg["Source"] = "Oneline XLS"
            xls_agg_frames.append(one_agg)
    except Exception as e:
        one_warns.append(f"Oneline XLS parse error: {e}")

if monthly_xls is not None:
    try:
        df_mon = read_any_excel(monthly_xls)
        mon_agg, mon_warns = aggregate_xls_by_category(df_mon, "Monthly")
        if not mon_agg.empty:
            mon_agg["Source"] = "Monthly XLS"
            xls_agg_frames.append(mon_agg)
    except Exception as e:
        mon_warns.append(f"Monthly XLS parse error: {e}")

if one_warns:
    for w in one_warns:
        st.error(w)
if mon_warns:
    for w in mon_warns:
        st.error(w)

if xls_agg_frames:
    xls_all = pd.concat(xls_agg_frames, ignore_index=True)
    # Show raw XLS sums
    st.subheader("XLS — grouped & summed by SE_RSV_CAT")
    st.dataframe(xls_all, use_container_width=True)

    # Flipped (categories as columns)
    st.subheader("XLS — categories as columns")
    st.dataframe(
        pivot_by_category(
            xls_all,
            ["Oil (Mbbl)", "Gas (MMcf)", "NGL (Mbbl)", "boe", "NPV9 ($M)" if "NPV9 ($M)" in xls_all.columns else "npv9", "NPV10 ($M)" if "NPV10 ($M)" in xls_all.columns else "npv10"],
        ),
        use_container_width=True,
    )

    # If PDF present, compare
    if not pdf_df.empty:
        # Build one combined XLS roll-up (Oneline + Monthly) by Category
        combined_xls = (
            xls_all.groupby("Category")[["Oil (Mbbl)", "Gas (MMcf)", "NGL (Mbbl)", "npv9", "npv10"]]
            .sum(numeric_only=True)
            .reset_index()
            .rename(columns={"npv9": "PV9 ($M)", "npv10": "PV10 ($M)"})
        )

        st.subheader("Reconciliation — PDF vs XLS (summed by SE_RSV_CAT)")
        recon = compare_frames(pdf_df, combined_xls, label="XLS", tol=0.51)
        # Style Consistent? col with green checks
        recon_display = recon.copy()
        recon_display["Consistent?"] = recon_display["Consistent?"].map(check_mark)
        st.dataframe(recon_display, use_container_width=True)

else:
    if oneline_xls is None and monthly_xls is None:
        st.info("Upload your Oneline and/or Monthly reports to check them against the PDF.")
