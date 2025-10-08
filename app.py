import re
import math
import pandas as pd
import streamlit as st

# ---------- PDF parsing deps ----------
try:
    import pdfplumber  # type: ignore
    HAS_PDFPLUMBER = True
except Exception:
    HAS_PDFPLUMBER = False

st.set_page_config(page_title="Reserves Tie-Out Checker", layout="wide")
st.title("📊 Reserves Tie‑Out Checker")
st.caption(
    "Cross‑check PDF (Table 1.1 / Cash Flows / Oneline) and Excel (Oneline + Monthly). "
    "Green ✅ / Red ❌ indicate consistency. PV shown as PV9 / PV10 (M$ = thousands of dollars)."
)

# ---------- Sidebar ----------
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
    case_name = st.text_input("Case/Project name (for CSV name)", "")

# ---------- Helpers ----------
NUM_RE = re.compile(r"^-?\d[\d,]*\.?\d*$")
PV_BOX_PAT = re.compile(r"(?i)P\.W\.,\s*M\$")

def _to_f(val):
    """String/number -> float with commas/$ stripped."""
    if val is None or (isinstance(val, float) and math.isnan(val)):
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

def _norm_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _sum_numeric(series: pd.Series) -> float:
    return pd.to_numeric(series, errors="coerce").sum(skipna=True)

def _slug(s: str) -> str:
    """Make a forgiving, comparable version of a column name."""
    s = (s or "").replace("\xa0", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    s = s.replace("@", "at")
    s = s.replace("mmbbls", "mbbl").replace("mbbls", "mbbl")
    s = s.replace("mmcf", "mmcf").replace("mboe", "mboe")
    return s

# ---------- Canonical metric names ----------
METRICS = [
    "Oil (Mbbl)",
    "Gas (MMcf)",
    "NGL (Mbbl)",
    "Net BOE (Mboe)",
    "PV9 (M$)",      # thousands of dollars
    "PV10 (M$)",     # thousands of dollars
]

# ---------- Regex (PDF) ----------
TABLE11_ROW_PAT = re.compile(
    r"(?i)(Total\s+Proved\s+Reserves|Proved\s+Developed\s+Producing\s+\(1PDP\)|Proved\s+Developed\s+Non-?Producing\s+\(3PDNP\)|Proved\s+Undeveloped\s+\(4PUD\)|Total\s+Probable\s+Reserves\s+\(5PROB\)|Total\s+Possible\s+Reserves\s+\(6POSS\)).*?"
    r"([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+\$\s*([0-9,]+)\s+\$\s*([0-9,]+)"
)
RSV_CAT_PAT = re.compile(r"(?i)SE[_\s]*RSV[_\s]*CAT\s*=\s*(1PDP|3PDNP|4PUD|5PROB|6POSS)")
ONELINE_CAT_TOTAL_PAT = re.compile(
    r"(?im)^\s*(1PDP|3PDNP|4PUD|5PROB|6POSS)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,][0-9,]*)\s+([0-9,][0-9,]*)\s*$"
)
ONELINE_GRAND_TOTAL_PAT = re.compile(
    r"(?im)^\s*Grand\s+Total\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,][0-9,]*)\s+([0-9,][0-9,]*)\s*$"
)

# ---------- Cash‑flow helpers ----------
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

    single = [w for w in headers if U(w["text"]) in (f"NET {token2}", f"NET\u00A0{token2}")]
    if single:
        w = min(single, key=lambda t: t["top"])
        return (w["x0"] + w["x1"]) / 2.0

    nets = [w for w in headers if U(w["text"]) == "NET"]
    for n in nets:
        same_line = [w for w in headers if abs(w["top"] - n["top"]) < 2.5 and w["x0"] > n["x0"]]
        t2 = [w for w in same_line if U(w["text"]) == token2]
        if t2:
            w = t2[0]
            x0 = min(n["x0"], w["x0"]); x1 = max(n["x1"], w["x1"])
            return (x0 + x1) / 2.0

    tokens = [w for w in headers if U(w["text"]) == token2]
    if tokens:
        w = min(tokens, key=lambda t: t["top"])
        return (w["x0"] + w["x1"]) / 2.0
    return None

def _extract_cashflow_totals_from_page(page) -> dict:
    """
    Extract from a CF page:
      • NET OIL / NET GAS / NET NGL totals from the *upper* table's TOTAL row
      • PV9 & PV10 (M$) from the P.W., M$ box at bottom
    """
    words = page.extract_words(
        use_text_flow=True, keep_blank_chars=False, extra_attrs=["x0", "x1", "top", "bottom"]
    )
    if not words:
        return {}

    # Assemble lines to find the economics header ("CUM. DISC. FCF")
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
        vol_tot_y = min(totals, key=lambda w: w["top"])["top"]  # fallback

    # Header centers
    x_oil = _find_header_center(words, vol_tot_y, "OIL")
    x_gas = _find_header_center(words, vol_tot_y, "GAS")
    x_ngl = _find_header_center(words, vol_tot_y, "NGL")

    # Numbers on the volumes TOTAL line
    nums_on_total = [
        w for w in words if abs(w["top"] - vol_tot_y) < 2.5 and NUM_RE.match(w["text"].strip())
    ]
    def at_col(xc):
        if xc is None or not nums_on_total:
            return math.nan
        w = _nearest_by_x(xc, nums_on_total)
        return _to_f(w["text"]) if w else math.nan

    oil = at_col(x_oil)
    gas = at_col(x_gas)
    ngl = at_col(x_ngl)

    # PV9 & PV10 from P.W., M$ box
    pv9 = math.nan
    pv10 = math.nan
    t = page.extract_text() or ""
    m = PV_BOX_PAT.search(t)
    if m:
        tail = t[m.start():]
        for rate_str, val_str in re.findall(r"(\d{1,2}\.\d{2})\s+([0-9][\d,]*\.\d+)", tail):
            r = _to_f(rate_str)
            v = _to_f(val_str)
            if abs(r - 9.0) < 0.01:
                pv9 = v
            elif abs(r - 10.0) < 0.01:
                pv10 = v

    return {"Oil (Mbbl)": oil, "Gas (MMcf)": gas, "NGL (Mbbl)": ngl, "PV9 (M$)": pv9, "PV10 (M$)": pv10}

# ---------- PDF parser ----------
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
                gas = _to_f(m.group(2))
                ngl = _to_f(m.group(3))
                oil = _to_f(m.group(4))
                boe = _to_f(m.group(5))
                pv10_m = _to_f(m.group(7))   # already in M$
                key = None
                if "Developed Producing" in label: key = "1PDP"
                elif "Non-Prod" in label or "NonProducing" in label or "3PDNP" in label: key = "3PDNP"
                elif "Undeveloped" in label: key = "4PUD"
                elif "Probable" in label: key = "5PROB"
                elif "Possible" in label: key = "6POSS"
                elif "Total Proved Reserves" in label: key = "TOTAL PROVED"
                if key:
                    rows.append({
                        "Source": "Table1.1",
                        "Category": key,
                        "Oil (Mbbl)": oil,
                        "Gas (MMcf)": gas,
                        "NGL (Mbbl)": ngl,
                        "Net BOE (Mboe)": boe,
                        "PV9 (M$)": math.nan,   # Table 1.1 prints PV10 only
                        "PV10 (M$)": pv10_m,
                    })

            # Cash‑flow pages (volumes totals + PV9/PV10)
            mcat = RSV_CAT_PAT.search(text)
            if mcat:
                cat = mcat.group(1)
                cf = _extract_cashflow_totals_from_page(page)
                rows.append({
                    "Source": "Cash Flows",
                    "Category": cat,
                    "Oil (Mbbl)": cf.get("Oil (Mbbl)", math.nan),
                    "Gas (MMcf)": cf.get("Gas (MMcf)", math.nan),
                    "NGL (Mbbl)": cf.get("NGL (Mbbl)", math.nan),
                    "Net BOE (Mboe)": math.nan,           # not on CF pages
                    "PV9 (M$)": cf.get("PV9 (M$)", math.nan),
                    "PV10 (M$)": cf.get("PV10 (M$)", math.nan),
                })

            # Oneline (PDF) grey category totals + Grand Total (PV10 only)
            for m in ONELINE_CAT_TOTAL_PAT.finditer(text):
                cat = m.group(1)
                oil = _to_f(m.group(2)); gas = _to_f(m.group(3)); ngl = _to_f(m.group(4))
                boe = _to_f(m.group(5)); npv10 = _to_f(m.group(7))
                rows.append({
                    "Source": "Oneline PDF",
                    "Category": cat,
                    "Oil (Mbbl)": oil,
                    "Gas (MMcf)": gas,
                    "NGL (Mbbl)": ngl,
                    "Net BOE (Mboe)": boe,
                    "PV9 (M$)": math.nan,
                    "PV10 (M$)": npv10 / 1_000.0,      # $ -> M$
                })
            mgt = ONELINE_GRAND_TOTAL_PAT.search(text)
            if mgt:
                oil = _to_f(mgt.group(1)); gas = _to_f(mgt.group(2)); ngl = _to_f(mgt.group(3))
                boe = _to_f(mgt.group(4)); npv10 = _to_f(mgt.group(6))
                rows.append({
                    "Source": "Oneline PDF",
                    "Category": "TOTAL PROVED",
                    "Oil (Mbbl)": oil,
                    "Gas (MMcf)": gas,
                    "NGL (Mbbl)": ngl,
                    "Net BOE (Mboe)": boe,
                    "PV9 (M$)": math.nan,
                    "PV10 (M$)": npv10 / 1_000.0,
                })

    if not rows:
        return None, "No recognizable sections found."
    return pd.DataFrame(rows), None

# ---------- Excel: Oneline (robust headers; all sheets) ----------
def _pick_col(df: pd.DataFrame, want: str, regexes: list[str]) -> str | None:
    """Return actual column name in df that best matches `want` slug or regex candidates."""
    slug_map = {_slug(c): c for c in df.columns}
    # exact slug
    if want in slug_map:
        return slug_map[want]
    # regex pass over slugs
    for rx in regexes:
        pat = re.compile(rx, flags=re.I)
        for s, orig in slug_map.items():
            if pat.search(s):
                return orig
    return None

def parse_oneline_xlsx(file):
    """
    Robustly find:
      SE_RSV_CAT, Net Res Oil (Mbbl), Net Res Gas (MMcf), Net Res NGL (Mbbl),
      [optional] Net Res (MBOE), NPV at 9%, NPV at 10%.
    Group & sum by SE_RSV_CAT; convert NPVs ($) -> PV9/PV10 (M$).
    """
    xls = pd.ExcelFile(file)
    frames = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(file, sheet_name=sheet)
        if df is None or df.empty:
            continue
        df = _norm_columns(df)

        # Desired slugs and flexible regexes
        want_cat  = "se rsv cat"
        want_oil  = "net res oil (mbbl)"; want_oil_slug = _slug(want_oil)
        want_gas  = "net res gas (mmcf)"; want_gas_slug = _slug(want_gas)
        want_ngl  = "net res ngl (mbbl)"; want_ngl_slug = _slug(want_ngl)
        want_boe  = "net res (mboe)";     want_boe_slug = _slug(want_boe)
        want_pv9  = "npv at 9%";          want_pv9_slug = _slug(want_pv9)
        want_pv10 = "npv at 10%";         want_pv10_slug = _slug(want_pv10)

        cat_col  = _pick_col(df, want_cat,  [r"\bse\s*[_ ]?\s*rsv\s*[_ ]?\s*cat\b"])
        oil_col  = _pick_col(df, want_oil_slug, [r"net\s*res.*oil.*mbbl"])
        gas_col  = _pick_col(df, want_gas_slug, [r"net\s*res.*gas.*mmcf"])
        ngl_col  = _pick_col(df, want_ngl_slug, [r"net\s*res.*ngl.*mbbl"])
        boe_col  = _pick_col(df, want_boe_slug, [r"net\s*res.*mboe"])
        pv9_col  = _pick_col(df, want_pv9_slug, [r"(npv|pv).*(9\b|9\s*%)"])
        pv10_col = _pick_col(df, want_pv10_slug,[r"(npv|pv).*(10\b|10\s*%)"])

        needed = [cat_col, oil_col, gas_col, ngl_col]
        if not all(needed):
            # Show a compact warning once per sheet that fails
            st.warning(
                f"Oneline XLS: '{sheet}' missing required columns. "
                f"Found: {list(df.columns)}"
            )
            continue

        # Coerce numerics (strip $, commas)
        for c in [oil_col, gas_col, ngl_col, boe_col, pv9_col, pv10_col]:
            if c and c in df:
                df[c] = pd.to_numeric(df[c].replace(r"[\$,]", "", regex=True), errors="coerce")

        # Group & sum by SE_RSV_CAT
        gcols = {oil_col: "Oil (Mbbl)", gas_col: "Gas (MMcf)", ngl_col: "NGL (Mbbl)"}
        if boe_col:
            gcols[boe_col] = "Net BOE (Mboe)"
        agg = df.groupby(df[cat_col].astype(str).str.strip()).agg({c: "sum" for c in gcols.keys()}).reset_index()
        agg.rename(columns={agg.columns[0]: "Category", **gcols}, inplace=True)
        agg["Source"] = "Oneline XLS"
        # PV9/PV10 (M$)
        if pv9_col:
            pv9 = df.groupby(df[cat_col].astype(str).str.strip())[pv9_col].sum().reset_index(drop=True)
            agg["PV9 (M$)"] = pv9 / 1_000.0
        else:
            agg["PV9 (M$)"] = math.nan
        if pv10_col:
            pv10 = df.groupby(df[cat_col].astype(str).str.strip())[pv10_col].sum().reset_index(drop=True)
            agg["PV10 (M$)"] = pv10 / 1_000.0
        else:
            agg["PV10 (M$)"] = math.nan

        frames.append(agg[["Source", "Category"] + METRICS])

    if frames:
        return pd.concat(frames, ignore_index=True)
    return pd.DataFrame(columns=["Source", "Category"] + METRICS)

# ---------- Excel: Monthly (tight headers) ----------
def parse_monthly_xlsx(file):
    """
    Expect: SE_RSV_CAT, Net Oil Prod, Net Gas Prod, Net NGL Prod.
    Sums to category totals. PV not present.
    """
    all_sheets = pd.read_excel(file, sheet_name=None)
    frames = []
    for _, df in (all_sheets or {}).items():
        if df is None or df.empty:
            continue
        df = _norm_columns(df)
        map_lower = {c.lower(): c for c in df.columns}
        def col(name): return map_lower.get(name.lower())

        cat = col("SE_RSV_CAT") or col("se_rsv_cat") or col("SE RSV CAT")
        oil = col("Net Oil Prod")
        gas = col("Net Gas Prod")
        ngl = col("Net NGL Prod")
        boe = col("Net BOE") or col("Net Equiv")

        if not (cat and oil and gas and ngl):
            continue

        for c in [oil, gas, ngl, boe]:
            if c in df:
                df[c] = pd.to_numeric(df[c].replace(r"[\$,]", "", regex=True), errors="coerce")

        slim = df[[c for c in [oil, gas, ngl, boe] if c]].copy()
        grouped = slim.groupby(df[cat].astype(str).str.strip()).agg(_sum_numeric).reset_index()
        grouped.rename(columns={
            grouped.columns[0]: "Category",
            oil: "Oil (Mbbl)", gas: "Gas (MMcf)", ngl: "NGL (Mbbl)", boe: "Net BOE (Mboe)"
        }, inplace=True)
        grouped["Source"] = "Monthly XLS"
        grouped["PV9 (M$)"] = math.nan
        grouped["PV10 (M$)"] = math.nan
        frames.append(grouped[["Source", "Category"] + METRICS])

    if frames:
        return pd.concat(frames, ignore_index=True)
    return pd.DataFrame(columns=["Source", "Category"] + METRICS)

# ---------- Consistency checks ----------
def check_consistency(df: pd.DataFrame) -> pd.DataFrame:
    out = []
    for cat in sorted(df["Category"].unique()):
        for metric in METRICS:
            vals = df.loc[df["Category"] == cat, metric].dropna().tolist()
            if strict and len(vals) < df["Source"].nunique():
                ok = False
            else:
                ok = within_tolerance(vals, abs_tol, rel_tol_pct) if vals else False
            out.append({
                "Category": cat,
                "Metric": metric,
                "Sources": int(df.loc[(df["Category"] == cat) & df[metric].notna()].shape[0]),
                "Min": pd.Series(vals).min() if vals else math.nan,
                "Max": pd.Series(vals).max() if vals else math.nan,
                "Consistent?": "✅" if ok else "❌",
            })
    return pd.DataFrame(out)

# ---------- UI ----------
lcol, rcol = st.columns(2)
with lcol:
    pdf_files = st.file_uploader("Upload PDF report(s)", type=["pdf"], accept_multiple_files=True)
with rcol:
    oneline_xls = st.file_uploader("Upload Oneline XLS/XLSX", type=["xls", "xlsx"], accept_multiple_files=False)
monthly_xls = st.file_uploader("Upload Monthly XLS/XLSX", type=["xls", "xlsx"], key="monthly", accept_multiple_files=False)

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
                st.warning("Oneline XLS: no recognizable columns found.")
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
                st.warning("Monthly XLS: no recognizable columns found.")
        except Exception as e:
            st.error(f"Monthly XLS parse error: {e}")

# ---------- Output ----------
if frames:
    merged = pd.concat(frames, ignore_index=True)

    st.subheader("Extracted figures (all sources)")
    st.dataframe(merged, use_container_width=True)

    st.subheader("Consistency checks (by file)")
    results = (
        merged.groupby(["File"])
        .apply(lambda g: check_consistency(g))
        .reset_index(level=0)
        .rename(columns={"level_0": "File"})
    )
    st.dataframe(results, use_container_width=True)

    st.subheader("Overall")
    overall = []
    for file, group in results.groupby("File"):
        ok = group["Consistent?"].eq("✅").all()
        overall.append({"File": file, "Pass?": "✅" if ok else "❌"})
    st.dataframe(pd.DataFrame(overall), use_container_width=True)

    st.download_button(
        "Download detailed CSV",
        data=merged.to_csv(index=False).encode("utf-8"),
        file_name=f"schaper_tieout_{(case_name or 'report').replace(' ', '_')}.csv",
        mime="text/csv",
    )
else:
    st.info("Upload at least one PDF and/or XLS to begin.")
