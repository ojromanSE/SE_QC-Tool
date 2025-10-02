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
        h = min(matches, key=lambda w: w["top"])  # pick the topmost
        return (h["x0"] + h["x1"]) / 2.0

    x_oil = find_header_x("NET OIL PROD")
    x_gas = find_header_x("NET GAS PROD")
    x_ngl = find_header_x("NET NGL PROD")
    x_boe = find_header_x("NET EQUIV")  # optional

    # Numeric words on TOTAL row
    numeric_on_total = [
        w for w in words
        if abs(w["top"] - tot_y) < 3 and re.match(r"^-?\d[\d,]*\.?\d*$", w["text"].strip())
    ]

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
        out["Net BOE (Mboe)"] = nearest_val(x_boe)
    return out
