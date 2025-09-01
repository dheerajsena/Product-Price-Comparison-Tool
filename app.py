# app.py
import re
from difflib import SequenceMatcher
from io import BytesIO

import pandas as pd
import streamlit as st

# --------------------------
# Page configuration & Header
# --------------------------
st.set_page_config(page_title="Product Price Comparison", layout="wide", page_icon="🚀")

st.markdown(
    """
    <div style='background-color:#002e5b;padding:12px 16px;border-radius:12px;'>
      <h1 style='text-align:center;color:#fff;margin:0;'>🚀 Product Price Comparison Dashboard</h1>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='text-align:center;font-size:16px;color:#333;margin-top:8px;'>"
    "Upload a Marlin price file and a Website price file — or download our templates below — to generate a detailed comparison report."
    "</p>",
    unsafe_allow_html=True,
)

# ---------------------------------
# Helpers: Template + Column Finding
# ---------------------------------
def make_template_bytes(template_name: str) -> bytes:
    """Create a 2-column Excel file: Variant Code, Variant Price."""
    df = pd.DataFrame({"Variant Code": [], "Variant Price": []})
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Prices", index=False)
        # Add a small 'README' sheet with directions
        pd.DataFrame(
            {
                "Instructions": [
                    "Fill only these two columns.",
                    "Variant Code should be the unique SKU/variant identifier.",
                    "Variant Price should be a number (Inc GST preferred).",
                ]
            }
        ).to_excel(writer, sheet_name="README", index=False)
    return bio.getvalue()

# Candidate text for variant code columns (normalized)
CODE_CANDIDATES = [
    "variantcode",
    "variant_code",
    "variant sku",
    "variantsku",
    "sku",
    "productcode",
    "product_code",
    "itemcode",
    "item_code",
    "code",
    "partnumber",
    "partno",
]

# Candidate text for price columns (normalized)
PRICE_CANDIDATES_PRIMARY = [
    # higher-priority (inc GST / web-facing)
    "variantprice",
    "price",
    "webprice",
    "websiteprice",
    "retail",
    "rrp",
    "sellprice",
    "sellingprice",
    "listprice",
]
# Hints to break ties toward INC over EXC GST when both exist
INC_HINTS = ["incgst", "inclgst", "incl_gst", "inc_gst"]
EXC_HINTS = ["exgst", "exc_gst", "exclgst", "excl_gst"]

def norm(s: str) -> str:
    s = s.strip().lower()
    s = re.sub(r"[^a-z0-9]+", "", s)  # keep alnum only
    return s

def best_match_column(columns, candidates, extra_bias_inc=None, extra_bias_exc=None):
    """
    Pick the best matching column from a list of names using:
    1) direct containment against candidate tokens
    2) simple fuzzy score
    Bias toward INC GST columns when both INC/EXC are present.
    """
    if not columns:
        return None

    normalized = {c: norm(c) for c in columns}

    # Direct containment scoring
    scores = {c: 0.0 for c in columns}
    for col, ncol in normalized.items():
        # base: candidate containment
        for cand in candidates:
            if cand in ncol:
                scores[col] += 1.0

        # fuzzy fallback vs each candidate
        fuzz = max(SequenceMatcher(None, ncol, cand).ratio() for cand in candidates)
        scores[col] += 0.4 * fuzz  # small fuzzy contribution

        # bias toward INC or EXC if present
        if extra_bias_inc and any(h in ncol for h in extra_bias_inc):
            scores[col] += 0.25
        if extra_bias_exc and any(h in ncol for h in extra_bias_exc):
            scores[col] += 0.10

    # pick highest scoring
    best = max(scores, key=lambda c: scores[c])
    # Require a minimal score to avoid nonsense picks
    return best if scores[best] >= 0.6 else None

def pick_sheet_and_columns(xls_file):
    """
    Read ALL sheets, find the sheet that most confidently contains one
    code column and one price column. Return (df_subset, meta).
    """
    # Read everything
    book = pd.read_excel(xls_file, sheet_name=None, dtype=str)
    best = None
    best_meta = None
    best_score = -1

    for sheet_name, df in book.items():
        if df is None or df.empty:
            continue

        # drop empty columns that pandas might create
        df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", na=False)]

        code_col = best_match_column(
            list(df.columns), CODE_CANDIDATES
        )
        price_col = best_match_column(
            list(df.columns),
            PRICE_CANDIDATES_PRIMARY,
            extra_bias_inc=INC_HINTS,
            extra_bias_exc=EXC_HINTS,
        )

        # score: 1 for each found + light reward if both present
        score = (1 if code_col else 0) + (1 if price_col else 0)
        if code_col and price_col:
            score += 0.5

        if score > best_score:
            best = (df, code_col, price_col, sheet_name)
            best_score = score

    if not best or best_score <= 0:
        raise ValueError("Could not detect suitable columns in any sheet.")

    df, code_col, price_col, sheet_name = best
    if not code_col or not price_col:
        raise ValueError(
            f"Auto-detection incomplete in sheet '{sheet_name}'. "
            f"Found code column: {code_col}, price column: {price_col}"
        )

    # keep only the two columns for safety
    sub = df[[code_col, price_col]].copy()

    # rename to standard labels for downstream logic
    sub.rename(columns={code_col: "Variant Code", price_col: "Variant Price"}, inplace=True)
    return sub, {"sheet": sheet_name, "code_col": code_col, "price_col": price_col}

def coerce_price(s) -> float:
    if pd.isna(s):
        return pd.NA
    if isinstance(s, (int, float)):
        return float(s)
    # strip currency and commas, handle parentheses for negatives
    txt = str(s)
    neg = False
    if "(" in txt and ")" in txt:
        neg = True
    txt = re.sub(r"[^\d\.\-]", "", txt)  # keep digits, dot, minus
    if txt.count(".") > 1:  # weird strings like "1.234.56"
        # last dot as decimal separator
        left, right = txt.rsplit(".", 1)
        txt = re.sub(r"\.", "", left) + "." + right
    try:
        val = float(txt)
        return -val if neg else val
    except Exception:
        return pd.NA

def clean_prices(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Variant Code"] = out["Variant Code"].astype(str).str.strip()
    out["Variant Price"] = out["Variant Price"].apply(coerce_price)
    # de-duplicate by latest occurrence
    out = out.dropna(subset=["Variant Code"]).drop_duplicates(subset=["Variant Code"], keep="last")
    return out

def make_report(marlin_df: pd.DataFrame, website_df: pd.DataFrame, meta_m, meta_w, tolerance=0.01) -> bytes:
    # Outer merge on code
    merged = website_df.merge(
        marlin_df,
        on="Variant Code",
        how="outer",
        suffixes=("_Website", "_Marlin"),
    )

    # Price difference (Website - Marlin)
    merged["Price Difference"] = merged["Variant Price_Website"] - merged["Variant Price_Marlin"]

    # Price match with tolerance (cents level default)
    def price_match(row):
        a = row["Variant Price_Website"]
        b = row["Variant Price_Marlin"]
        if pd.isna(a) or pd.isna(b):
            return "N/A"
        return "Match" if abs(a - b) <= tolerance else "Mismatch"

    merged["Price Match"] = merged.apply(price_match, axis=1)

    def compare_label(row):
        a = row["Variant Price_Website"]
        b = row["Variant Price_Marlin"]
        if pd.isna(a) and not pd.isna(b):
            return "Only in Marlin"
        if pd.isna(b) and not pd.isna(a):
            return "Only in Website"
        if pd.isna(a) and pd.isna(b):
            return "Missing in both"
        if abs(a - b) <= tolerance:
            return "Equal (within tolerance)"
        return "Website higher" if (a - b) > 0 else "Marlin higher"

    merged["Comparison"] = merged.apply(compare_label, axis=1)

    # Build Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Order columns nicely
        col_order = [
            "Variant Code",
            "Variant Price_Website",
            "Variant Price_Marlin",
            "Price Difference",
            "Price Match",
            "Comparison",
        ]
        merged[col_order].to_excel(writer, sheet_name="Full Data", index=False)
        merged[merged["Price Match"] == "Match"][col_order].to_excel(writer, sheet_name="Matched", index=False)
        merged[merged["Price Match"] == "Mismatch"][col_order].to_excel(writer, sheet_name="Mismatched", index=False)
        merged[merged["Comparison"] == "Only in Website"][col_order].to_excel(writer, sheet_name="Only in Website", index=False)
        merged[merged["Comparison"] == "Only in Marlin"][col_order].to_excel(writer, sheet_name="Only in Marlin", index=False)

        # Summary sheet
        summary = pd.DataFrame(
            [
                ["Detected Marlin sheet", meta_m["sheet"]],
                ["Marlin code column", meta_m["code_col"]],
                ["Marlin price column", meta_m["price_col"]],
                ["Detected Website sheet", meta_w["sheet"]],
                ["Website code column", meta_w["code_col"]],
                ["Website price column", meta_w["price_col"]],
                ["Total Website rows", len(website_df)],
                ["Total Marlin rows", len(marlin_df)],
                ["Matches", (merged["Price Match"] == "Match").sum()],
                ["Mismatches", (merged["Price Match"] == "Mismatch").sum()],
                ["Only in Website", (merged["Comparison"] == "Only in Website").sum()],
                ["Only in Marlin", (merged["Comparison"] == "Only in Marlin").sum()],
            ],
            columns=["Metric", "Value"],
        )
        summary.to_excel(writer, sheet_name="Summary", index=False)

    return output.getvalue()

# ----------------
# Template section
# ----------------
with st.expander("📥 Download blank templates (fill & re-upload)"):
    tcol1, tcol2 = st.columns(2)
    with tcol1:
        st.download_button(
            "Download Marlin Template",
            data=make_template_bytes("Marlin"),
            file_name="Marlin_Price_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with tcol2:
        st.download_button(
            "Download Website Template",
            data=make_template_bytes("Website"),
            file_name="Website_Price_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ---------------------
# File uploads + action
# ---------------------
col1, col2 = st.columns(2)
with col1:
    marlin_file = st.file_uploader("Upload Marlin Price File (.xlsx)", type=["xlsx"], key="marlin")
with col2:
    website_file = st.file_uploader("Upload Website Price File (.xlsx)", type=["xlsx"], key="website")

run = st.button("Run Comparison", use_container_width=True)

if run:
    if marlin_file is None or website_file is None:
        st.error("Please upload both Excel files to proceed.")
    else:
        try:
            with st.spinner("Auto-detecting columns and comparing…"):
                # Auto-detect (sheet + columns), rename to standard, clean
                m_raw, meta_m = pick_sheet_and_columns(marlin_file)
                w_raw, meta_w = pick_sheet_and_columns(website_file)

                m = clean_prices(m_raw)
                w = clean_prices(w_raw)

                report_bytes = make_report(m, w, meta_m, meta_w, tolerance=0.01)

            st.success("Report ready!")
            st.download_button(
                label="Download Comparison Report",
                data=report_bytes,
                file_name="Price_Comparison_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            # Show a quick preview + what we auto-detected
            with st.expander("🔎 Auto-detection details"):
                lcol, rcol = st.columns(2)
                with lcol:
                    st.markdown("**Marlin detection**")
                    st.write(meta_m)
                    st.dataframe(m.head(10), use_container_width=True)
                with rcol:
                    st.markdown("**Website detection**")
                    st.write(meta_w)
                    st.dataframe(w.head(10), use_container_width=True)

        except Exception as e:
            st.error(f"Comparison failed: {e}")
