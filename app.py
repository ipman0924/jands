from __future__ import annotations

import io
import json
import os
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

# ─────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DD_CACHE_CSV  = os.path.join(BASE_DIR, "dd_feed_cache.csv")
DD_CACHE_META = os.path.join(BASE_DIR, "dd_feed_cache_meta.json")
DD_FEED_URL = (
    "https://portal.dickerdata.com.au/Download?file=raqJl1tyq4ZEzYDHtVUsOuSl4zAncwx"
    "%2Fzg%2Bx0ubLlU%2FenFx3QTFg2jNviQgIJcThjqyfwPHYyhsoCvjDX8xtrx1WQ2i3huc%2B%2F"
    "vGHrRkVJXCT9rf08l5fbXhHHBXMAayhGWNlC62UBLZwBnuPEASD%2BdB18joopRax&fileName="
    "DickerDataDataFeedCSV.csv&displaySaveAs=True"
)
CACHE_TTL_HOURS = 24


# ─────────────────────────────────────────────────────────────
# DD Feed helpers
# ─────────────────────────────────────────────────────────────

def _dd_cache_age_hours() -> float | None:
    if not os.path.exists(DD_CACHE_META):
        return None
    with open(DD_CACHE_META) as f:
        meta = json.load(f)
    ts = datetime.fromisoformat(meta["downloaded_at"])
    return (datetime.utcnow() - ts).total_seconds() / 3600


def _dd_cache_date_str() -> str:
    if not os.path.exists(DD_CACHE_META):
        return "No cache"
    with open(DD_CACHE_META) as f:
        meta = json.load(f)
    return datetime.fromisoformat(meta["downloaded_at"]).strftime("%d %b %Y %H:%M UTC")


def _download_dd_feed() -> tuple[pd.DataFrame, str]:
    resp = requests.get(DD_FEED_URL, timeout=120)
    resp.raise_for_status()
    df = pd.read_csv(io.BytesIO(resp.content), low_memory=False)
    keep = [c for c in ["VendorStockCode", "RRPEx", "DealerEx"] if c in df.columns]
    df = df[keep].copy()
    df.to_csv(DD_CACHE_CSV, index=False)
    now = datetime.utcnow().isoformat()
    with open(DD_CACHE_META, "w") as f:
        json.dump({"downloaded_at": now}, f)
    return df, now


def load_dd_feed(force: bool = False) -> tuple[pd.DataFrame | None, str]:
    age = _dd_cache_age_hours()
    if not force and age is not None and age < CACHE_TTL_HOURS:
        df = pd.read_csv(DD_CACHE_CSV, low_memory=False)
        return df, f"Cached — {_dd_cache_date_str()}"
    try:
        df, ts = _download_dd_feed()
        label = datetime.fromisoformat(ts).strftime("%d %b %Y %H:%M UTC")
        return df, f"Downloaded — {label}"
    except Exception as e:
        if os.path.exists(DD_CACHE_CSV):
            df = pd.read_csv(DD_CACHE_CSV, low_memory=False)
            return df, f"Download failed ({e}); using stale cache — {_dd_cache_date_str()}"
        return None, f"Download failed and no cache: {e}"


# ─────────────────────────────────────────────────────────────
# SAM Report parsing
# ─────────────────────────────────────────────────────────────

@st.cache_data(show_spinner="Parsing SAM report…")
def parse_sam_report(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", header=2)
    df.columns = df.columns.str.strip()
    needed = ["No.", "Vendor Item No.", "Unit price Ex GST", "Trade Price Ex", "Trade Line Discount"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"SAM report missing columns: {missing}")
    df = df[needed].copy()
    df["No."] = df["No."].astype(str).str.strip()
    df["Vendor Item No."] = df["Vendor Item No."].astype(str).str.strip()
    df["Unit price Ex GST"]    = pd.to_numeric(df["Unit price Ex GST"],    errors="coerce")
    df["Trade Price Ex"]       = pd.to_numeric(df["Trade Price Ex"],       errors="coerce")
    df["Trade Line Discount"]  = pd.to_numeric(df["Trade Line Discount"],  errors="coerce")
    df = df.dropna(subset=["No.", "Vendor Item No."])
    df = df[df["No."].str.strip() != ""]
    return df


# ─────────────────────────────────────────────────────────────
# Quote parsing — Excel only
# ─────────────────────────────────────────────────────────────

@st.cache_data(show_spinner="Parsing Excel quote…")
def parse_quote_excel(file_bytes: bytes) -> pd.DataFrame:
    xl = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    sheet = "Quote Product View" if "Quote Product View" in xl.sheet_names else xl.sheet_names[0]
    df = xl.parse(sheet)
    df.columns = df.columns.str.strip()
    col_map = {
        "Quote Product Code":        "jands_sku",
        "Product Description":       "description",
        "List Price Per Unit (Ex GST)": "list_price",
        "Final Discount %":          "quoted_discount_pct",
        "Final Unit Price":          "quoted_price",
        "Quantity":                  "qty",
    }
    for src, dst in col_map.items():
        if src in df.columns:
            df = df.rename(columns={src: dst})
    for col in ["list_price", "quoted_discount_pct", "quoted_price", "qty"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df["jands_sku"] = df["jands_sku"].astype(str).str.strip()
    df = df[~df["jands_sku"].str.lower().isin(["", "nan", "none", "nat"])]
    return df


# ─────────────────────────────────────────────────────────────
# Shared: cross-derive trade price / discount within one source
# ─────────────────────────────────────────────────────────────

def _derive_price_disc(df: pd.DataFrame, price_col: str, disc_col: str, list_col: str) -> pd.DataFrame:
    lp = pd.to_numeric(df[list_col],  errors="coerce")
    tp = pd.to_numeric(df[price_col], errors="coerce")
    dc = pd.to_numeric(df[disc_col],  errors="coerce")

    mask_price = tp.isna() & lp.notna() & dc.notna()
    df.loc[mask_price, price_col] = lp[mask_price] * (1 - dc[mask_price] / 100)

    lp = pd.to_numeric(df[list_col],  errors="coerce")
    tp = pd.to_numeric(df[price_col], errors="coerce")
    mask_disc = dc.isna() & lp.notna() & tp.notna() & (lp != 0)
    df.loc[mask_disc, disc_col] = (1 - tp[mask_disc] / lp[mask_disc]) * 100
    return df


# ─────────────────────────────────────────────────────────────
# Matching + calculations  (quote mode)
# ─────────────────────────────────────────────────────────────

def build_comparison(
    quote_df: pd.DataFrame,
    sam_df: pd.DataFrame,
    dd_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    sam_lookup = sam_df.rename(columns={
        "No.":              "jands_sku",
        "Vendor Item No.":  "vendor_sku",
        "Unit price Ex GST":"sam_list_price",
        "Trade Price Ex":   "sam_trade_price",
        "Trade Line Discount": "sam_discount_pct",
    })
    merged = quote_df.merge(sam_lookup, on="jands_sku", how="left")

    for col in ["list_price", "quoted_price", "quoted_discount_pct"]:
        if col not in merged.columns:
            merged[col] = None

    merged["list_price"] = pd.to_numeric(merged["list_price"], errors="coerce").combine_first(
        pd.to_numeric(merged["sam_list_price"], errors="coerce")
    )
    merged = _derive_price_disc(merged, "quoted_price",    "quoted_discount_pct", "list_price")
    merged = _derive_price_disc(merged, "sam_trade_price", "sam_discount_pct",    "list_price")

    dd_lookup = dd_df.rename(columns={
        "VendorStockCode": "vendor_sku",
        "RRPEx":   "dd_rrp",
        "DealerEx":"dd_price",
    })
    dd_lookup["vendor_sku"] = dd_lookup["vendor_sku"].astype(str).str.strip()
    merged["vendor_sku"]    = merged["vendor_sku"].astype(str).str.strip()
    merged = merged.merge(dd_lookup, on="vendor_sku", how="left")

    merged["dd_discount_pct"] = (
        (1 - merged["dd_price"] / merged["dd_rrp"]) * 100
    ).where(merged["dd_rrp"].notna() & (merged["dd_rrp"] != 0))

    effective = pd.to_numeric(merged["quoted_price"], errors="coerce").combine_first(
        pd.to_numeric(merged["sam_trade_price"], errors="coerce")
    )
    merged["price_diff"] = (effective - merged["dd_price"]).where(merged["dd_price"].notna())

    merged["new_pct_to_match_dd"] = (
        (1 - merged["dd_price"] / merged["list_price"]) * 100
    ).where(
        merged["dd_price"].notna()
        & merged["list_price"].notna()
        & (merged["list_price"] != 0)
    )

    unmatched = merged[merged["vendor_sku"].isna() | merged["dd_price"].isna()].copy()
    return merged.reset_index(drop=True), unmatched.reset_index(drop=True)


# ─────────────────────────────────────────────────────────────
# Manual entry SKU lookup  (supports JANDS SKU or Vendor SKU)
# ─────────────────────────────────────────────────────────────

def lookup_single_sku(
    vendor_sku: str,
    jands_sku_hint: str,
    qty: int,
    sam_df: pd.DataFrame,
    dd_df: pd.DataFrame,
) -> dict:
    sam_by_vendor = sam_df.set_index("Vendor Item No.")
    sam_by_jands  = sam_df.set_index("No.")
    dd_by_vendor  = dd_df.set_index("VendorStockCode") if "VendorStockCode" in dd_df.columns else pd.DataFrame()

    vsku = vendor_sku.strip()
    jsku = jands_sku_hint.strip()

    if not vsku and jsku:
        if jsku in sam_by_jands.index:
            jr = sam_by_jands.loc[jsku]
            if isinstance(jr, pd.DataFrame):
                jr = jr.iloc[0]
            vsku = str(jr.get("Vendor Item No.", "")).strip()

    row: dict = {
        "vendor_sku": vsku,
        "jands_sku": "",
        "description": "",
        "qty": qty,
        "list_price": None,
        "sam_trade_price": None,
        "sam_discount_pct": None,
        "quoted_price": None,
        "quoted_discount_pct": None,
        "dd_rrp": None,
        "dd_price": None,
        "dd_discount_pct": None,
        "price_diff": None,
        "new_pct_to_match_dd": None,
    }

    if vsku and vsku in sam_by_vendor.index:
        sr = sam_by_vendor.loc[vsku]
        if isinstance(sr, pd.DataFrame):
            sr = sr.iloc[0]
        row["jands_sku"] = str(sr.get("No.", ""))
        list_p  = pd.to_numeric(sr.get("Unit price Ex GST"), errors="coerce")
        trade_p = pd.to_numeric(sr.get("Trade Price Ex"),    errors="coerce")
        disc    = pd.to_numeric(sr.get("Trade Line Discount"),errors="coerce")
        if pd.isna(trade_p) and not pd.isna(list_p) and not pd.isna(disc):
            trade_p = list_p * (1 - disc / 100)
        if pd.isna(disc) and not pd.isna(list_p) and not pd.isna(trade_p) and list_p != 0:
            disc = (1 - trade_p / list_p) * 100
        row["list_price"]       = None if pd.isna(list_p)  else float(list_p)
        row["sam_trade_price"]  = None if pd.isna(trade_p) else float(trade_p)
        row["sam_discount_pct"] = None if pd.isna(disc)    else float(disc)
    elif jsku:
        row["jands_sku"] = jsku

    if vsku and not dd_by_vendor.empty and vsku in dd_by_vendor.index:
        dr = dd_by_vendor.loc[vsku]
        if isinstance(dr, pd.DataFrame):
            dr = dr.iloc[0]
        row["dd_rrp"]   = dr.get("RRPEx")
        row["dd_price"] = dr.get("DealerEx")

    if row["dd_rrp"] and row["dd_price"] and row["dd_rrp"] != 0:
        row["dd_discount_pct"] = (1 - row["dd_price"] / row["dd_rrp"]) * 100

    eff = row["sam_trade_price"]
    if eff is not None and row["dd_price"] is not None:
        row["price_diff"] = eff - row["dd_price"]

    if row["dd_price"] and row["list_price"] and row["list_price"] != 0:
        row["new_pct_to_match_dd"] = (1 - row["dd_price"] / row["list_price"]) * 100

    return row


def lookup_manual_skus(entries: list[dict], sam_df: pd.DataFrame, dd_df: pd.DataFrame) -> pd.DataFrame:
    rows = [
        lookup_single_sku(
            e.get("vendor_sku", ""),
            e.get("jands_sku", ""),
            e.get("qty", 1),
            sam_df, dd_df,
        )
        for e in entries
        if e.get("vendor_sku", "").strip() or e.get("jands_sku", "").strip()
    ]
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ─────────────────────────────────────────────────────────────
# Formatting helpers
# ─────────────────────────────────────────────────────────────

def fmt_price(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "N/A"
    try:
        return f"${float(v):,.2f}"
    except (ValueError, TypeError):
        return "N/A"


def fmt_pct(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "N/A"
    try:
        return f"{float(v):.1f}%"
    except (ValueError, TypeError):
        return "N/A"


def build_display_df(df: pd.DataFrame) -> pd.DataFrame:
    def col(name):
        return df[name] if name in df.columns else pd.Series([None] * len(df), index=df.index)

    display = pd.DataFrame(index=df.index)
    display["JANDS SKU"]         = col("jands_sku")
    display["Vendor SKU"]        = col("vendor_sku")
    display["Description"]       = col("description")
    display["Qty"]               = col("qty")
    display["List Price"]        = col("list_price").apply(fmt_price)
    display["SAM Trade Price"]   = col("sam_trade_price").apply(fmt_price)
    display["SAM Disc %"]        = col("sam_discount_pct").apply(fmt_pct)
    display["Quoted Price"]      = col("quoted_price").apply(fmt_price)
    display["Quoted Disc %"]     = col("quoted_discount_pct").apply(fmt_pct)
    display["Competing Price"]   = col("dd_price").apply(fmt_price)
    display["Competitor RRP"]    = col("dd_rrp").apply(fmt_price)
    display["Competitor Disc %"] = col("dd_discount_pct").apply(fmt_pct)
    display["vs Competitor ($)"] = col("price_diff").apply(fmt_price)
    display["New % to Match"]    = col("new_pct_to_match_dd").apply(fmt_pct)
    return display


def add_totals_row(display_df: pd.DataFrame, raw_df: pd.DataFrame) -> pd.DataFrame:
    qty = pd.to_numeric(raw_df.get("qty", pd.Series([1] * len(raw_df))), errors="coerce").fillna(1)

    def weighted_sum(col_name: str) -> str:
        if col_name not in raw_df.columns:
            return ""
        vals = pd.to_numeric(raw_df[col_name], errors="coerce")
        total = (vals * qty).sum()
        return fmt_price(total) if not pd.isna(total) else ""

    totals = {c: "" for c in display_df.columns}
    totals["JANDS SKU"]         = "TOTAL"
    totals["Qty"]               = int(qty.sum())
    totals["List Price"]        = weighted_sum("list_price")
    totals["SAM Trade Price"]   = weighted_sum("sam_trade_price")
    totals["Quoted Price"]      = weighted_sum("quoted_price")
    totals["Competing Price"]   = weighted_sum("dd_price")
    totals["Competitor RRP"]    = weighted_sum("dd_rrp")
    totals["vs Competitor ($)"] = weighted_sum("price_diff")

    return pd.concat([display_df, pd.DataFrame([totals])], ignore_index=True)


def style_table(df_display: pd.DataFrame) -> object:
    last_row = len(df_display) - 1

    def color_diff(val):
        if val == "N/A" or val == "":
            return ""
        try:
            num = float(str(val).replace("$", "").replace(",", ""))
        except ValueError:
            return ""
        if num > 0:
            return "color: #c0392b; font-weight: bold"
        if num < 0:
            return "color: #27ae60; font-weight: bold"
        return ""

    def style_row(row):
        if row.name == last_row:
            return ["font-weight: bold; border-top: 2px solid #555"] * len(row)
        return [""] * len(row)

    styler = df_display.style.apply(style_row, axis=1)
    map_fn = "map" if hasattr(styler, "map") else "applymap"
    if "vs Competitor ($)" in df_display.columns:
        styler = getattr(styler, map_fn)(color_diff, subset=["vs Competitor ($)"])
    if "New % to Match" in df_display.columns:
        styler = getattr(styler, map_fn)(
            lambda v: "font-weight: bold" if v not in ("N/A", "") else "",
            subset=["New % to Match"],
        )
    return styler


# ─────────────────────────────────────────────────────────────
# Export to Excel
# ─────────────────────────────────────────────────────────────

def export_to_excel(df_display: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df_display.to_excel(writer, index=False, sheet_name="Comparison")
        wb = writer.book
        ws = writer.sheets["Comparison"]
        header_fmt = wb.add_format({"bold": True, "bg_color": "#1a3c6e", "font_color": "white", "border": 1})
        red_fmt    = wb.add_format({"bold": True, "font_color": "#c0392b"})
        green_fmt  = wb.add_format({"bold": True, "font_color": "#27ae60"})
        bold_fmt   = wb.add_format({"bold": True})
        total_fmt  = wb.add_format({"bold": True, "top": 2})

        cols = list(df_display.columns)
        for col_num, col_name in enumerate(cols):
            ws.write(0, col_num, col_name, header_fmt)
            ws.set_column(col_num, col_num, max(len(col_name) + 4, 14))

        diff_idx  = cols.index("vs Competitor ($)") if "vs Competitor ($)" in cols else None
        match_idx = cols.index("New % to Match")    if "New % to Match"    in cols else None
        last_row  = len(df_display) - 1

        for row_num, row in enumerate(df_display.itertuples(index=False), start=1):
            is_total = (row_num - 1) == last_row
            for col_num, val in enumerate(row):
                fmt = total_fmt if is_total else None
                if not is_total:
                    if diff_idx is not None and col_num == diff_idx:
                        raw = str(val).replace("$", "").replace(",", "").strip()
                        try:
                            num = float(raw)
                            fmt = red_fmt if num > 0 else (green_fmt if num < 0 else None)
                        except ValueError:
                            pass
                    elif match_idx is not None and col_num == match_idx and val != "N/A":
                        fmt = bold_fmt
                ws.write(row_num, col_num, val, fmt)

    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# Streamlit App
# ─────────────────────────────────────────────────────────────

st.set_page_config(page_title="JANDS Price Comparator", layout="wide", page_icon="📊")

# ── Session state init ───────────────────────────────────────
for key, default in [
    ("dd_df",       None),
    ("dd_status",   "Not loaded"),
    ("sam_df",      None),
    ("sam_status",  "Not loaded"),
    ("mode",        "quote"),
    ("manual_rows", [{"vendor_sku": "", "jands_sku": "", "qty": 1}]),
    ("result_df",   None),
    ("unmatched_df",None),
    ("quote_hash",  None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── Auto-load DD feed ────────────────────────────────────────
if st.session_state.dd_df is None:
    with st.spinner("Checking competitor price feed…"):
        _dd_df, _dd_status = load_dd_feed()
        st.session_state.dd_df    = _dd_df
        st.session_state.dd_status = _dd_status

# ─────────────────────────────────────────────────────────────
# Title + status bar
# ─────────────────────────────────────────────────────────────

st.markdown("# 📊 JANDS Price Comparator")
st.markdown(
    f"<p style='font-size:0.78rem;color:#888;margin:-8px 0 4px;'>"
    f"SAM Report: <strong>{st.session_state.sam_status}</strong>"
    f"&nbsp;|&nbsp;"
    f"Competitor Feed: <strong>{st.session_state.dd_status}</strong>"
    f"</p>",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────
# Data Sources expander
# ─────────────────────────────────────────────────────────────

with st.expander("⚙️ Data Sources & Settings", expanded=(st.session_state.sam_df is None)):
    ds_left, ds_mid, ds_right = st.columns([2, 2, 2])

    with ds_left:
        st.markdown("**SAM Report**")
        sam_file = st.file_uploader("Upload SAM Report (.xlsx)", type=["xlsx"], key="sam_uploader")
        if sam_file:
            try:
                st.session_state.sam_df = parse_sam_report(sam_file.read())
                st.session_state.sam_status = (
                    f"Loaded — {sam_file.name} ({len(st.session_state.sam_df):,} items)"
                )
                st.success(f"Loaded {len(st.session_state.sam_df):,} items")
            except Exception as e:
                st.error(f"Failed to parse: {e}")
                st.session_state.sam_df    = None
                st.session_state.sam_status = f"Error: {e}"

    with ds_mid:
        st.markdown("**Competitor Price Feed**")
        st.caption(st.session_state.dd_status)
        if st.button("🔄 Force Re-download"):
            with st.spinner("Downloading…"):
                _dd_df, _dd_status = load_dd_feed(force=True)
                st.session_state.dd_df    = _dd_df
                st.session_state.dd_status = _dd_status
            st.rerun()

    with ds_right:
        st.markdown("**How to use**")
        st.caption("• Upload SAM report once per session (left panel)")
        st.caption("• Upload a customer quote (Excel) or use Manual Entry mode")
        st.caption("• Red = JANDS pricier than competitor, Green = JANDS cheaper")

st.divider()

# ─────────────────────────────────────────────────────────────
# Mode toggle + quote upload
# ─────────────────────────────────────────────────────────────

col_upload, col_toggle = st.columns([5, 1])
with col_toggle:
    mode = st.radio("Entry mode", ["Upload Quote", "Manual Entry"], label_visibility="collapsed")
    st.session_state.mode = "quote" if mode == "Upload Quote" else "manual"

# ── Quote mode ────────────────────────────────────────────────
if st.session_state.mode == "quote":
    with col_upload:
        quote_file = st.file_uploader(
            "Upload customer quote (Excel .xlsx)",
            type=["xlsx"],
            key="quote_uploader",
        )

    if quote_file:
        if st.session_state.sam_df is None:
            st.warning("Upload the SAM report first (Data Sources above).")
        elif st.session_state.dd_df is None:
            st.warning("Competitor feed not available. Try re-downloading in Data Sources.")
        else:
            file_bytes = quote_file.read()
            file_hash  = hash(file_bytes)
            if file_hash != st.session_state.quote_hash:
                try:
                    quote_df = parse_quote_excel(file_bytes)
                    if quote_df.empty:
                        st.warning("No line items found in the uploaded quote.")
                    else:
                        result, unmatched = build_comparison(
                            quote_df, st.session_state.sam_df, st.session_state.dd_df
                        )
                        st.session_state.result_df    = result
                        st.session_state.unmatched_df = unmatched
                        st.session_state.quote_hash   = file_hash
                except Exception as e:
                    st.error(f"Failed to parse quote: {e}")

# ── Manual entry mode ─────────────────────────────────────────
else:
    with col_upload:
        st.markdown("**Manual SKU Entry** — enter Vendor SKU, JANDS SKU, or both")

    if st.session_state.sam_df is None:
        st.warning("Upload the SAM report first (Data Sources above).")
    else:
        # Header row
        h1, h2, h3, h4, h5 = st.columns([3, 3, 1, 1, 0.5])
        h1.markdown("**Vendor SKU**")
        h2.markdown("**JANDS SKU**")
        h3.markdown("**Qty**")

        rows_to_delete = []
        updated_rows   = []
        for i, row in enumerate(st.session_state.manual_rows):
            c1, c2, c3, c4, c5 = st.columns([3, 3, 1, 1, 0.5])
            with c1:
                vsku = st.text_input("Vendor SKU", value=row["vendor_sku"], key=f"vsku_{i}", label_visibility="collapsed")
            with c2:
                jsku = st.text_input("JANDS SKU",  value=row["jands_sku"],  key=f"jsku_{i}", label_visibility="collapsed")
            with c3:
                qty  = st.number_input("Qty", value=int(row["qty"]), min_value=1, key=f"qty_{i}", label_visibility="collapsed")
            with c5:
                if st.button("✕", key=f"del_{i}") and len(st.session_state.manual_rows) > 1:
                    rows_to_delete.append(i)
            updated_rows.append({"vendor_sku": vsku, "jands_sku": jsku, "qty": qty})

        st.session_state.manual_rows = [
            r for i, r in enumerate(updated_rows) if i not in rows_to_delete
        ]

        ca, cb, _ = st.columns([1, 1, 4])
        with ca:
            if st.button("➕ Add row"):
                st.session_state.manual_rows.append({"vendor_sku": "", "jands_sku": "", "qty": 1})
                st.rerun()
        with cb:
            if st.button("🔍 Look up prices", type="primary"):
                valid = [r for r in st.session_state.manual_rows
                         if r["vendor_sku"].strip() or r["jands_sku"].strip()]
                if not valid:
                    st.warning("Enter at least one SKU.")
                elif st.session_state.dd_df is None:
                    st.warning("Competitor feed not available.")
                else:
                    result = lookup_manual_skus(valid, st.session_state.sam_df, st.session_state.dd_df)
                    if result.empty:
                        st.warning("No results found.")
                    else:
                        st.session_state.result_df    = result.reset_index(drop=True)
                        st.session_state.unmatched_df = result[result["dd_price"].isna()].copy()
                        st.session_state.quote_hash   = None

# ─────────────────────────────────────────────────────────────
# Results
# ─────────────────────────────────────────────────────────────

result_df = st.session_state.result_df

if result_df is not None and not result_df.empty:
    st.divider()

    # ── Summary metrics ──────────────────────────────────────
    total_items   = len(result_df)
    matched_count = int(result_df["dd_price"].notna().sum()) if "dd_price" in result_df.columns else 0
    pricier_count = int((result_df.get("price_diff", pd.Series(dtype=float)) > 0).sum())
    pot_savings   = result_df.loc[
        result_df.get("price_diff", pd.Series(dtype=float)) > 0, "price_diff"
    ].sum() if "price_diff" in result_df.columns else 0.0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Line Items",            total_items)
    m2.metric("Matched to Competitor",        matched_count)
    m3.metric("JANDS Pricier Than Competitor",pricier_count)
    m4.metric("Potential Savings (excl. GST)",f"${pot_savings:,.2f}")

    st.divider()

    # ── Comparison table ─────────────────────────────────────
    st.markdown("### Comparison Table")
    display_df = build_display_df(result_df)
    display_with_totals = add_totals_row(display_df, result_df)
    styled = style_table(display_with_totals)
    st.dataframe(styled, use_container_width=True, hide_index=True)

    # ── Export ───────────────────────────────────────────────
    excel_bytes = export_to_excel(display_with_totals)
    st.download_button(
        label="📥 Export to Excel",
        data=excel_bytes,
        file_name=f"JANDS_Price_Comparison_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    # ── Row management ───────────────────────────────────────
    mgmt_left, mgmt_right = st.columns([2, 3])

    with mgmt_left:
        st.markdown("**Remove rows**")
        row_labels = [
            f"{i+1} — {row.get('jands_sku') or row.get('vendor_sku', f'Row {i+1}')}"
            for i, row in result_df.iterrows()
        ]
        to_remove = st.multiselect(
            "Select rows to remove",
            options=list(range(len(result_df))),
            format_func=lambda i: row_labels[i],
            label_visibility="collapsed",
        )
        if st.button("🗑 Remove selected rows", disabled=not to_remove):
            st.session_state.result_df = result_df.drop(index=to_remove).reset_index(drop=True)
            if st.session_state.unmatched_df is not None:
                st.session_state.unmatched_df = st.session_state.unmatched_df[
                    st.session_state.unmatched_df.index.isin(
                        st.session_state.result_df.index
                    )
                ].reset_index(drop=True)
            st.rerun()

    with mgmt_right:
        st.markdown("**Add a row manually**")
        a1, a2, a3, a4 = st.columns([3, 3, 1, 1])
        with a1:
            add_vsku = st.text_input("Vendor SKU", key="add_vsku", label_visibility="collapsed", placeholder="Vendor SKU")
        with a2:
            add_jsku = st.text_input("JANDS SKU",  key="add_jsku", label_visibility="collapsed", placeholder="JANDS SKU")
        with a3:
            add_qty  = st.number_input("Qty", min_value=1, value=1, key="add_qty", label_visibility="collapsed")
        with a4:
            if st.button("➕ Add", key="add_row_btn"):
                if not add_vsku.strip() and not add_jsku.strip():
                    st.warning("Enter a Vendor SKU or JANDS SKU.")
                elif st.session_state.sam_df is None:
                    st.warning("SAM report not loaded.")
                elif st.session_state.dd_df is None:
                    st.warning("Competitor feed not available.")
                else:
                    new_row = lookup_single_sku(
                        add_vsku, add_jsku, add_qty,
                        st.session_state.sam_df, st.session_state.dd_df,
                    )
                    new_df = pd.DataFrame([new_row])
                    st.session_state.result_df = pd.concat(
                        [st.session_state.result_df, new_df], ignore_index=True
                    )
                    st.rerun()

    # ── Unmatched items ──────────────────────────────────────
    unmatched_df = st.session_state.unmatched_df
    if unmatched_df is not None and not unmatched_df.empty:
        with st.expander(f"⚠️ Unmatched items ({len(unmatched_df)}) — not found in competitor feed"):
            show_cols = [c for c in ["jands_sku", "vendor_sku", "description", "list_price", "sam_trade_price", "quoted_price"] if c in unmatched_df.columns]
            um = unmatched_df[show_cols].copy()
            um.columns = [c.replace("_", " ").title() for c in um.columns]
            st.dataframe(um, use_container_width=True, hide_index=True)

elif st.session_state.mode == "quote" and st.session_state.sam_df is not None:
    st.info("Upload a customer quote above (Excel) to see the price comparison.")