"""Core DNO logic: portfolio parsing, checkup building, DNO template filling.

Framework-free so it can be tested on its own and reused by the web app.
"""
import io
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

MODES = ("mpn", "master_id", "sku", "page_id", "secondary_id")

TEMPORARY_REASONS = [
    "Excess Inventory", "Listing Issues", "Ongoing Negotiations",
    "Production Delays - Single Item", "Production Delays - Missing Component",
    "Production Delays - Component", "Seasonal Item",
    "SKU Rationalization - Buybox", "SKU Rationalization - Margin",
    "Waiting - Inventory Position", "Waiting - Launch",
]
PERMANENT_REASONS = [
    "Marketplace Regulations", "New Generation", "Partner Termination",
    "Production Discontinued", "Amaczar Legacy DNO", "Listing Deprecated",
]

UNDNO_DELETE = "<<Delete>>"
DNO_ATTRIBUTES = [
    "listing_dno_note", "listing_dno_reason_code", "listing_dno_is_temporary",
    "listing_dno_start_date", "listing_dno_end_date",
]

TEMPLATE_HEADERS = [
    "listing_action", "listing_id", "listing_dno_note", "listing_dno_reason_code",
    "listing_dno_is_temporary", "listing_dno_start_date", "listing_dno_end_date",
    "listing_is_shipable", "listing_sync_avail_inv",
]


# ---------- Helpers ----------
def safe_str(x):
    return str(x).strip() if pd.notna(x) else ""


def normalize_cols(df):
    return {c.strip().lower().replace(" ", "_"): c for c in df.columns}


def find_sheet_like(xls, candidates):
    for s in xls.sheet_names:
        low = s.strip().lower()
        for cand in candidates:
            if cand in low:
                return s
    return None


def find_col(norm_map, candidates):
    for cand in candidates:
        if cand in norm_map:
            return norm_map[cand]
    for k, v in norm_map.items():
        for cand in candidates:
            if cand in k:
                return v
    return None


def validate_date_str(s):
    try:
        datetime.strptime(s, "%Y-%m-%d")
        return True
    except (ValueError, TypeError):
        return False


def parse_values(text):
    """Split comma/newline separated user input into a clean list."""
    return [s.strip() for s in (text or "").replace("\r", "\n").replace(",", "\n").split("\n") if s.strip()]


def _read_sheet(source, sheet):
    try:
        df = pd.read_excel(source, sheet_name=sheet, header=2, dtype=str)
        if df.shape[1] < 2:
            raise ValueError("fallback")
    except Exception:
        df = pd.read_excel(source, sheet_name=sheet, header=0, dtype=str)
    return df.fillna("").map(lambda x: x.strip() if isinstance(x, str) else x)


# ---------- Step 1: portfolio ----------
PROD_REQUIRED = {
    "sku": ["listing_mp_primary_id", "listing_mp_primaryid", "mp_primary_id", "sku", "product_sku", "mp_primary"],
    "listing_id": ["listing_id", "listingid", "mp_listing_id", "mp_list_id"],
    "marketplace": ["listing_mp_name", "listingmpname", "mp_name", "marketplace"],
    "seller": ["listing_mp_seller_name", "listingmp_seller_name", "seller_name", "seller_account",
               "sellername", "seller_id", "sellerid"],
    "master": ["master_id", "masterid", "master"],
}
# "Marketplace Page ID" holds the ASIN / Item ID value; "Marketplace Page ID Type" says which kind it is.
ASIN_CANDIDATES = ["listing_mp_page_id", "listing_mp_pageid", "marketplace_page_id", "mp_page_id", "page_id", "asin", "mp_pageid"]
PAGE_TYPE_CANDIDATES = ["listing_mp_page_id_type", "marketplace_page_id_type", "mp_page_id_type", "page_id_type"]
SECONDARY_ID_CANDIDATES = ["listing_mp_secondary_id", "marketplace_secondary_id", "mp_secondary_id",
                           "listing_secondary_id", "secondary_id", "secondaryid"]


def _key(s):
    """Lower-case and drop non-alphanumerics: 'Item ID' / 'ITEM_ID' -> 'itemid'."""
    return "".join(c for c in str(s).lower() if c.isalnum())


def _strip_zeros(s):
    """UPC/GTIN compare ignoring leading zeros (Excel often drops them)."""
    return str(s).strip().lstrip("0")


def load_portfolio(file_like):
    """Parse the portfolio workbook. Returns a picklable dict of state."""
    data = file_like.read() if hasattr(file_like, "read") else file_like
    xls = pd.ExcelFile(io.BytesIO(data))

    product_sheet = find_sheet_like(
        xls, ["product listing", "product listings", "product_listing", "product_listings"])
    if not product_sheet:
        raise ValueError("Product Listings sheet not found. Sheets: " + ", ".join(xls.sheet_names))

    prod_df = _read_sheet(io.BytesIO(data), product_sheet)
    prod_norm = normalize_cols(prod_df)

    cols, missing = {}, []
    for key, cands in PROD_REQUIRED.items():
        cols[key] = find_col(prod_norm, cands)
        if not cols[key]:
            missing.append(key)
    if missing:
        raise ValueError(f"Missing required Product Listings columns: {missing}. "
                         f"Found: {list(prod_df.columns)}")
    # keep "...page_id_type" out of the page-id search so the substring fallback can't pick it
    page_norm = {k: v for k, v in prod_norm.items() if not k.endswith("_type")}
    cols["asin"] = find_col(page_norm, ASIN_CANDIDATES)  # the Marketplace Page ID column
    cols["page_type"] = find_col(prod_norm, PAGE_TYPE_CANDIDATES)
    cols["secondary_id"] = find_col(prod_norm, SECONDARY_ID_CANDIDATES)
    # current DNO / shipping status, shown in the checkup table (optional columns)
    cols["ship"] = prod_norm.get("listing_is_shipable")
    cols["dno_temp"] = prod_norm.get("listing_dno_is_temporary")
    cols["dno_reason"] = prod_norm.get("listing_dno_reason_code")
    cols["dno_start"] = prod_norm.get("listing_dno_start_date")
    cols["sync"] = prod_norm.get("listing_sync_avail_inv")

    master_df, master_cols = None, {}
    master_sheet = find_sheet_like(
        xls, ["master products", "master_products", "master product", "masterproduct"])
    if master_sheet:
        master_df = _read_sheet(io.BytesIO(data), master_sheet)
        mnorm = normalize_cols(master_df)
        master_cols = {
            "mpn": find_col(mnorm, ["mpn", "mpn_code", "mpn_number"]),
            "master_id": find_col(mnorm, ["master_id", "masterid", "master"]),
        }

    marketplaces = sorted({safe_str(x) for x in prod_df[cols["marketplace"]] if safe_str(x)})
    sellers = sorted({safe_str(x) for x in prod_df[cols["seller"]] if safe_str(x)})
    if not marketplaces or not sellers:
        raise ValueError("No marketplaces or sellers detected in Product Listings sheet.")

    # sellers that actually sell in each marketplace (drives the dependent dropdown)
    mp_col, seller_col = cols["marketplace"], cols["seller"]
    sellers_by_mp = {
        mp: sorted({safe_str(s) for s in grp[seller_col] if safe_str(s)})
        for mp, grp in prod_df[prod_df[mp_col] != ""].groupby(mp_col)
    }

    return {
        "prod_df": prod_df, "cols": cols,
        "master_df": master_df, "master_cols": master_cols,
        "marketplaces": marketplaces, "sellers": sellers,
        "sellers_by_mp": sellers_by_mp,
        "brand": detect_brand(master_df, master_cols),
    }


BRAND_CANDIDATES = ["master_brand_name", "master_brand", "brand_name", "brand"]


def detect_brand(master_df, master_cols):
    """Most common brand in the Master Products sheet, or None if there is no brand column."""
    if master_df is None:
        return None
    norm = normalize_cols(master_df)
    col = next((norm[c] for c in BRAND_CANDIDATES if c in norm), None)
    if not col:
        return None
    vals = master_df[col].map(safe_str)
    vals = vals[vals != ""]
    return vals.value_counts().index[0] if len(vals) else None


_CANON = ["sku", "listing_id", "marketplace", "seller", "master", "asin", "page_type", "secondary_id",
          "ship", "dno_temp", "dno_reason", "dno_start", "sync"]


def _dno_type(temp, reason):
    """Current DNO status from listing_dno_is_temporary + reason code."""
    # listing_dno_is_temporary: Y = temporary DNO, N = permanent DNO, blank = no DNO applied
    t = str(temp).strip().upper()
    if t in ("Y", "YES", "TRUE", "1"):
        return "Temporary"
    if t in ("N", "NO", "FALSE", "0"):
        return "Permanent"
    return "-"


def load_portfolios(files):
    """files: list of (file_name, bytes). Parses each portfolio and merges them into one
    state, tagging every Product Listings row with its brand. One file works the same way."""
    if not files:
        raise ValueError("No portfolio file uploaded.")

    parts = []
    for fname, data in files:
        try:
            st_ = load_portfolio(data)
        except ValueError as e:
            raise ValueError(f"{fname}: {e}")
        parts.append((st_["brand"] or fname.rsplit(".", 1)[0], st_))

    present = {k: False for k in _CANON}
    master_present = {"mpn": False, "master_id": False}
    orig_columns, prod_frames, master_frames = [], [], []

    for brand, st_ in parts:
        df = st_["prod_df"]
        orig_columns += [c for c in df.columns if c not in orig_columns]
        ren = {}
        for k in _CANON:
            if st_["cols"].get(k):
                ren[st_["cols"][k]] = "_c_" + k
                present[k] = True
        df = df.rename(columns=ren)
        for k in _CANON:
            if "_c_" + k not in df.columns:
                df["_c_" + k] = ""
        df["_c_brand"] = brand
        prod_frames.append(df)

        mdf, mc = st_["master_df"], st_["master_cols"]
        if mdf is not None:
            mren = {mc[k]: "_c_" + k for k in ("mpn", "master_id") if mc.get(k)}
            for k in mren.values():
                master_present[k[3:]] = True
            mdf = mdf.rename(columns=mren)
            for k in ("mpn", "master_id"):
                if "_c_" + k not in mdf.columns:
                    mdf["_c_" + k] = ""
            master_frames.append(mdf[["_c_mpn", "_c_master_id"]])

    prod_df = pd.concat(prod_frames, ignore_index=True)
    cols = {k: ("_c_" + k if present[k] else None) for k in _CANON}
    cols["brand"] = "_c_brand"
    master_df = pd.concat(master_frames, ignore_index=True) if master_frames else None
    master_cols = {k: ("_c_" + k if master_present[k] else None) for k in ("mpn", "master_id")}

    mp_col, seller_col = cols["marketplace"], cols["seller"]
    brands = list(dict.fromkeys(b for b, _ in parts))
    return {
        "prod_df": prod_df, "cols": cols,
        "master_df": master_df, "master_cols": master_cols,
        "marketplaces": sorted({safe_str(x) for x in prod_df[mp_col] if safe_str(x)}),
        "sellers": sorted({safe_str(x) for x in prod_df[seller_col] if safe_str(x)}),
        "sellers_by_mp": {
            mp: sorted({safe_str(s) for s in grp[seller_col] if safe_str(s)})
            for mp, grp in prod_df[prod_df[mp_col] != ""].groupby(mp_col)
        },
        "brands": brands,
        "brand": brands[0] if len(brands) == 1 else "Multi-Brand",
        "orig_columns": orig_columns,
    }


def capabilities(state):
    mc = state["master_cols"]
    return {
        "page_id": state["cols"]["asin"] is not None,
        "secondary_id": state["cols"].get("secondary_id") is not None,
        "mpn": bool(state["master_df"] is not None and mc.get("mpn") and mc.get("master_id")),
    }


# ---------- Step 2: checkup ----------
def build_checkup(state, mode, values, scope_all, marketplaces, sellers):
    """Return (rows, missing). rows: one dict per (input value, listing_id). missing: the
    inputs that matched no LIDs, as (value, reason) pairs."""
    if mode not in MODES:
        raise ValueError(f"Unknown mode: {mode}")
    cols, prod = state["cols"], state["prod_df"]
    caps = capabilities(state)
    if mode == "page_id" and not caps["page_id"]:
        raise ValueError("Marketplace Page ID column not found in Product Listings; Page ID mode unavailable.")
    if mode == "secondary_id" and not caps["secondary_id"]:
        raise ValueError("Secondary ID column not found in Product Listings; Secondary ID mode unavailable.")
    if mode == "mpn" and not caps["mpn"]:
        raise ValueError("Master Products sheet or its 'mpn'/'master_id' columns not found; MPN mode unavailable.")
    if not scope_all and not marketplaces:
        raise ValueError("Select at least one marketplace for 'LIDs Marketplace-wise'.")

    def in_scope(df):
        if scope_all:
            return df
        df = df[df[cols["marketplace"]].isin(marketplaces)]
        if sellers:  # no seller chosen = every seller in the selected marketplaces
            df = df[df[cols["seller"]].isin(sellers)]
        return df

    def field(rep, k):
        return safe_str(rep[cols[k]]) if cols.get(k) else ""

    def rows_for(base, key):
        """key: dict of leading columns identifying the input value. Empty list = no LIDs."""
        base = in_scope(base)
        out = []
        if base.empty:
            return out
        for (mp, seller), sub in base.groupby([cols["marketplace"], cols["seller"]], sort=True):
            lids = sub[cols["listing_id"]].replace("", pd.NA).dropna().astype(str).unique().tolist()
            for lid in lids:
                rep = sub[sub[cols["listing_id"]] == lid].iloc[0]
                temp, reason = field(rep, "dno_temp"), field(rep, "dno_reason")
                dno_type = _dno_type(temp, reason)
                out.append({**key, "master_id": safe_str(rep[cols["master"]]),
                            "brand": safe_str(rep[cols["brand"]]) if cols.get("brand") else "",
                            "listing_id": safe_str(lid), "sku": safe_str(rep[cols["sku"]]),
                            "marketplace": safe_str(mp), "seller_account": safe_str(seller),
                            "able_to_ship": field(rep, "ship"),
                            "dno": "N" if dno_type == "-" else "Y",  # a value in the temporary field = DNO applied
                            "dno_temporary": temp,
                            "dno_type": dno_type,
                            "dno_reason_code": reason,
                            "sync_inventory": field(rep, "sync")})
        return out

    if mode == "mpn" and not values:
        mc, mdf = state["master_cols"], state["master_df"]
        values = mdf[mc["mpn"]].replace("", pd.NA).dropna().astype(str).unique().tolist()
    if not values:
        raise ValueError("Enter at least one value.")

    rows, missing = [], []  # missing: (value, reason) for inputs that produced no LIDs
    for v in values:
        before = len(rows)
        base_found = True  # did the value exist in the portfolio at all (before the marketplace filter)?
        if mode == "sku":
            base = prod[prod[cols["sku"]] == v]
            rows += rows_for(base, {})
            base_found = not base.empty
        elif mode == "page_id":  # ASIN, Item ID, etc. all live in Marketplace Page ID
            base = prod[prod[cols["asin"]] == v]
            rows += rows_for(base, {"page_id": v})
            base_found = not base.empty
        elif mode == "secondary_id":
            sec = prod[cols["secondary_id"]].map(_strip_zeros)
            base = prod[sec == _strip_zeros(v)]
            rows += rows_for(base, {"secondary_id": v})
            base_found = not base.empty
        elif mode == "master_id":
            base = prod[prod[cols["master"]] == v]
            rows += rows_for(base, {"master_id": v})
            base_found = not base.empty
        else:  # mpn
            mc, mdf = state["master_cols"], state["master_df"]
            mids = mdf[mdf[mc["mpn"]] == v][mc["master_id"]].replace("", pd.NA).dropna().astype(str).unique().tolist()
            if not mids:
                missing.append((v, "not found in Master Products"))
                continue
            base_found = False
            for mid in mids:
                base = prod[prod[cols["master"]] == mid]
                base_found = base_found or not base.empty
                rows += rows_for(base, {"mpn": v, "master_id": mid})
        if len(rows) == before:
            missing.append((v, "no LIDs for the selected marketplace / seller" if base_found
                            else "not found in the portfolio"))

    return rows, missing


def checkup_to_xlsx(rows):
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False, sheet_name="Checkup")
    return buf.getvalue()


# ---------- Step 3: DNO file ----------
def generate_dno(template_path, listing_ids, cfg):
    """Fill the formatted template. cfg keys: temporary(bool), reason, start, end,
    shippable, sync_inv, note. Returns xlsx bytes."""
    undno = cfg.get("undno", False)
    start, end = cfg.get("start", "").strip(), cfg.get("end", "").strip()
    if not undno:
        if not validate_date_str(start):
            raise ValueError("Invalid start date. Use YYYY-MM-DD.")
        if cfg["temporary"] and not validate_date_str(end):
            raise ValueError("Invalid end date for Temporary DNO. Use YYYY-MM-DD.")
        allowed = TEMPORARY_REASONS if cfg["temporary"] else PERMANENT_REASONS
        if cfg.get("reason") not in allowed:
            raise ValueError("Reason code is not valid for the selected DNO type.")
    if not listing_ids:
        raise ValueError("No listing IDs found in Checkup.")

    wb = load_workbook(template_path)
    if "Product Listings" not in wb.sheetnames:
        raise ValueError("Template must contain a 'Product Listings' sheet.")
    ws = wb["Product Listings"]

    header_row, first_data_row = 3, 4
    header_map = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val is not None:
            header_map[str(val).strip().lower().replace(" ", "_")] = col
    missing = [h for h in TEMPLATE_HEADERS if h not in header_map]
    if missing:
        raise ValueError(f"Template 'Product Listings' sheet missing required columns: {missing}")

    for r in range(first_data_row, ws.max_row + 1):
        for h in TEMPLATE_HEADERS:
            ws.cell(row=r, column=header_map[h]).value = None

    for i, lid in enumerate(listing_ids):
        r = first_data_row + i
        if undno:  # un-DNO = every DNO attribute set to <<Delete>>
            values = {h: UNDNO_DELETE for h in DNO_ATTRIBUTES}
        else:
            values = {
                "listing_dno_note": cfg.get("note", "").strip(),
                "listing_dno_reason_code": cfg["reason"],
                "listing_dno_is_temporary": "Y" if cfg["temporary"] else "N",
                "listing_dno_start_date": start,
                "listing_dno_end_date": end if cfg["temporary"] else "",
            }
        values.update({
            "listing_action": "Update Existing",
            "listing_id": lid,
            "listing_is_shipable": cfg.get("shippable", "") or "",
            "listing_sync_avail_inv": cfg.get("sync_inv", "") or "",
        })
        for h, v in values.items():
            ws.cell(row=r, column=header_map[h]).value = v

    if cfg.get("discontinue_mids") and not undno:
        _mark_manufacturer_discontinued(wb, cfg["discontinue_mids"], start)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _hkey(s):
    return "".join(c for c in str(s).lower() if c.isalnum())


def _mark_manufacturer_discontinued(wb, mids, date):
    """Master Products sheet: one row per Master ID, Manufacturer Discontinued? = Y and the
    Manufacturer Discontinued Date = the DNO start date (same text format as the listing rows)."""
    if "Master Products" not in wb.sheetnames:
        raise ValueError("Template must contain a 'Master Products' sheet to mark Manufacturer Discontinued.")
    ws = wb["Master Products"]
    header_row, first_data_row = 3, 4
    found = {}
    for col in range(1, ws.max_column + 1):
        # row 3 holds the system names (master_id, is_discontinued, ...), row 2 the friendly labels
        k = _hkey(ws.cell(row=header_row, column=col).value or "")
        label = _hkey(ws.cell(row=header_row - 1, column=col).value or "")
        if k == "masterid":
            found.setdefault("master_id", col)
        elif k in ("action", "masteraction"):
            found.setdefault("action", col)
        elif k == "isdiscontinued" or ("manufacturerdiscontinued" in label and "date" not in label):
            found.setdefault("flag", col)
        elif k == "discontinueddate" or "manufacturerdiscontinueddate" in label:
            found.setdefault("date", col)
    missing = [n for n, key in (("Master ID", "master_id"), ("Manufacturer Discontinued?", "flag"),
                                ("Manufacturer Discontinued Date", "date")) if key not in found]
    if missing:
        raise ValueError(f"Template 'Master Products' sheet (headers on row 3) is missing: {missing}")

    for r in range(first_data_row, ws.max_row + 1):  # clear leftovers in the columns we fill
        for col in found.values():
            ws.cell(row=r, column=col).value = None
    for i, mid in enumerate(dict.fromkeys(mids)):
        r = first_data_row + i
        if "action" in found:
            ws.cell(row=r, column=found["action"]).value = "Update Existing"
        ws.cell(row=r, column=found["master_id"]).value = mid
        ws.cell(row=r, column=found["flag"]).value = "Y"
        ws.cell(row=r, column=found["date"]).value = date
