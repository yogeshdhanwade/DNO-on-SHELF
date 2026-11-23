"""Product Catalog Auto-DNO on Shelf - Streamlit app (uses core.py).

Run:  streamlit run streamlit_app.py
"""
import base64
import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import core

st.set_page_config(page_title="Product Catalog Auto-DNO on Shelf", layout="wide")
st.title("Product Catalog Auto-DNO on Shelf")
st.caption("Bulk DNO & un-DNO for every LID on Shelf, from portfolio to upload file in minutes")


@st.cache_resource
def template_store():
    """Shared across sessions while the app process is alive."""
    return {}


store = template_store()
ss = st.session_state

# Default template: any .xlsx with "template" in its name, committed next to this file in the repo.
APP_DIR = os.path.dirname(os.path.abspath(__file__))
if "bytes" not in store:
    candidates = sorted((n for n in os.listdir(APP_DIR)
                         if n.lower().endswith(".xlsx") and "template" in n.lower()),
                        key=lambda n: (n != "DNO_Template.xlsx", n))  # DNO_Template.xlsx wins if present
    if candidates:
        with open(os.path.join(APP_DIR, candidates[0]), "rb") as f:
            store["bytes"], store["name"] = f.read(), candidates[0]

MODE_LABELS = {"mpn": "MPN", "master_id": "Master ID (MID)", "sku": "SKU",
               "page_id": "Item ID / ASIN", "secondary_id": "GTIN (UPC)"}
MODE_NOUNS = {"mpn": "MPNs", "master_id": "Master IDs", "sku": "SKUs",
              "page_id": "Item IDs / ASINs", "secondary_id": "GTINs (UPCs)"}


if "bytes" not in store:
    st.error("DNO template not found. Add an .xlsx with 'template' in its name next to streamlit_app.py on GitHub.")

# ---------- Upload ----------
st.header("Upload Brand's Portfolio")
pfs = st.file_uploader("Brand portfolio - add one or more (.xlsx with Product Listings & optional Master Products)",
                       type=["xlsx", "xlsm"], accept_multiple_files=True, key="pf")
if pfs:
    key = tuple((f.name, f.size) for f in pfs)
    if ss.get("pf_key") != key:
        try:
            with st.spinner("Reading portfolio..."):
                ss["portfolio"] = core.load_portfolios([(f.name, f.getvalue()) for f in pfs])
            ss["pf_key"] = key
            ss.pop("checkup", None)
            ss.pop("dno_file", None)
        except Exception as e:
            ss.pop("portfolio", None)
            ss.pop("pf_key", None)
            st.error(str(e))

if not pfs:  # all files removed from the uploader -> start clean
    for k in ("portfolio", "pf_key", "checkup", "brand"):
        ss.pop(k, None)

state = ss.get("portfolio")
if not state:
    st.info("Upload a portfolio file to continue.")
    st.stop()

caps = core.capabilities(state)
brand = state["brand"]
st.caption(f"{len(state['brands'])} brand(s) - {len(state['prod_df'])} listing rows - "
           f"{len(state['marketplaces'])} marketplace(s) - {len(state['sellers'])} seller(s)")

# ---------- Find listings ----------
st.header("Fetch Listing IDs by Identifier")
modes = [m for m in core.MODES if caps.get(m, True)]
mode = st.segmented_control("Look up listings by", modes, default=modes[0],
                            format_func=MODE_LABELS.get, key="mode")
if mode is None:  # user clicked the selected option again
    st.info("Pick an option above to continue.")
    st.stop()
hidden = [MODE_LABELS[m] for m in core.MODES if m not in modes]
if hidden:
    st.caption("Not available for this portfolio (column not found): " + ", ".join(hidden))
    st.caption("Columns found in Product Listings: " + ", ".join(map(str, state["orig_columns"])))
st.caption("Brands: **" + ", ".join(state["brands"]) + "**")
hint = " - leave blank to use all from Master Products" if mode == "mpn" else ""
values = st.text_area(f"Paste {MODE_NOUNS[mode]} (comma or one per line){hint}", height=120)

# Marketplaces / sellers that actually have LIDs for the pasted identifiers (nothing pasted yet = all).
parsed = core.parse_values(values)
avail_key = (mode, tuple(parsed), ss.get("pf_key"))
if ss.get("avail_key") != avail_key:
    avail = {mp: list(s) for mp, s in state["sellers_by_mp"].items()}
    if parsed or mode == "mpn":
        try:
            prev, _ = core.build_checkup(state, mode, parsed, True, [], [])
            avail = {}
            for r in prev:
                avail.setdefault(r["marketplace"], set()).add(r["seller_account"])
            avail = {mp: sorted(s) for mp, s in avail.items()}
        except ValueError:
            pass
    ss["avail_key"], ss["avail"] = avail_key, avail
avail = ss["avail"]

ALL_LIDS, MP_LIDS = "All LIDs", "LIDs Marketplace-wise"
scope = st.segmented_control("Listing Coverage", [ALL_LIDS, MP_LIDS],
                             default=ALL_LIDS, key="scope") or ALL_LIDS
mps, sellers = [], []
if scope == MP_LIDS:
    c1, c2 = st.columns(2)
    # a soft colour per marketplace tag (by name; position-based as a fallback)
    palette = ["#2dd4bf", "#38bdf8", "#a78bfa", "#fb7185", "#fbbf24",
               "#34d399", "#fb923c", "#f472b6", "#a3e635", "#818cf8"]
    css = []
    for i, c in enumerate(palette, 1):
        css.append(f'.st-key-sel_mps [data-baseweb="tag"]:nth-child({len(palette)}n+{i})'
                   f'{{background:{c}33 !important;color:{c} !important;border:1px solid {c}88;}}')
    for i, name in enumerate(state["marketplaces"]):
        c = palette[i % len(palette)]
        safe = name.replace("\\", "\\\\").replace('"', '\\"')
        css.append(f'.st-key-sel_mps [data-baseweb="tag"]:has(span[title="{safe}"])'
                   f'{{background:{c}33 !important;color:{c} !important;border:1px solid {c}88;}}')
    st.markdown("<style>" + "".join(css) + "</style>", unsafe_allow_html=True)
    mp_opts = sorted(avail)
    ss["sel_mps"] = [m for m in ss.get("sel_mps", []) if m in mp_opts]  # drop picks that no longer apply
    mps = c1.multiselect("Marketplace(s)", mp_opts, key="sel_mps",
                         placeholder="No marketplace has LIDs for these identifiers" if not mp_opts else "Choose marketplace(s)")
    # only sellers that sell in the chosen marketplaces (for the pasted identifiers)
    seller_opts = sorted({s for mp in mps for s in avail.get(mp, [])})
    ss["sel_sellers"] = [s for s in ss.get("sel_sellers", []) if s in seller_opts]
    # a marketplace with a single seller gets that seller picked for you, whatever else is chosen
    handled = set(ss.get("auto_mps", []))
    for mp in mps:
        if mp not in handled and len(avail.get(mp, [])) == 1 and avail[mp][0] not in ss["sel_sellers"]:
            ss["sel_sellers"] = ss["sel_sellers"] + avail[mp]
    ss["auto_mps"] = list(mps)
    sellers = c2.multiselect("Seller(s)", seller_opts, key="sel_sellers",
                             placeholder="Select a marketplace first" if not mps else "Choose seller(s)")
    # a marketplace where no seller is ticked keeps all of its sellers
    ticked = set(sellers)
    keep = set(sellers)
    for mp in mps:
        if not set(avail.get(mp, [])) & ticked:
            keep |= set(avail.get(mp, []))
    sellers = sorted(keep)

if st.button("Fetch LIDs", type="primary"):
    ss.pop("dno_file", None)
    try:
        ss["checkup"], ss["missing"] = core.build_checkup(state, mode, parsed, scope == ALL_LIDS, mps, sellers)
        ss["ck_mode"] = mode
        ss["ck_id"] = ss.get("ck_id", 0) + 1  # new table -> new editor, all rows ticked again
        ss["brand"] = brand
    except ValueError as e:
        ss.pop("checkup", None)
        st.error(str(e))

rows = ss.get("checkup")
if rows is None:
    st.stop()
missing = ss.get("missing", [])


def show_missing():
    """Footnote listing the inputs that matched no LIDs, grouped by reason."""
    by_reason = {}
    for value, why in missing:
        by_reason.setdefault(why, []).append(value)
    for why, vals in by_reason.items():
        more = f" ... and {len(vals) - 25} more" if len(vals) > 25 else ""
        st.caption(f"No LIDs ({why}): " + ", ".join(vals[:25]) + more)


if not rows:
    st.warning("No listing IDs found.")
    show_missing()
    st.stop()

all_lids = list(dict.fromkeys(r["listing_id"] for r in rows))
safe_brand = "".join(c for c in ss["brand"] if c.isalnum() or c in ("_", "-", " ")).strip() or "Brand"

# Results table: rows are grouped by Master ID, and Master IDs that cover 2+ SKUs share one colour.
df = pd.DataFrame(rows).fillna("")
if len(state["brands"]) < 2 and "brand" in df.columns:
    df = df.drop(columns="brand")  # only worth showing when several brands were uploaded
# group by Master ID so SKUs of one Master ID sit together
df = df.sort_values("master_id", kind="stable", key=lambda s: s.where(s != "", "￿")).reset_index(drop=True)
multi = [m for m, n in df[df["master_id"] != ""].groupby("master_id")["sku"].nunique().items() if n > 1]
tints = ["#2dd4bf", "#38bdf8", "#a78bfa", "#fb7185", "#fbbf24", "#34d399", "#fb923c", "#f472b6"]
tint_of = {m: tints[i % len(tints)] for i, m in enumerate(multi)}


def tint_row(row):
    c = tint_of.get(row["master_id"])
    return [f"background-color: {c}2e" if c else ""] * len(row)


df.insert(0, "Apply", True)  # every matched LID starts ticked

# Default view: the searched identifier first (never repeated), then a short list of columns. Columns
# that are mostly empty are left out; every attribute stays available from the eye icon in the table toolbar.
ident = ss.get("ck_mode", mode)
wanted = [ident, "brand", "master_id", "listing_id", "sku", "marketplace", "seller_account", "able_to_ship", "dno"]
always = {ident, "brand", "listing_id", "marketplace", "seller_account"}
visible = []
for c in wanted:
    if c not in df.columns or c in visible:
        continue
    if c not in always and df[c].isin(["", "-"]).mean() > 0.5:
        continue
    visible.append(c)

LABELS = {"mpn": "MPN", "master_id": "Master ID", "sku": "SKU", "page_id": "Item ID / ASIN",
          "secondary_id": "GTIN (UPC)", "brand": "Brand", "listing_id": "LID", "marketplace": "Marketplace",
          "seller_account": "Seller account", "able_to_ship": "Able to ship", "dno": "DNO (Y/N)",
          "dno_temporary": "DNO temporary", "dno_type": "DNO type", "dno_reason_code": "DNO reason code",
          "sync_inventory": "Sync inventory"}
config = {c: st.column_config.TextColumn(LABELS.get(c, c)) for c in df.columns if c != "Apply"}
config["Apply"] = st.column_config.CheckboxColumn("Apply", default=True, width="small")

if multi:
    st.caption("Rows sharing a colour are SKUs under the same Master ID.")
st.caption("Untick any listing you do NOT want to DNO / un-DNO. Use the eye icon in the table's toolbar to show more columns.")
edited = st.data_editor(
    df.style.apply(tint_row, axis=1), use_container_width=True, hide_index=True,
    key=f"checkup_editor_{ss.get('ck_id', 0)}", height=min(600, 38 * (len(df) + 1) + 3),
    column_order=["Apply"] + visible, column_config=config,
    disabled=[c for c in df.columns if c != "Apply"],
)
show_missing()
picked = edited[edited["Apply"]]
listing_ids = list(dict.fromkeys(picked["listing_id"]))
sku_count = picked.loc[picked["sku"] != "", "sku"].nunique()
st.write(f"**{len(listing_ids)}** of {len(all_lids)} listing ID(s) selected.")
if not listing_ids:
    st.warning("No listing IDs selected. Tick at least one to continue.")
    st.stop()

# ---------- DNO ----------
st.header("DNO Settings")
if "bytes" not in store:
    st.warning("DNO template is missing from the repository.")
    st.stop()

dno_type = st.radio("DNO Type", ["Temporary", "Permanent", "UN - DNO"], horizontal=True)
undno = dno_type == "UN - DNO"
temp = dno_type == "Temporary"
reason, start, end, note = "", "", "", ""
if not undno:
    reason = st.selectbox("Reason Code", core.TEMPORARY_REASONS if temp else core.PERMANENT_REASONS)
    c1, c2 = st.columns(2)
    start_d = c1.date_input("Start Date", value=None, format="YYYY-MM-DD")
    end_d = c2.date_input("End Date", value=None, format="YYYY-MM-DD") if temp else None
    start = start_d.isoformat() if start_d else ""
    end = end_d.isoformat() if end_d else ""
c3, c4 = st.columns(2)
ship = c3.selectbox("Able to ship to marketplace (Shippable)", ["", "Y", "N"])
sync = c4.selectbox("Sync inventory to marketplace (Sync Avail Inv)", ["", "Y", "N"])
disc_mids = []
if not undno:
    note = st.text_input("DNO Note (optional)")
    mids = list(dict.fromkeys(m for m in picked["master_id"] if m))
    if st.checkbox(f"Mark Manufacturer Discontinued on Master Products ({len(mids)} Master ID(s))",
                   key="mark_discontinued", help="Adds each selected Master ID to the Master Products sheet with Manufacturer "
                        "Discontinued? = Y and the Manufacturer Discontinued Date = the DNO start date."):
        disc_mids = mids
        if not mids:
            st.warning("The selected listings have no Master IDs, so nothing will be added to Master Products.")

if st.button("Generate UN-DNO File" if undno else "Generate DNO File", type="primary"):
    cfg = {"undno": undno, "temporary": temp, "reason": reason, "start": start, "end": end,
           "shippable": ship, "sync_inv": sync, "note": note, "discontinue_mids": disc_mids}
    try:
        data = core.generate_dno(io.BytesIO(store["bytes"]), listing_ids, cfg)
        out_name = f"{safe_brand}_{sku_count}_UN-DNO.xlsx" if undno else f"{safe_brand}_{sku_count}_{dno_type} DNO.xlsx"
        st.success(f"{out_name} generated - {len(listing_ids)} listing ID(s)"
                   + (f"; Master Products: {len(disc_mids)} Master ID(s) marked Manufacturer Discontinued" if disc_mids else "")
                   + ". Downloading...")
        # trigger the browser download straight away (no extra download button)
        b64 = base64.b64encode(data).decode()
        components.html(
            f'<a id="dl" download="{out_name}" '
            f'href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"></a>'
            '<script>document.getElementById("dl").click();</script>',
            height=0,
        )
    except ValueError as e:
        st.error(str(e))
