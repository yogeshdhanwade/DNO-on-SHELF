"""Shelf Auto-DNO - Streamlit app (uses core.py).

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

st.set_page_config(page_title="Shelf Auto-DNO", layout="wide")
st.title("Shelf Auto-DNO")
st.caption("Bulk DNO & un-DNO for Shelf LIDs")


@st.cache_resource
def template_store():
    """Shared across sessions while the app process is alive."""
    return {}


store = template_store()
ss = st.session_state

# Default template: any .xlsx with "template" in its name, committed next to this file in the repo.
APP_DIR = os.path.dirname(os.path.abspath(__file__))
if "bytes" not in store:
    candidates = sorted(n for n in os.listdir(APP_DIR)
                        if n.lower().endswith(".xlsx") and "template" in n.lower())
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
pf = st.file_uploader("Brand portfolio (.xlsx with Product Listings & optional Master Products)",
                      type=["xlsx", "xlsm"], key="pf")
if pf is not None:
    key = (pf.name, pf.size)
    if ss.get("pf_key") != key:
        try:
            with st.spinner("Reading portfolio..."):
                ss["portfolio"] = core.load_portfolio(pf.getvalue())
            ss["pf_key"] = key
            ss.pop("checkup", None)
            ss.pop("dno_file", None)
        except Exception as e:
            ss.pop("portfolio", None)
            ss.pop("pf_key", None)
            st.error(str(e))

state = ss.get("portfolio")
if not state:
    st.info("Upload a portfolio file to continue.")
    st.stop()

caps = core.capabilities(state)
brand = state.get("brand") or os.path.splitext(ss.get("pf_key", ("Brand",))[0])[0]
st.caption(f"{len(state['prod_df'])} listing rows - {len(state['marketplaces'])} marketplace(s) - "
           f"{len(state['sellers'])} seller(s)")

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
    st.caption("Columns found in Product Listings: " + ", ".join(map(str, state["prod_df"].columns)))
st.caption(f"Brand: **{brand}**" + ("" if state.get("brand") else " (taken from the portfolio file name; no brand column found)"))
ALL_LIDS, MP_LIDS = "All LIDs", "LIDs Marketplace-wise"
scope = st.segmented_control("Listing Coverage", [ALL_LIDS, MP_LIDS],
                             default=ALL_LIDS, key="scope") or ALL_LIDS
mps, sellers = [], []
if scope == MP_LIDS:
    c1, c2 = st.columns(2)
    mps = c1.multiselect("Marketplace(s)", state["marketplaces"], key="sel_mps")
    # only sellers that sell in the chosen marketplaces
    seller_opts = sorted({s for mp in mps for s in state["sellers_by_mp"].get(mp, [])})
    ss["sel_sellers"] = [s for s in ss.get("sel_sellers", []) if s in seller_opts]
    sellers = c2.multiselect("Seller(s)", seller_opts, key="sel_sellers",
                             placeholder="Select a marketplace first" if not mps else "Choose seller(s)")
hint = " - leave blank to use all from Master Products" if mode == "mpn" else ""
values = st.text_area(f"Paste {MODE_NOUNS[mode]} (comma or one per line){hint}", height=120)

if st.button("Fetch LIDs", type="primary"):
    ss.pop("dno_file", None)
    try:
        rows = core.build_checkup(state, mode, core.parse_values(values),
                                  scope == ALL_LIDS, mps, sellers)
        ss["checkup"] = rows
        ss["brand"] = brand
    except ValueError as e:
        ss.pop("checkup", None)
        st.error(str(e))

rows = ss.get("checkup")
if not rows:
    st.stop()

found = [r for r in rows if r["listing_id"]]
listing_ids = list(dict.fromkeys(r["listing_id"] for r in found))
sku_count = len({r["sku"] for r in found if r["sku"]})
st.write(f"**{len(listing_ids)}** unique listing ID(s); {len(rows) - len(found)} input(s) with no match.")
st.dataframe(pd.DataFrame(rows), use_container_width=True, height=300)
safe_brand = "".join(c for c in ss["brand"] if c.isalnum() or c in ("_", "-", " ")).strip() or "Brand"
st.download_button(
    "Download Checkup (.xlsx)", core.checkup_to_xlsx(rows),
    file_name=f"Checkup_{safe_brand}_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
if not listing_ids:
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
if not undno:
    note = st.text_input("DNO Note (optional)")

if st.button("Generate UN-DNO File" if undno else "Generate DNO File", type="primary"):
    cfg = {"undno": undno, "temporary": temp, "reason": reason, "start": start, "end": end,
           "shippable": ship, "sync_inv": sync, "note": note}
    try:
        data = core.generate_dno(io.BytesIO(store["bytes"]), listing_ids, cfg)
        out_name = f"{safe_brand}_{sku_count}_UN-DNO.xlsx" if undno else f"{safe_brand}_{sku_count}_{dno_type} DNO.xlsx"
        st.success(f"{out_name} generated - {len(listing_ids)} listing ID(s). Downloading...")
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
