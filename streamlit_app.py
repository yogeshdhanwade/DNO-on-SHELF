"""DNO Desk - Streamlit app (uses core.py).

Run:  streamlit run streamlit_app.py
"""
import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st

import core

st.set_page_config(page_title="DNO Desk", page_icon="🗂️", layout="wide")

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ss = st.session_state

CSS = """
<style>
.block-container { padding-top: 1.5rem; max-width: 1100px; }
.hero { background: linear-gradient(120deg, #4f46e5 0%, #0ea5a4 100%); color: #fff;
        border-radius: 16px; padding: 22px 28px; margin-bottom: 14px; }
.hero h1 { margin: 0; font-size: 2rem; color: #fff; padding: 0; }
.hero p { margin: 4px 0 0; opacity: .92; font-size: 1.02rem; }
.steps { display: flex; gap: 10px; margin: 6px 0 18px; flex-wrap: wrap; }
.step { flex: 1; min-width: 150px; padding: 10px 14px; border-radius: 12px; font-weight: 600;
        background: rgba(128,128,128,.12); color: rgba(128,128,128,.95); }
.step.done { background: rgba(16,185,129,.16); color: #059669; }
.step.now { background: rgba(79,70,229,.16); color: #4f46e5; outline: 2px solid rgba(79,70,229,.45); }
.badge { display: inline-block; padding: 3px 12px; border-radius: 999px; font-weight: 700; font-size: .85rem; }
.badge.temp { background: rgba(245,158,11,.2); color: #b45309; }
.badge.perm { background: rgba(239,68,68,.18); color: #b91c1c; }
div[data-testid="stMetric"] { background: rgba(128,128,128,.08); border-radius: 12px; padding: 10px 14px; }
div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {
    background: linear-gradient(120deg, #4f46e5, #0ea5a4); border: 0; color: #fff; font-weight: 600; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


@st.cache_resource
def template_store():
    """Holds the default template once it has been read from the repo."""
    return {}


store = template_store()

# Default template: DNO_Template.xlsx committed next to this file in the repo.
DEFAULT_TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "DNO_Template.xlsx")
if "bytes" not in store and os.path.exists(DEFAULT_TEMPLATE):
    with open(DEFAULT_TEMPLATE, "rb") as f:
        store["bytes"] = f.read()

MODE_LABELS = {"mpn": "MPN", "master_id": "Master ID (MID)", "sku": "SKU",
               "page_id": "Item ID / ASIN", "secondary_id": "GTIN (UPC)"}
MODE_NOUNS = {"mpn": "MPNs", "master_id": "Master IDs", "sku": "SKUs",
              "page_id": "Item IDs / ASINs", "secondary_id": "GTINs (UPCs)"}


# ---------- Sidebar ----------
with st.sidebar:
    st.markdown("### 🗂️ DNO Desk")
    st.markdown(
        "1. **Upload** the brand portfolio\n"
        "2. **Find** listing IDs by MPN, MID, SKU, Item ID / ASIN or GTIN\n"
        "3. **Set** the DNO details and download the file"
    )
    st.divider()
    if st.button("🔄 Start over", use_container_width=True):
        for k in list(ss.keys()):
            del ss[k]
        st.rerun()

# ---------- Header + progress ----------
st.markdown('<div class="hero"><h1>🗂️ DNO Desk</h1>'
            '<p>Portfolio in, listing IDs found, DNO file out.</p></div>', unsafe_allow_html=True)
progress = st.empty()


def render_progress():
    got_portfolio = bool(ss.get("portfolio"))
    got_lids = any(r["listing_id"] for r in ss.get("checkup") or [])
    got_file = bool(ss.get("dno_file"))
    done = [got_portfolio, got_lids, got_file]
    now = done.index(False) if False in done else -1
    labels = ["1 · Portfolio", "2 · Listing IDs", "3 · DNO file"]
    cells = []
    for i, label in enumerate(labels):
        cls = "done" if done[i] else ("now" if i == now else "")
        mark = "✓ " if done[i] else ""
        cells.append(f'<div class="step {cls}">{mark}{label}</div>')
    progress.markdown('<div class="steps">' + "".join(cells) + "</div>", unsafe_allow_html=True)


def main():
    if "bytes" not in store:
        st.error("DNO_Template.xlsx not found in the repository. Add it next to streamlit_app.py on GitHub.")

    # ---------- Upload ----------
    with st.container(border=True):
        st.subheader("📁 Brand portfolio")
        pf = st.file_uploader("Drop the brand portfolio here (.xlsx with Product Listings & optional Master Products)",
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
            st.info("Upload a portfolio file to get started.")
            st.stop()

        brand = state.get("brand") or os.path.splitext(ss.get("pf_key", ("Brand",))[0])[0]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Brand", brand)
        c2.metric("Listings", f"{len(state['prod_df']):,}")
        c3.metric("Marketplaces", len(state["marketplaces"]))
        c4.metric("Sellers", len(state["sellers"]))
        if not state.get("brand"):
            st.caption("Brand taken from the portfolio file name (no brand column found).")

    # ---------- Find listings ----------
    caps = core.capabilities(state)
    with st.container(border=True):
        st.subheader("🔎 Find listing IDs")
        modes = [m for m in core.MODES if caps.get(m, True)]
        mode = st.segmented_control("Look up listings by", modes, default=modes[0],
                                    format_func=MODE_LABELS.get, key="mode")
        if mode is None:  # user clicked the selected option again
            st.info("Pick an option above to continue.")
            st.stop()
        hidden = [MODE_LABELS[m] for m in core.MODES if m not in modes]
        if hidden:
            with st.expander("Some options aren't available for this portfolio"):
                st.write("Column not found for: " + ", ".join(hidden))
                st.caption("Columns in Product Listings: " + ", ".join(map(str, state["prod_df"].columns)))

        scope = st.segmented_control("Scope", ["All Listing IDs", "Specific Listing IDs"],
                                     default="All Listing IDs", key="scope") or "All Listing IDs"
        mps, sellers = [], []
        if scope == "Specific Listing IDs":
            a, b = st.columns(2)
            mps = a.multiselect("Marketplace(s)", state["marketplaces"], key="sel_mps")
            # only sellers that sell in the chosen marketplaces
            seller_opts = sorted({s for mp in mps for s in state["sellers_by_mp"].get(mp, [])})
            ss["sel_sellers"] = [s for s in ss.get("sel_sellers", []) if s in seller_opts]
            sellers = b.multiselect("Seller(s)", seller_opts, key="sel_sellers",
                                    placeholder="Select a marketplace first" if not mps else "Choose seller(s)")
        hint = " - leave blank to use all from Master Products" if mode == "mpn" else ""
        values = st.text_area(f"Paste {MODE_NOUNS[mode]} (comma or one per line){hint}", height=120)

        if st.button("Find listing IDs", type="primary"):
            ss.pop("dno_file", None)
            try:
                ss["checkup"] = core.build_checkup(state, mode, core.parse_values(values),
                                                   scope == "All Listing IDs", mps, sellers)
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
    safe_brand = "".join(c for c in ss["brand"] if c.isalnum() or c in ("_", "-", " ")).strip() or "Brand"

    with st.container(border=True):
        st.subheader("✅ Results")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Listing IDs", f"{len(listing_ids):,}")
        m2.metric("SKUs", f"{sku_count:,}")
        m3.metric("Marketplaces", len({r["marketplace"] for r in found}))
        m4.metric("No match", len(rows) - len(found))
        st.dataframe(pd.DataFrame(rows), use_container_width=True, height=280)
        st.download_button("⬇️ Download checkup (.xlsx)", core.checkup_to_xlsx(rows),
                           file_name=f"Checkup_{safe_brand}_{datetime.now():%Y%m%d_%H%M%S}.xlsx", mime=XLSX)
    if not listing_ids:
        st.stop()

    # ---------- DNO ----------
    with st.container(border=True):
        st.subheader("⚙️ DNO details")
        if "bytes" not in store:
            st.warning("DNO template is missing from the repository.")
            st.stop()

        dno_type = st.segmented_control("DNO type", ["Temporary", "Permanent"],
                                        default="Temporary", key="dno_type") or "Temporary"
        temp = dno_type == "Temporary"
        st.markdown(f'<span class="badge {"temp" if temp else "perm"}">{dno_type.upper()} DNO</span>',
                    unsafe_allow_html=True)
        reason = st.selectbox("Reason code", core.TEMPORARY_REASONS if temp else core.PERMANENT_REASONS)
        d1, d2 = st.columns(2)
        start_d = d1.date_input("Start date", value=None, format="YYYY-MM-DD")
        end_d = d2.date_input("End date", value=None, format="YYYY-MM-DD") if temp else None
        s1, s2 = st.columns(2)
        ship = s1.selectbox("Shippable", ["", "Y", "N"])
        sync = s2.selectbox("Sync avail inv", ["", "Y", "N"])
        note = st.text_input("DNO note (optional)")

        fname = f"{safe_brand}_{sku_count}_{dno_type} DNO.xlsx"
        st.caption(f"File name: `{fname}`")

        if st.button("Generate DNO file", type="primary"):
            cfg = {"temporary": temp, "reason": reason,
                   "start": start_d.isoformat() if start_d else "",
                   "end": end_d.isoformat() if end_d else "",
                   "shippable": ship, "sync_inv": sync, "note": note}
            try:
                ss["dno_file"] = (core.generate_dno(io.BytesIO(store["bytes"]), listing_ids, cfg), fname)
                ss["_just_made"] = True
            except ValueError as e:
                ss.pop("dno_file", None)
                st.error(str(e))

        if ss.get("dno_file"):
            data, name = ss["dno_file"]
            st.success(f"Your DNO file is ready - {len(listing_ids):,} listing ID(s).")
            st.download_button("⬇️ Download DNO file", data, file_name=name, mime=XLSX, type="primary")
            if ss.pop("_just_made", False):
                st.balloons()


try:
    main()
finally:
    render_progress()
