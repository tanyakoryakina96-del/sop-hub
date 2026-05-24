"""S&OP Hub — entry point.

Renders the title, a data-version chip (demand is the S&OP cycle clock —
CONTRACTS §4 / §7.1), and points the user at the sidebar pages. No business
logic lives here.
"""

import streamlit as st

import config  # noqa: F401 — side effect: registers the Plotly "tko" template
from data import session_db

st.set_page_config(
    page_title="S&OP Hub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Per-session DuckDB: each visitor gets their own file, pre-seeded with
# anonymized sample data so dashboards render on first paint. See CONTRACTS §5.6.
session_db.cleanup_stale_sessions()
session_db.seed_if_empty()

# ---------------------------------------------------------------------------
# Header — title + demand-version chip
# ---------------------------------------------------------------------------

versions = st.session_state.get("data_versions", {})
demand_v = versions.get("demand")
chip_text = f"Demand v{demand_v}" if demand_v else "No data yet"
chip_bg = config.TKO["colors"]["plasma"] if demand_v else config.TKO["colors"]["cold_slate"]

left, right = st.columns([6, 2])
with left:
    st.title("S&OP Hub")
    st.caption("Self-service S&OP reporting — Demand, Supply, Financial, Scenario")
with right:
    st.markdown(
        f"""
        <div style="text-align:right; padding-top: 1.2rem;">
          <span style="
            background:{chip_bg};
            color:{config.TKO['colors']['ice_white']};
            padding: 0.3rem 0.75rem;
            border-radius: 999px;
            font-size: 0.85rem;
            font-weight: 600;
            letter-spacing: 0.02em;
          ">{chip_text}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.divider()

# ---------------------------------------------------------------------------
# Landing copy — points to the sidebar pages
# ---------------------------------------------------------------------------

st.markdown(
    """
### Get started

1. **Upload** your CSVs — start with `sku_master` and `demand` (the demand
   upload is what anchors the dashboard's default date range).
2. **Review** each S&OP gate: Demand → Supply → Financial → Scorecard.
3. **Run scenarios** on the Scenario Planner and freeze the Exec Summary view
   for the Exec S&OP meeting.
4. **Export** any page (or the full bundled deck from Exec Summary) as PPTX
   or Excel — one click.

Use the sidebar to navigate. All pages share the same filter bar; filter state
persists within the session.
    """
)

# Domain-by-domain summary chips
if versions:
    st.subheader("Active data versions")
    cols = st.columns(4)
    for i, domain in enumerate(("demand", "supply", "financial", "sku_master")):
        with cols[i]:
            v = versions.get(domain)
            st.metric(label=domain, value=f"v{v}" if v else "—")
