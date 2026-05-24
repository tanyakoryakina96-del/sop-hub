"""Demand Review dashboard.

Reads through `data.query` only — no direct DuckDB access (rule §3.6).
Constructs a `Filters` from widget values and writes it to
`st.session_state["filters"]` so other pages inherit the same window.
"""

import sys
from pathlib import Path

import plotly.graph_objects as go
import streamlit as st

# Allow imports from the project root when Streamlit loads this file directly.
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import config  # noqa: F401, E402 — registers Plotly "tko" template
from data import query  # noqa: E402
from exports import xlsx_builder  # noqa: E402

st.set_page_config(page_title="Demand — S&OP Hub", page_icon="📈", layout="wide")

_hdr_l, _hdr_r = st.columns([6, 2])
with _hdr_l:
    st.title("📈 Demand Review")
    st.caption("Forecast accuracy, bias, volume bridge, and channel mix.")
_export_slot = _hdr_r.empty()  # filled below once filters/data are resolved

# ---------------------------------------------------------------------------
# Filter bar (SPEC §10.2)
# ---------------------------------------------------------------------------

if "filters" not in st.session_state:
    st.session_state["filters"] = query.default_filters()

current: query.Filters = st.session_state["filters"]

# Pull dim options once per render (cached)
brand_opts = query.list_brands()
cat_opts = query.list_categories()
chan_opts = query.list_channels()
reg_opts = query.list_regions()
sku_df = query.list_skus()
sku_opts = sku_df["sku_code"].tolist() if not sku_df.empty else []
sku_label = {
    row.sku_code: f"{row.sku_code} — {row.sku_name}"
    for row in sku_df.itertuples()
}

# Clamp persisted selections to current option set (a prior upload may have
# removed a brand/sku that's still in session_state).
def _clamp(values: list[str], opts: list[str]) -> list[str]:
    return [v for v in values if v in opts]


with st.expander("🔍 Filters", expanded=True):
    r1 = st.columns([2, 2, 2, 2, 2, 1])
    period_from = r1[0].date_input(
        "Period from", value=current.period_from, key="flt_period_from"
    )
    period_to = r1[1].date_input(
        "Period to", value=current.period_to, key="flt_period_to"
    )
    brands = r1[2].multiselect(
        "Brand", options=brand_opts,
        default=_clamp(current.brands, brand_opts),
        key="flt_brands",
    )
    categories = r1[3].multiselect(
        "Category", options=cat_opts,
        default=_clamp(current.categories, cat_opts),
        key="flt_categories",
    )
    channels = r1[4].multiselect(
        "Channel", options=chan_opts,
        default=_clamp(current.channels, chan_opts),
        key="flt_channels",
    )
    r1[5].markdown("&nbsp;", unsafe_allow_html=True)
    reset = r1[5].button("Reset", use_container_width=True)

    r2 = st.columns(2)
    regions = r2[0].multiselect(
        "Region", options=reg_opts,
        default=_clamp(current.regions, reg_opts),
        key="flt_regions",
    )
    skus = r2[1].multiselect(
        "SKU", options=sku_opts,
        default=_clamp(current.skus, sku_opts),
        format_func=lambda c: sku_label.get(c, c),
        key="flt_skus",
    )

if reset:
    st.session_state["filters"] = query.default_filters()
    for k in ("flt_period_from", "flt_period_to", "flt_brands",
              "flt_categories", "flt_channels", "flt_regions", "flt_skus"):
        st.session_state.pop(k, None)
    st.rerun()

# Assemble the canonical Filters from widget values and store.
f = query.Filters(
    period_from=period_from,
    period_to=period_to,
    brands=brands,
    categories=categories,
    channels=channels,
    regions=regions,
    skus=skus,
)
st.session_state["filters"] = f

# ---------------------------------------------------------------------------
# Data gate
# ---------------------------------------------------------------------------

dv = query.current_data_version("demand")
if dv is None:
    st.warning(
        "No demand data uploaded yet. Visit the **Upload** page to load a "
        "`demand` CSV before reviewing this dashboard."
    )
    st.stop()

# ---------------------------------------------------------------------------
# KPI row
# ---------------------------------------------------------------------------

kpis = query.demand_kpis(f)

import math  # noqa: E402 — used for the empty-window NaN check below


def _fmt(v: float, fmt: str) -> str:
    return "—" if v is None or math.isnan(v) else format(v, fmt)


k1, k2, k3, k4 = st.columns(4)
k1.metric("Forecast Accuracy", _fmt(kpis.fa_pct, ".1f") + ("" if math.isnan(kpis.fa_pct) else "%"))
k2.metric("MAPE",              _fmt(kpis.mape,   ".1f") + ("" if math.isnan(kpis.mape)   else "%"))
k3.metric("Bias",              _fmt(kpis.bias,   "+.1f") + ("" if math.isnan(kpis.bias)   else "%"))
k4.metric("Volume (cases)",    _fmt(kpis.volume_total, ",.0f"))

st.divider()

# ---------------------------------------------------------------------------
# Charts — two-column grid
# ---------------------------------------------------------------------------

left, right = st.columns(2)

with left:
    st.subheader("Volume Bridge")
    bridge = query.demand_volume_bridge(f)
    if bridge.empty or bridge["value"].abs().sum() == 0:
        st.info("No volume in window.")
    else:
        fig = go.Figure(go.Waterfall(
            x=bridge["stage"],
            y=bridge["value"],
            measure=["absolute", "relative", "relative", "total"],
            text=[f"{v:,.0f}" for v in bridge["value"]],
            textposition="outside",
            connector={"line": {"color": config.TKO["colors"]["cold_slate"]}},
            increasing={"marker": {"color": config.SERIES["consensus_fcst"]}},
            decreasing={"marker": {"color": config.SERIES["statistical_fcst"]}},
            totals={"marker": {"color": config.SERIES["signal"]}},
        ))
        fig.update_layout(height=380, showlegend=False,
                          yaxis_title="Cases")
        st.plotly_chart(fig, use_container_width=True)

with right:
    st.subheader("Forecast vs Actuals")
    series = query.demand_fcst_vs_actuals(f)
    if series.empty:
        st.info("No data in window.")
    else:
        fig = go.Figure()
        fig.add_bar(
            x=series["period_date"], y=series["actuals"], name="Actuals",
            marker_color=config.SERIES["actuals"], opacity=0.85,
        )
        fig.add_scatter(
            x=series["period_date"], y=series["consensus_fcst"],
            name="Consensus Fcst", mode="lines+markers",
            line=dict(color=config.SERIES["consensus_fcst"], width=3),
        )
        fig.add_scatter(
            x=series["period_date"], y=series["statistical_fcst"],
            name="Statistical Fcst", mode="lines+markers",
            line=dict(color=config.SERIES["statistical_fcst"],
                      width=2, dash="dash"),
        )
        fig.update_layout(
            height=380, hovermode="x unified",
            yaxis_title="Cases",
            legend=dict(orientation="h", y=-0.2, x=0),
        )
        st.plotly_chart(fig, use_container_width=True)

# ---------------------------------------------------------------------------
# Channel mix (full-width)
# ---------------------------------------------------------------------------

st.subheader("Channel Mix")
mix = query.demand_channel_mix(f)
if mix.empty:
    st.info("No channel data in window.")
else:
    pivot = mix.pivot_table(
        index="period_date", columns="channel_name",
        values="volume_share", aggfunc="sum",
    ).fillna(0).sort_index()
    fig = go.Figure()
    for col in pivot.columns:
        fig.add_bar(x=pivot.index, y=pivot[col], name=col)
    fig.update_layout(
        barmode="stack", height=380,
        yaxis=dict(title="Share (%)", range=[0, 100]),
        legend=dict(orientation="h", y=-0.2, x=0),
    )
    st.plotly_chart(fig, use_container_width=True)

# ---------------------------------------------------------------------------
# MAPE by SKU table
# ---------------------------------------------------------------------------

st.subheader("Forecast Accuracy by SKU")
sku_acc = query.demand_mape_by_sku(f)
if sku_acc.empty:
    st.info("No SKU-level data in window.")
else:
    rag_colors = config.RAG_COLORS
    void = config.TKO["colors"]["void"]

    def _rag_style(val: str) -> str:
        bg = rag_colors.get(val, "")
        return f"background-color: {bg}; color: {void}; font-weight: 600;"

    styled = (
        sku_acc.style
        .map(_rag_style, subset=["rag"])
        .format({
            "mape":   "{:.1f}",
            "fa_pct": "{:.1f}",
            "bias":   "{:+.1f}",
            "volume": "{:,.0f}",
        })
    )
    st.dataframe(
        styled, use_container_width=True, hide_index=True,
        column_config={
            "sku_code": "SKU",
            "sku_name": "Name",
            "brand":    "Brand",
            "mape":     "MAPE %",
            "fa_pct":   "FA %",
            "bias":     "Bias %",
            "volume":   "Volume (cases)",
            "rag":      "RAG",
        },
    )

st.caption(
    f"Data: demand v{dv} · Window: {f.period_from:%Y-%m} → {f.period_to:%Y-%m}"
)

# ---------------------------------------------------------------------------
# Excel export (Phase 3 — demand only; CONTRACTS §5.4 / §6)
# ---------------------------------------------------------------------------

with _export_slot.container():
    st.markdown("&nbsp;", unsafe_allow_html=True)  # vertical nudge to align w/ title
    try:
        _xlsx_bytes, _xlsx_name = xlsx_builder.build(modules=["demand"], filters=f)
    except Exception as e:  # noqa: BLE001 — surface to user, don't crash the page
        st.error(f"Excel export failed: {e}")
    else:
        st.download_button(
            label="⬇ Export to Excel",
            data=_xlsx_bytes,
            file_name=_xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
