"""Financial Review dashboard.

Reads through ``data.query`` only (rule §3.6). Shares the ``filters`` session
key with the other dashboards; region selections are ignored by the financial
queries because ``fact_financial`` has no region column.
"""

import math
import sys
from pathlib import Path

import plotly.graph_objects as go
import streamlit as st

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import config  # noqa: F401, E402 — registers Plotly "tko" template
from data import query  # noqa: E402
from exports import xlsx_builder  # noqa: E402

st.set_page_config(page_title="Financial — S&OP Hub", page_icon="💰", layout="wide")

_hdr_l, _hdr_r = st.columns([6, 2])
with _hdr_l:
    st.title("💰 Financial Review")
    st.caption("Revenue vs Budget & LE, GM%, channel mix, brand P&L.")
_export_slot = _hdr_r.empty()

# ---------------------------------------------------------------------------
# Filter bar (shared shape)
# ---------------------------------------------------------------------------

if "filters" not in st.session_state:
    st.session_state["filters"] = query.default_filters()

current: query.Filters = st.session_state["filters"]

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
        help="Region filter is ignored on this page (financial data has no region).",
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

f = query.Filters(
    period_from=period_from, period_to=period_to,
    brands=brands, categories=categories, channels=channels,
    regions=regions, skus=skus,
)
st.session_state["filters"] = f

# ---------------------------------------------------------------------------
# Data gate
# ---------------------------------------------------------------------------

fv = query.current_data_version("financial")
if fv is None:
    st.warning(
        "No financial data uploaded yet. Visit the **Upload** page to load a "
        "`financial` CSV before reviewing this dashboard."
    )
    st.stop()

# ---------------------------------------------------------------------------
# KPI tiles
# ---------------------------------------------------------------------------

kpis = query.financial_kpis(f)
gm_trend = query.financial_gm_trend(f)


def _fmt(v: float, fmt: str) -> str:
    return "—" if v is None or math.isnan(v) else format(v, fmt)


k1, k2, k3, k4 = st.columns(4)
k1.metric(
    f"Revenue Actual ({config.CURRENCY})",
    _fmt(kpis.revenue_actual, ",.0f"),
)
k2.metric(
    "Revenue vs Budget",
    _fmt(kpis.revenue_vs_budget_pct, "+.1f")
    + ("" if math.isnan(kpis.revenue_vs_budget_pct) else "%"),
    help=f"≥ {config.RAG['rev_vs_budget']['green']}% green; "
         f"≥ {config.RAG['rev_vs_budget']['amber']}% amber.",
)
k3.metric(
    "Revenue vs LE",
    _fmt(kpis.revenue_vs_le_pct, "+.1f")
    + ("" if math.isnan(kpis.revenue_vs_le_pct) else "%"),
)
k4.metric(
    "GM %",
    _fmt(kpis.gm_pct, ".1f") + ("" if math.isnan(kpis.gm_pct) else "%"),
)

# Sparkline under the GM% tile (SPEC §5.3)
if not gm_trend.empty:
    spark = go.Figure()
    spark.add_scatter(
        x=gm_trend["period_date"], y=gm_trend["gm_pct"],
        mode="lines", line=dict(color=config.SERIES["signal"], width=2),
        hovertemplate="%{x|%Y-%m}: %{y:.1f}%<extra></extra>",
    )
    spark.update_layout(
        height=80, margin=dict(l=0, r=0, t=0, b=0),
        showlegend=False,
        xaxis=dict(visible=False), yaxis=dict(visible=False),
        paper_bgcolor=config.TKO["colors"]["void"],
        plot_bgcolor=config.TKO["colors"]["void"],
    )
    k4.plotly_chart(spark, use_container_width=True, config={"displayModeBar": False})

st.divider()

# ---------------------------------------------------------------------------
# Revenue Waterfall (Budget → LE → Actuals) + Channel mix
# ---------------------------------------------------------------------------

left, right = st.columns(2)

with left:
    st.subheader("Revenue Waterfall (Budget → LE → Actuals)")
    wf = query.financial_revenue_waterfall(f)
    if wf.empty or wf["value"].abs().sum() == 0:
        st.info("No revenue in window.")
    else:
        budget = float(wf.loc[wf["stage"] == "Budget", "value"].iloc[0])
        le     = float(wf.loc[wf["stage"] == "LE",     "value"].iloc[0])
        actual = float(wf.loc[wf["stage"] == "Actuals", "value"].iloc[0])
        # Waterfall: Budget (absolute) → ΔLE → ΔActual → Actual total
        fig = go.Figure(go.Waterfall(
            x=["Budget", "Δ to LE", "Δ to Actual", "Actual"],
            y=[budget, le - budget, actual - le, actual],
            measure=["absolute", "relative", "relative", "total"],
            text=[
                f"{budget:,.0f}",
                f"{(le - budget):+,.0f}",
                f"{(actual - le):+,.0f}",
                f"{actual:,.0f}",
            ],
            textposition="outside",
            connector={"line": {"color": config.TKO["colors"]["cold_slate"]}},
            increasing={"marker": {"color": config.SERIES["consensus_fcst"]}},
            decreasing={"marker": {"color": config.SERIES["statistical_fcst"]}},
            totals={"marker": {"color": config.SERIES["signal"]}},
        ))
        fig.update_layout(
            height=380, showlegend=False,
            yaxis_title=f"Revenue ({config.CURRENCY})",
        )
        st.plotly_chart(fig, use_container_width=True)

with right:
    st.subheader("Revenue by Channel")
    chan = query.financial_by_channel(f)
    if chan.empty or chan["revenue_actual"].abs().sum() == 0:
        st.info("No channel revenue in window.")
    else:
        fig = go.Figure(go.Treemap(
            labels=chan["channel_name"],
            parents=[""] * len(chan),
            values=chan["revenue_actual"],
            marker=dict(
                colors=chan["revenue_actual"],
                colorscale=[
                    [0.0, config.TKO["colors"]["graphite"]],
                    [1.0, config.TKO["colors"]["plasma"]],
                ],
            ),
            texttemplate="<b>%{label}</b><br>%{value:,.0f}<br>%{percentRoot:.1%}",
            hovertemplate=(
                "<b>%{label}</b><br>"
                + f"Revenue ({config.CURRENCY}): %{{value:,.0f}}<br>"
                + "Share: %{percentRoot:.1%}<extra></extra>"
            ),
        ))
        fig.update_layout(
            height=380, margin=dict(l=8, r=8, t=24, b=8),
        )
        st.plotly_chart(fig, use_container_width=True)

# ---------------------------------------------------------------------------
# P&L Summary by Brand
# ---------------------------------------------------------------------------

st.subheader("P&L Summary by Brand")
pnl = query.financial_pnl_summary(f)
if pnl.empty:
    st.info("No P&L data in window.")
else:
    fmt_cur = f"{{:,.0f}} {config.CURRENCY}"
    st.dataframe(
        pnl.style.format({
            "revenue":     "{:,.0f}",
            "gm":          "{:,.0f}",
            "promo_spend": "{:,.0f}",
            "net_revenue": "{:,.0f}",
        }),
        use_container_width=True, hide_index=True,
        column_config={
            "brand":       "Brand",
            "revenue":     f"Revenue ({config.CURRENCY})",
            "gm":          f"GM ({config.CURRENCY})",
            "promo_spend": f"Promo Spend ({config.CURRENCY})",
            "net_revenue": f"Net Revenue ({config.CURRENCY})",
        },
    )

# ---------------------------------------------------------------------------
# YTD vs Full-Year Budget (one bar per brand)
# ---------------------------------------------------------------------------

st.subheader(f"YTD Revenue vs Full-Year Budget ({f.period_to.year})")
ytd = query.financial_ytd_progress(f)
if ytd.empty or (ytd["full_year_budget"].fillna(0) == 0).all():
    st.info("No budget data for the selected fiscal year.")
else:
    ytd_plot = ytd[ytd["full_year_budget"] > 0].copy()
    fig = go.Figure()
    fig.add_bar(
        x=ytd_plot["full_year_budget"], y=ytd_plot["brand"], orientation="h",
        marker_color=config.TKO["colors"]["cold_slate"],
        name="Full-Year Budget",
        hovertemplate="%{y}<br>Budget: %{x:,.0f}<extra></extra>",
    )
    fig.add_bar(
        x=ytd_plot["ytd_actual"], y=ytd_plot["brand"], orientation="h",
        marker_color=config.SERIES["signal"],
        name="YTD Actual",
        hovertemplate="%{y}<br>YTD Actual: %{x:,.0f}<extra></extra>",
    )
    fig.update_layout(
        barmode="overlay", height=max(220, 36 * len(ytd_plot) + 80),
        xaxis_title=f"Revenue ({config.CURRENCY})",
        legend=dict(orientation="h", y=-0.25, x=0),
        margin=dict(l=140, r=24, t=24, b=48),
    )
    st.plotly_chart(fig, use_container_width=True)

    ytd_disp = ytd.copy()
    st.dataframe(
        ytd_disp.style.format({
            "ytd_actual":       "{:,.0f}",
            "full_year_budget": "{:,.0f}",
            "ytd_pct":          "{:.1f}%",
        }),
        use_container_width=True, hide_index=True,
        column_config={
            "brand":             "Brand",
            "ytd_actual":        f"YTD Actual ({config.CURRENCY})",
            "full_year_budget":  f"Full-Year Budget ({config.CURRENCY})",
            "ytd_pct":           "YTD %",
        },
    )

st.caption(
    f"Data: financial v{fv} · "
    f"Window: {f.period_from:%Y-%m} → {f.period_to:%Y-%m} · "
    f"Currency: {config.CURRENCY}"
)

# ---------------------------------------------------------------------------
# Excel export — financial module
# ---------------------------------------------------------------------------

with _export_slot.container():
    st.markdown("&nbsp;", unsafe_allow_html=True)
    try:
        _xlsx_bytes, _xlsx_name = xlsx_builder.build(
            modules=["financial"], filters=f,
        )
    except Exception as e:  # noqa: BLE001
        st.error(f"Excel export failed: {e}")
    else:
        st.download_button(
            label="⬇ Export to Excel",
            data=_xlsx_bytes,
            file_name=_xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
