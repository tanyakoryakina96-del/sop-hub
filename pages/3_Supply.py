"""Supply Review dashboard.

Reads through ``data.query`` only — no direct DuckDB access (rule §3.6).
Shares the ``filters`` session key with the other dashboards; channel/region
selections are ignored by the supply queries because ``fact_supply`` is
plant-keyed (see ``_supply_where`` in ``data/query.py``).
"""

import math
import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import config  # noqa: F401, E402 — registers Plotly "tko" template
from data import query  # noqa: E402
from exports import xlsx_builder  # noqa: E402

st.set_page_config(page_title="Supply — S&OP Hub", page_icon="🏭", layout="wide")

_hdr_l, _hdr_r = st.columns([6, 2])
with _hdr_l:
    st.title("🏭 Supply Review")
    st.caption("Inventory, fill rate, production adherence, capacity utilization.")
_export_slot = _hdr_r.empty()

# ---------------------------------------------------------------------------
# Filter bar (shared shape with Demand — SPEC §10.2)
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
        help="Channel filter is ignored on this page (supply is plant-keyed).",
    )
    r1[5].markdown("&nbsp;", unsafe_allow_html=True)
    reset = r1[5].button("Reset", use_container_width=True)

    r2 = st.columns(2)
    regions = r2[0].multiselect(
        "Region", options=reg_opts,
        default=_clamp(current.regions, reg_opts),
        key="flt_regions",
        help="Region filter is ignored on this page (supply is plant-keyed).",
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

sv = query.current_data_version("supply")
if sv is None:
    st.warning(
        "No supply data uploaded yet. Visit the **Upload** page to load a "
        "`supply` CSV before reviewing this dashboard."
    )
    st.stop()

dv_demand = query.current_data_version("demand")
if dv_demand is None:
    st.info(
        "No demand data uploaded — DOS metrics will appear as `—` until a "
        "demand CSV is loaded."
    )

# ---------------------------------------------------------------------------
# KPI tiles
# ---------------------------------------------------------------------------

kpis = query.supply_kpis(f)
fr_trend = query.supply_fill_rate_trend(f)


def _fmt(v: float, fmt: str) -> str:
    return "—" if v is None or math.isnan(v) else format(v, fmt)


k1, k2, k3 = st.columns(3)
k1.metric(
    "Days of Supply (DOS)",
    _fmt(kpis.dos_avg, ".1f") + ("" if math.isnan(kpis.dos_avg) else " days"),
    help=f"Target {config.RAG['dos']['target_days']} days "
         f"± {config.RAG['dos']['green_band']} (green) / "
         f"± {config.RAG['dos']['amber_band']} (amber).",
)
k2.metric(
    "Fill Rate",
    _fmt(kpis.fill_rate, ".1f") + ("" if math.isnan(kpis.fill_rate) else "%"),
    help=f"≥ {config.RAG['fill_rate']['green']}% green; "
         f"≥ {config.RAG['fill_rate']['amber']}% amber.",
)
k3.metric(
    "Production Adherence",
    _fmt(kpis.production_adherence, ".1f")
    + ("" if math.isnan(kpis.production_adherence) else "%"),
)

# Sparkline under the fill rate tile (SPEC §5.2)
if not fr_trend.empty:
    spark = go.Figure()
    spark.add_scatter(
        x=fr_trend["period_date"], y=fr_trend["fill_rate"],
        mode="lines", line=dict(color=config.SERIES["consensus_fcst"], width=2),
        hovertemplate="%{x|%Y-%m}: %{y:.1f}%<extra></extra>",
    )
    spark.update_layout(
        height=80, margin=dict(l=0, r=0, t=0, b=0),
        showlegend=False,
        xaxis=dict(visible=False), yaxis=dict(visible=False),
        paper_bgcolor=config.TKO["colors"]["void"],
        plot_bgcolor=config.TKO["colors"]["void"],
    )
    k2.plotly_chart(spark, use_container_width=True, config={"displayModeBar": False})

st.divider()

# ---------------------------------------------------------------------------
# Top row — Inventory by SKU (left) + Production Adherence (right)
# ---------------------------------------------------------------------------

left, right = st.columns(2)

with left:
    st.subheader("Inventory by SKU (DOS at latest period)")
    dos_df = query.supply_dos_by_sku(f)
    if dos_df.empty or dos_df["dos"].isna().all():
        st.info("No DOS data — needs supply + demand for the latest period.")
    else:
        d = dos_df.dropna(subset=["dos"]).sort_values("dos")
        bar_colors = [config.RAG_COLORS[r] for r in d["rag"]]
        fig = go.Figure(go.Bar(
            x=d["dos"], y=d["sku_name"], orientation="h",
            marker_color=bar_colors,
            hovertemplate="%{y}<br>DOS: %{x:.1f} days<extra></extra>",
        ))
        fig.add_vline(
            x=config.RAG["dos"]["target_days"],
            line=dict(color=config.TKO["colors"]["acid"], width=2, dash="dot"),
            annotation_text=f"Target {config.RAG['dos']['target_days']}d",
            annotation_position="top",
        )
        fig.update_layout(
            height=max(360, 24 * len(d) + 80),
            xaxis_title="Days of Supply",
            margin=dict(l=140, r=24, t=40, b=48),
        )
        st.plotly_chart(fig, use_container_width=True)

with right:
    st.subheader("Production Adherence by Plant")
    adh = query.supply_production_adherence(f)
    if adh.empty:
        st.info("No production data in window.")
    else:
        fig = go.Figure()
        fig.add_bar(
            x=adh["plant_name"], y=adh["plan"], name="Plan",
            marker_color=config.SERIES["budget"],
        )
        fig.add_bar(
            x=adh["plant_name"], y=adh["actual"], name="Actual",
            marker_color=config.SERIES["consensus_fcst"],
        )
        fig.update_layout(
            barmode="group", height=380,
            yaxis_title="Cases",
            legend=dict(orientation="h", y=-0.2, x=0),
        )
        st.plotly_chart(fig, use_container_width=True)

        adh_disp = adh.copy()
        adh_disp["adherence_pct"] = adh_disp["adherence_pct"].round(1)
        st.dataframe(
            adh_disp.style.format({
                "plan": "{:,.0f}", "actual": "{:,.0f}",
                "adherence_pct": "{:.1f}%",
            }),
            use_container_width=True, hide_index=True,
            column_config={
                "plant_code": "Plant", "plant_name": "Name",
                "plan": "Plan (cases)", "actual": "Actual (cases)",
                "adherence_pct": "Adherence",
            },
        )

# ---------------------------------------------------------------------------
# Capacity Utilization (bullet-style)
# ---------------------------------------------------------------------------

st.subheader("Capacity Utilization by Plant")
cap = query.supply_capacity_utilization(f)
if cap.empty:
    st.info("No capacity data in window.")
else:
    fig = go.Figure()
    fig.add_bar(
        x=cap["capacity"], y=cap["plant_name"], orientation="h",
        name="Capacity", marker_color=config.TKO["colors"]["cold_slate"],
        hovertemplate="%{y}<br>Capacity: %{x:,.0f}<extra></extra>",
    )
    fig.add_bar(
        x=cap["actual"], y=cap["plant_name"], orientation="h",
        name="Actual", marker_color=config.SERIES["signal"],
        hovertemplate="%{y}<br>Actual: %{x:,.0f}<extra></extra>",
    )
    fig.update_layout(
        barmode="overlay", height=max(220, 36 * len(cap) + 80),
        xaxis_title="Cases",
        legend=dict(orientation="h", y=-0.25, x=0),
        margin=dict(l=140, r=24, t=24, b=48),
    )
    st.plotly_chart(fig, use_container_width=True)

    cap_disp = cap.copy()
    cap_disp["utilization_pct"] = cap_disp["utilization_pct"].round(1)
    st.dataframe(
        cap_disp.style.format({
            "actual": "{:,.0f}", "capacity": "{:,.0f}",
            "utilization_pct": "{:.1f}%",
        }),
        use_container_width=True, hide_index=True,
        column_config={
            "plant_code": "Plant", "plant_name": "Name",
            "actual": "Actual (cases)", "capacity": "Capacity (cases)",
            "utilization_pct": "Utilization",
        },
    )

# ---------------------------------------------------------------------------
# Inventory Cover Heatmap (SKU × Month)
# ---------------------------------------------------------------------------

st.subheader("Inventory Cover (SKU × Month)")
heat = query.supply_inventory_heatmap(f)
if heat.empty:
    st.info("No inventory data in window — need both supply and demand.")
else:
    pivot = heat.pivot_table(
        index="sku_code", columns="period_date",
        values="dos", aggfunc="mean",
    ).sort_index()
    pivot.columns = [pd.Timestamp(c).strftime("%Y-%m") for c in pivot.columns]

    cfg = config.RAG["dos"]
    tgt = cfg["target_days"]
    # Diverging colorscale centered at target. Acid = on-target,
    # plasma = under (risk), neon_indigo = over (excess).
    colorscale = [
        [0.0, config.RAG_COLORS["red"]],
        [0.5, config.RAG_COLORS["green"]],
        [1.0, config.RAG_COLORS["amber"]],
    ]
    zmin = tgt - cfg["amber_band"] * 2
    zmax = tgt + cfg["amber_band"] * 2
    fig = go.Figure(go.Heatmap(
        z=pivot.values, x=pivot.columns, y=pivot.index,
        colorscale=colorscale, zmin=zmin, zmid=tgt, zmax=zmax,
        hovertemplate="%{y} · %{x}<br>DOS: %{z:.1f} days<extra></extra>",
        colorbar=dict(title="DOS", tickformat=".0f"),
    ))
    fig.update_layout(
        height=max(320, 22 * len(pivot.index) + 80),
        margin=dict(l=80, r=24, t=24, b=48),
    )
    st.plotly_chart(fig, use_container_width=True)

# ---------------------------------------------------------------------------
# Supply Gaps (next 3 months)
# ---------------------------------------------------------------------------

st.subheader("Supply Gaps — Projected Stockouts (next 3 months)")
gaps = query.supply_gaps(f)
if gaps.empty:
    st.success("No projected shortfalls in the next 3 months.")
else:
    gaps_disp = gaps.copy()
    gaps_disp["period_date"] = pd.to_datetime(gaps_disp["period_date"]).dt.strftime("%Y-%m")
    st.dataframe(
        gaps_disp.style.format({"shortfall": "{:,.0f}"}),
        use_container_width=True, hide_index=True,
        column_config={
            "sku_code": "SKU", "sku_name": "Name",
            "period_date": "Period",
            "shortfall": "Shortfall (native UOM)",
        },
    )

st.caption(
    f"Data: supply v{sv}"
    + (f" · demand v{dv_demand}" if dv_demand else "")
    + f" · Window: {f.period_from:%Y-%m} → {f.period_to:%Y-%m}"
)

# ---------------------------------------------------------------------------
# Excel export — supply module (CONTRACTS §5.4 / §6 phase 4)
# ---------------------------------------------------------------------------

with _export_slot.container():
    st.markdown("&nbsp;", unsafe_allow_html=True)
    try:
        _xlsx_bytes, _xlsx_name = xlsx_builder.build(modules=["supply"], filters=f)
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
