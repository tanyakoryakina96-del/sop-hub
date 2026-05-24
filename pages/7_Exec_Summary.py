"""Executive Summary — single-page view for the Exec S&OP meeting (SPEC §5.6).

Curated: 4 KPI cards (FA, Fill Rate, Revenue vs Budget, GM%) + volume trend
+ revenue-vs-budget chart + 3-scenario comparison + bundled-deck export.
Filter freeze toggle locks the current scope; the freeze is an honor-system
convenience, not a security boundary (single-user app, SPEC §0 Rule 6).
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
from exports import pptx_builder, xlsx_builder  # noqa: E402

st.set_page_config(page_title="Exec Summary — S&OP Hub", page_icon="📈", layout="wide")

_hdr_l, _hdr_r = st.columns([6, 2])
with _hdr_l:
    st.title("📈 Executive Summary")
    st.caption("Projector-ready snapshot for the Exec S&OP gate (SPEC §5.6).")
_export_slot = _hdr_r.empty()

# ---------------------------------------------------------------------------
# Filter freeze toggle + filter bar
# ---------------------------------------------------------------------------

if "filters" not in st.session_state:
    st.session_state["filters"] = query.default_filters()
if "exec_filters_frozen" not in st.session_state:
    st.session_state["exec_filters_frozen"] = False
if "scenario_adjustments" not in st.session_state:
    st.session_state["scenario_adjustments"] = {
        "base": 0.0, "upside": 5.0, "downside": -5.0,
    }

current: query.Filters = st.session_state["filters"]
frozen: bool = bool(st.session_state["exec_filters_frozen"])

freeze_col, _ = st.columns([1, 5])
if frozen:
    if freeze_col.button("🔓 Unfreeze filters", use_container_width=True):
        st.session_state["exec_filters_frozen"] = False
        st.rerun()
else:
    if freeze_col.button("🔒 Freeze filters for Exec", use_container_width=True):
        st.session_state["exec_filters_frozen"] = True
        st.rerun()


def _filter_label(values: list[str]) -> str:
    return ", ".join(values) if values else "All"


if frozen:
    st.success(
        "**Filters frozen for Exec review.** "
        f"Window {current.period_from:%Y-%m} → {current.period_to:%Y-%m} · "
        f"Brands: {_filter_label(current.brands)} · "
        f"Categories: {_filter_label(current.categories)} · "
        f"Channels: {_filter_label(current.channels)} · "
        f"Regions: {_filter_label(current.regions)} · "
        f"SKUs: {_filter_label(current.skus)}"
    )
    f = current
else:
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
            default=_clamp(current.brands, brand_opts), key="flt_brands",
        )
        categories = r1[3].multiselect(
            "Category", options=cat_opts,
            default=_clamp(current.categories, cat_opts), key="flt_categories",
        )
        channels = r1[4].multiselect(
            "Channel", options=chan_opts,
            default=_clamp(current.channels, chan_opts), key="flt_channels",
        )
        r1[5].markdown("&nbsp;", unsafe_allow_html=True)
        reset = r1[5].button("Reset", use_container_width=True)

        r2 = st.columns(2)
        regions = r2[0].multiselect(
            "Region", options=reg_opts,
            default=_clamp(current.regions, reg_opts), key="flt_regions",
        )
        skus = r2[1].multiselect(
            "SKU", options=sku_opts,
            default=_clamp(current.skus, sku_opts),
            format_func=lambda c: sku_label.get(c, c), key="flt_skus",
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

versions = {
    d: query.current_data_version(d) for d in ("demand", "supply", "financial")
}
if not any(versions.values()):
    st.warning(
        "No data uploaded yet. Visit the **Upload** page to load demand, "
        "supply, or financial CSVs before running the Exec Summary."
    )
    st.stop()

# ---------------------------------------------------------------------------
# 4 KPI cards — projector-friendly, large numbers
# ---------------------------------------------------------------------------

dem = query.demand_kpis(f)
sup = query.supply_kpis(f)
fin = query.financial_kpis(f)


def _fmt_pct(v: float | None, *, sign: bool = False) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "—"
    return f"{v:+.1f}%" if sign else f"{v:.1f}%"


k1, k2, k3, k4 = st.columns(4)
k1.metric("Forecast Accuracy", _fmt_pct(dem.fa_pct))
k2.metric("Fill Rate",         _fmt_pct(sup.fill_rate))
k3.metric(
    "Revenue vs Budget",
    _fmt_pct(fin.revenue_vs_budget_pct, sign=True),
    help=f"Actual revenue: {fin.revenue_actual:,.0f} {config.CURRENCY}"
         if not math.isnan(fin.revenue_actual) else None,
)
k4.metric("GM %", _fmt_pct(fin.gm_pct))

st.markdown("---")

# ---------------------------------------------------------------------------
# Charts — volume trend + revenue waterfall
# ---------------------------------------------------------------------------

left, right = st.columns(2)

with left:
    st.markdown("##### Volume — Forecast vs Actuals")
    series = query.demand_fcst_vs_actuals(f)
    if series.empty:
        st.info("No demand data in window.")
    else:
        fig = go.Figure()
        fig.add_bar(
            x=series["period_date"], y=series["actuals"], name="Actuals",
            marker_color=config.SERIES["actuals"], opacity=0.85,
        )
        fig.add_scatter(
            x=series["period_date"], y=series["consensus_fcst"],
            name="Consensus", mode="lines+markers",
            line=dict(color=config.SERIES["consensus_fcst"], width=3),
        )
        fig.update_layout(
            hovermode="x unified", yaxis_title="Cases",
            legend=dict(orientation="h", y=-0.18, x=0),
            margin=dict(l=48, r=20, t=24, b=60),
            height=380,
        )
        st.plotly_chart(fig, use_container_width=True)

with right:
    st.markdown("##### Revenue — Budget → LE → Actuals")
    wf = query.financial_revenue_waterfall(f)
    if wf.empty or wf["value"].abs().sum() == 0:
        st.info("No financial data in window.")
    else:
        budget = float(wf.loc[wf["stage"] == "Budget", "value"].iloc[0])
        le     = float(wf.loc[wf["stage"] == "LE",     "value"].iloc[0])
        actual = float(wf.loc[wf["stage"] == "Actuals", "value"].iloc[0])
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
            showlegend=False,
            yaxis_title=f"Revenue ({config.CURRENCY})",
            margin=dict(l=64, r=20, t=24, b=40),
            height=380,
        )
        st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# ---------------------------------------------------------------------------
# 3-scenario comparison
# ---------------------------------------------------------------------------

st.markdown("##### Scenario Comparison")

adj: dict = st.session_state["scenario_adjustments"]
st.caption(
    f"Inputs (set on Scenario page): "
    f"Base {adj['base']:+.1f}% · "
    f"Upside {adj['upside']:+.1f}% · "
    f"Downside {adj['downside']:+.1f}%"
)

baseline = query.scenario_baseline(f)
price = query.scenario_price_proxy(f) if versions["financial"] else pd.DataFrame(
    columns=["sku_code", "channel_code", "avg_price"]
)

if baseline.empty:
    st.info(
        "No forward planning horizon in this window — scenarios have no "
        "months to project."
    )
else:
    merged = baseline.merge(price, on=["sku_code", "channel_code"], how="left")
    rows = []
    for name in ("Base", "Upside", "Downside"):
        pct = float(adj.get(name.lower(), 0.0))
        vol = merged["consensus_volume_cases"] * (1.0 + pct / 100.0)
        rev = (vol * merged["avg_price"]).fillna(0.0)
        rows.append({
            "Scenario": name,
            "% vs Consensus": pct,
            "Volume (cases)": float(vol.sum()),
            f"Revenue ({config.CURRENCY})": float(rev.sum()),
        })
    summary = pd.DataFrame(rows)
    base_rev = float(summary.loc[summary["Scenario"] == "Base",
                                 f"Revenue ({config.CURRENCY})"].iloc[0])
    summary["Δ vs Base %"] = summary[f"Revenue ({config.CURRENCY})"].apply(
        lambda v: (v - base_rev) / base_rev * 100 if base_rev else float("nan")
    )

    rag_colors = config.RAG_COLORS
    void = config.TKO["colors"]["void"]

    def _scen_style(row: pd.Series) -> list[str]:
        styles = [""] * len(row)
        if row["Scenario"] == "Upside":
            styles[0] = f"background-color: {rag_colors['green']}; color: {void}; font-weight: 700;"
        elif row["Scenario"] == "Downside":
            styles[0] = f"background-color: {rag_colors['red']}; color: {void}; font-weight: 700;"
        else:
            styles[0] = "font-weight: 700;"
        return styles

    st.dataframe(
        summary.style.apply(_scen_style, axis=1).format({
            "% vs Consensus": "{:+.1f}%",
            "Volume (cases)": "{:,.0f}",
            f"Revenue ({config.CURRENCY})": "{:,.0f}",
            "Δ vs Base %": "{:+.1f}%",
        }),
        use_container_width=True, hide_index=True,
    )

# ---------------------------------------------------------------------------
# Footer — versions + freeze status
# ---------------------------------------------------------------------------

st.caption(
    "Data versions: "
    + " · ".join(
        f"{d} v{v}" if v else f"{d} —"
        for d, v in versions.items()
    )
    + f" · Window: {f.period_from:%Y-%m} → {f.period_to:%Y-%m}"
    + (" · 🔒 Filters frozen" if frozen else "")
)

# ---------------------------------------------------------------------------
# Export full report — bundled deck with every module that has data
# ---------------------------------------------------------------------------

_bundle_modules = (
    ["scorecard"]
    + [d for d, v in versions.items() if v is not None]
    + (["scenario"] if not baseline.empty else [])
)


def _filters_key(filters: query.Filters) -> tuple:
    return (
        filters.period_from, filters.period_to,
        tuple(filters.brands), tuple(filters.categories),
        tuple(filters.channels), tuple(filters.regions), tuple(filters.skus),
        tuple(_bundle_modules),
        adj.get("base", 0.0), adj.get("upside", 0.0), adj.get("downside", 0.0),
    )


_PPTX_STATE = "exec_pptx_state"
current_key = _filters_key(f)

with _export_slot.container():
    st.markdown("&nbsp;", unsafe_allow_html=True)

    try:
        _xlsx_bytes, _xlsx_name = xlsx_builder.build(
            modules=_bundle_modules, filters=f, scenario_adjustments=adj,
        )
    except Exception as e:  # noqa: BLE001
        st.error(f"Excel export failed: {e}")
    else:
        st.download_button(
            label="⬇ Full Report (Excel)",
            data=_xlsx_bytes,
            file_name=_xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    state = st.session_state.get(_PPTX_STATE)
    label = "⟳ Rebuild Full Deck" if state else "📊 Build Full Deck (PPTX)"
    if st.button(label, use_container_width=True, key="exec_build_pptx"):
        with st.spinner("Rendering full deck (kaleido)…"):
            try:
                data, name = pptx_builder.build(
                    modules=_bundle_modules, filters=f,
                    scenario_adjustments=adj,
                )
                st.session_state[_PPTX_STATE] = {
                    "key": current_key, "data": data,
                    "name": name, "error": None,
                }
            except Exception as e:  # noqa: BLE001
                st.session_state[_PPTX_STATE] = {
                    "key": current_key, "data": None,
                    "name": None, "error": str(e),
                }
        state = st.session_state[_PPTX_STATE]

    if state:
        if state["error"]:
            st.error(f"PPTX failed: {state['error']}")
        elif state["data"]:
            stale = state["key"] != current_key
            st.download_button(
                label="⬇ Download PPTX" + (" (stale)" if stale else ""),
                data=state["data"],
                file_name=state["name"],
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                help="Filters or scenarios changed — rebuild for fresh deck."
                if stale else None,
            )
