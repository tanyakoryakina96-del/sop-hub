"""Scenario Planner — three fixed scenarios with % volume adjustment (SPEC §5.5).

Sliders write to ``st.session_state["scenario_adjustments"]`` (in-session only,
not persisted — see CONTRACTS §4). The page computes scenario volume × trailing
3-month avg price per (SKU, channel) and aggregates to a side-by-side comparison.
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

st.set_page_config(page_title="Scenario — S&OP Hub", page_icon="🎲", layout="wide")

_hdr_l, _hdr_r = st.columns([6, 2])
with _hdr_l:
    st.title("🎲 Scenario Planner")
    st.caption(
        "Three fixed scenarios — Base / Upside / Downside — applied uniformly "
        "as a % adjustment to the forward consensus volume. Inputs are not "
        "persisted across restarts (SPEC §5.5)."
    )
_export_slot = _hdr_r.empty()

# ---------------------------------------------------------------------------
# Filter bar (shared shape — SPEC §10.2)
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
# Data gate — need demand for baseline; financial for revenue
# ---------------------------------------------------------------------------

dv = query.current_data_version("demand")
fv = query.current_data_version("financial")
if dv is None:
    st.warning(
        "No demand data uploaded yet. Visit the **Upload** page and load a "
        "demand CSV before using the Scenario Planner."
    )
    st.stop()
if fv is None:
    st.info(
        "No financial data uploaded — scenarios will show volume only "
        "(no revenue projection). Upload a financial CSV to enable the "
        "revenue column."
    )

# ---------------------------------------------------------------------------
# Sliders → session_state
# ---------------------------------------------------------------------------

DEFAULTS = {"base": 0.0, "upside": 5.0, "downside": -5.0}
if "scenario_adjustments" not in st.session_state:
    st.session_state["scenario_adjustments"] = dict(DEFAULTS)
adj: dict = st.session_state["scenario_adjustments"]

st.markdown("##### Scenario inputs · % volume vs consensus")
s1, s2, s3 = st.columns(3)
adj["base"] = s1.slider(
    "Base", min_value=-30.0, max_value=30.0,
    value=float(adj.get("base", 0.0)), step=0.5, format="%+.1f%%",
    key="scen_base",
    help="Reference scenario. Usually 0% = take consensus as-is.",
)
adj["upside"] = s2.slider(
    "Upside", min_value=-30.0, max_value=30.0,
    value=float(adj.get("upside", 5.0)), step=0.5, format="%+.1f%%",
    key="scen_upside",
    help="Optimistic adjustment to forward consensus volume.",
)
adj["downside"] = s3.slider(
    "Downside", min_value=-30.0, max_value=30.0,
    value=float(adj.get("downside", -5.0)), step=0.5, format="%+.1f%%",
    key="scen_downside",
    help="Pessimistic adjustment to forward consensus volume.",
)
st.session_state["scenario_adjustments"] = adj

# ---------------------------------------------------------------------------
# Queries — baseline (volume) + price proxy
# ---------------------------------------------------------------------------

baseline = query.scenario_baseline(f)
price = query.scenario_price_proxy(f) if fv is not None else pd.DataFrame(
    columns=["sku_code", "channel_code", "avg_price"]
)

if baseline.empty:
    st.warning(
        "**No forward planning horizon in this window.** "
        "Consensus volume after the latest actuals period is empty — "
        "scenarios have no months to project. Widen the period window so it "
        "extends past the actuals frontier, or upload demand data with "
        "forward-looking consensus."
    )
    st.stop()

# Pairs in baseline but missing a price → contribute volume but not revenue.
merged = baseline.merge(price, on=["sku_code", "channel_code"], how="left")
missing_price = int(merged["avg_price"].isna().sum())
if missing_price and fv is not None:
    st.info(
        f"**{missing_price}** SKU-channel pair(s) excluded from revenue — "
        "no revenue history in trailing 3 months. Volume column still includes them."
    )

# ---------------------------------------------------------------------------
# Scenario summary (volume + revenue) — aggregate across (sku, channel)
# ---------------------------------------------------------------------------

scen_rows: list[dict] = []
for name in ("Base", "Upside", "Downside"):
    pct = float(adj.get(name.lower(), 0.0))
    vol = merged["consensus_volume_cases"] * (1.0 + pct / 100.0)
    rev = (vol * merged["avg_price"]).fillna(0.0)
    scen_rows.append({
        "Scenario": name,
        "% vs Consensus": pct,
        "Volume (cases)": float(vol.sum()),
        f"Revenue ({config.CURRENCY})": float(rev.sum()),
    })
summary = pd.DataFrame(scen_rows)

base_revenue = float(summary.loc[summary["Scenario"] == "Base",
                                 f"Revenue ({config.CURRENCY})"].iloc[0])
base_volume = float(summary.loc[summary["Scenario"] == "Base",
                                "Volume (cases)"].iloc[0])

summary["Δ Revenue vs Base"] = summary[f"Revenue ({config.CURRENCY})"] - base_revenue
summary["Δ Revenue %"] = summary[f"Revenue ({config.CURRENCY})"].apply(
    lambda v: (v - base_revenue) / base_revenue * 100 if base_revenue else float("nan")
)

# ---------------------------------------------------------------------------
# Render — KPI cards + table + chart
# ---------------------------------------------------------------------------

c1, c2, c3 = st.columns(3)
for col, name in zip((c1, c2, c3), ("Base", "Upside", "Downside")):
    row = summary[summary["Scenario"] == name].iloc[0]
    rev = row[f"Revenue ({config.CURRENCY})"]
    vol = row["Volume (cases)"]
    delta_pct = row["Δ Revenue %"]
    delta_label = (
        None if name == "Base" or math.isnan(delta_pct)
        else f"{delta_pct:+.1f}% vs Base"
    )
    col.metric(
        f"{name} · {row['% vs Consensus']:+.1f}%",
        f"{rev:,.0f} {config.CURRENCY}" if fv is not None else f"{vol:,.0f} cases",
        delta=delta_label,
    )

st.markdown("##### Comparison")

# Style: highlight Upside green / Downside red on the scenario column.
rag_colors = config.RAG_COLORS


def _scen_style(row: pd.Series) -> list[str]:
    styles = [""] * len(row)
    if row["Scenario"] == "Upside":
        styles[0] = (
            f"background-color: {rag_colors['green']}; "
            f"color: {config.TKO['colors']['void']}; font-weight: 700;"
        )
    elif row["Scenario"] == "Downside":
        styles[0] = (
            f"background-color: {rag_colors['red']}; "
            f"color: {config.TKO['colors']['void']}; font-weight: 700;"
        )
    else:
        styles[0] = "font-weight: 700;"
    return styles


st.dataframe(
    summary.style
    .apply(_scen_style, axis=1)
    .format({
        "% vs Consensus": "{:+.1f}%",
        "Volume (cases)": "{:,.0f}",
        f"Revenue ({config.CURRENCY})": "{:,.0f}",
        "Δ Revenue vs Base": "{:+,.0f}",
        "Δ Revenue %": "{:+.1f}%",
    }),
    use_container_width=True, hide_index=True,
)

# Bar chart — revenue per scenario.
palette = {
    "Base":     config.SERIES["consensus_fcst"],
    "Upside":   config.RAG_COLORS["green"],
    "Downside": config.RAG_COLORS["red"],
}
fig_metric = f"Revenue ({config.CURRENCY})" if fv is not None else "Volume (cases)"
fig = go.Figure(go.Bar(
    x=summary["Scenario"], y=summary[fig_metric],
    marker_color=[palette[s] for s in summary["Scenario"]],
    text=[f"{v:,.0f}" for v in summary[fig_metric]],
    textposition="outside",
))
fig.update_layout(
    title=f"Scenario {fig_metric}",
    showlegend=False, yaxis_title=fig_metric,
    margin=dict(l=56, r=24, t=56, b=40),
)
st.plotly_chart(fig, use_container_width=True)

st.caption(
    f"Forward planning rows: {len(baseline):,} (SKU-channel pairs) · "
    f"Base volume: {base_volume:,.0f} cases · "
    f"Window: {f.period_from:%Y-%m} → {f.period_to:%Y-%m}"
)

# ---------------------------------------------------------------------------
# Exports — bundled deck with scenario module
# ---------------------------------------------------------------------------

_versions = {
    d: query.current_data_version(d) for d in ("demand", "supply", "financial")
}
_bundle_modules = (
    ["scorecard"]
    + [d for d, v in _versions.items() if v is not None]
    + ["scenario"]
)


def _filters_key(filters: query.Filters) -> tuple:
    return (
        filters.period_from, filters.period_to,
        tuple(filters.brands), tuple(filters.categories),
        tuple(filters.channels), tuple(filters.regions), tuple(filters.skus),
        tuple(_bundle_modules),
        adj["base"], adj["upside"], adj["downside"],
    )


_PPTX_STATE = "scenario_pptx_state"
current_key = _filters_key(f)

with _export_slot.container():
    st.markdown("&nbsp;", unsafe_allow_html=True)

    try:
        _xlsx_bytes, _xlsx_name = xlsx_builder.build(
            modules=_bundle_modules, filters=f,
            scenario_adjustments=adj,
        )
    except Exception as e:  # noqa: BLE001
        st.error(f"Excel export failed: {e}")
    else:
        st.download_button(
            label="⬇ Excel (bundled)",
            data=_xlsx_bytes,
            file_name=_xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    state = st.session_state.get(_PPTX_STATE)
    label = "⟳ Rebuild PPTX" if state else "📊 Build PPTX"
    if st.button(label, use_container_width=True, key="scenario_build_pptx"):
        with st.spinner("Rendering charts (kaleido)…"):
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
