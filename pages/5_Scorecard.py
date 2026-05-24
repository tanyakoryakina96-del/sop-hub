"""KPI Scorecard — six tier-1 KPIs with RAG status (SPEC §5.4).

All KPI math + threshold semantics live in ``query.scorecard_all`` — this page
only renders. Same filter bar shape as the other dashboards so the shared
``filters`` session key remains compatible.
"""

import math
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import config  # noqa: F401, E402 — registers Plotly "tko" template
from data import query  # noqa: E402
from exports import pptx_builder, xlsx_builder  # noqa: E402

st.set_page_config(page_title="Scorecard — S&OP Hub", page_icon="🎯", layout="wide")

_hdr_l, _hdr_r = st.columns([6, 2])
with _hdr_l:
    st.title("🎯 KPI Scorecard")
    st.caption("Executive view: tier-1 KPIs with RAG status. Thresholds in `config.RAG`.")
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
        help="Channels apply to demand/financial; supply is plant-keyed and ignores them.",
    )
    r1[5].markdown("&nbsp;", unsafe_allow_html=True)
    reset = r1[5].button("Reset", use_container_width=True)

    r2 = st.columns(2)
    regions = r2[0].multiselect(
        "Region", options=reg_opts,
        default=_clamp(current.regions, reg_opts),
        key="flt_regions",
        help="Regions apply to demand only.",
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
# Data gate — need at least one fact table to compute any KPI
# ---------------------------------------------------------------------------

versions = {
    d: query.current_data_version(d) for d in ("demand", "supply", "financial")
}
if not any(versions.values()):
    st.warning(
        "No data uploaded yet. Visit the **Upload** page to load demand, supply, "
        "or financial CSVs before reviewing the Scorecard."
    )
    st.stop()

missing = [d for d, v in versions.items() if v is None]
if missing:
    st.info(
        "Missing data for: "
        + ", ".join(missing)
        + ". KPIs that need this data will appear as `—`."
    )

# ---------------------------------------------------------------------------
# Scorecard table
# ---------------------------------------------------------------------------

card = query.scorecard_all(f)


def _fmt_value(row: pd.Series) -> str:
    v = row["value"]
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "—"
    kpi = row["kpi_name"]
    unit = row["unit"]
    if kpi in ("Forecast Bias", "GM% vs Budget"):
        return f"{v:+.2f}{unit}" if unit == "pp" else f"{v:+.1f}%"
    if unit == "%":
        return f"{v:.1f}%"
    if unit == "days":
        return f"{v:.1f} days"
    if unit == "pp":
        return f"{v:+.2f}pp"
    return f"{v:.2f}"


display = card.copy()
display["Value"] = display.apply(_fmt_value, axis=1)
display = display.rename(columns={
    "kpi_name": "KPI",
    "rag": "RAG",
    "threshold_explainer": "Thresholds",
})[["KPI", "Value", "RAG", "Thresholds"]]


rag_colors = config.RAG_COLORS
void = config.TKO["colors"]["void"]


def _rag_style(val: str) -> str:
    bg = rag_colors.get(val, "")
    return f"background-color: {bg}; color: {void}; font-weight: 700; text-align: center;"


st.dataframe(
    display.style.map(_rag_style, subset=["RAG"]),
    use_container_width=True, hide_index=True,
    column_config={
        "KPI":        st.column_config.TextColumn("KPI", width="medium"),
        "Value":      st.column_config.TextColumn("Value", width="small"),
        "RAG":        st.column_config.TextColumn("RAG", width="small"),
        "Thresholds": st.column_config.TextColumn("Thresholds", width="large"),
    },
)

# Compact summary chips below the table — counts by status.
rag_counts = card["rag"].value_counts().to_dict()
c1, c2, c3 = st.columns(3)
c1.metric("🟢 Green", rag_counts.get("green", 0))
c2.metric("🟡 Amber", rag_counts.get("amber", 0))
c3.metric("🔴 Red",   rag_counts.get("red", 0))

st.caption(
    "Data versions: "
    + " · ".join(
        f"{d} v{v}" if v else f"{d} —"
        for d, v in versions.items()
    )
    + f" · Window: {f.period_from:%Y-%m} → {f.period_to:%Y-%m}"
)

# ---------------------------------------------------------------------------
# Exports — bundled deck (scorecard + any uploaded domain)
# ---------------------------------------------------------------------------
# Excel is rendered eagerly (sub-second). PPTX is gated behind a build button
# because kaleido's first chart render is 5–10s; we keep page interactions
# snappy and let the user opt in when they actually want a deck.

_bundle_modules = ["scorecard"] + [d for d, v in versions.items() if v is not None]


def _filters_key(filters: query.Filters) -> tuple:
    return (
        filters.period_from, filters.period_to,
        tuple(filters.brands), tuple(filters.categories),
        tuple(filters.channels), tuple(filters.regions), tuple(filters.skus),
        tuple(_bundle_modules),
    )


_PPTX_STATE = "scorecard_pptx_state"
current_key = _filters_key(f)

with _export_slot.container():
    st.markdown("&nbsp;", unsafe_allow_html=True)

    # XLSX — instant, render on every page load
    try:
        _xlsx_bytes, _xlsx_name = xlsx_builder.build(
            modules=_bundle_modules, filters=f,
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

    # PPTX — gated. Persist the artifact in session_state so the download
    # button survives Streamlit's per-interaction reruns until filters change.
    state = st.session_state.get(_PPTX_STATE)
    label = "⟳ Rebuild PPTX" if state else "📊 Build PPTX"
    if st.button(label, use_container_width=True, key="scorecard_build_pptx"):
        with st.spinner("Rendering charts (kaleido)…"):
            try:
                data, name = pptx_builder.build(
                    modules=_bundle_modules, filters=f,
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
                help="Filters changed — rebuild for an up-to-date deck." if stale else None,
            )
