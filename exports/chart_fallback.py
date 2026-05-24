"""Matplotlib chart renderer — kaleido fallback for PPTX export.

Used when ``exports.pptx_builder._fig_to_png`` (Plotly + kaleido) returns
``None`` because chromium IPC failed or timed out. Matplotlib runs entirely
in-process via the Agg backend — no subprocess, no chromium.

Charts here are deliberately simpler than the Plotly versions on the
dashboard pages: the goal is "data faithfully conveyed in a slide image,"
not visual parity with the interactive on-screen charts. TKO colors are
applied so the slides still feel branded.

This is a one-trick module: ``render(kind, df, *, title, width_px, height_px)``
returns PNG bytes. Add a new chart kind by extending ``_DISPATCH`` and
writing a ``_render_<kind>`` function. Each render function is ~25–40 LOC;
keeping them inline rather than abstracting a shared chart factory (per
CLAUDE.md rule §0 #8: inline > abstract).
"""

from __future__ import annotations

from io import BytesIO

import pandas as pd
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure

import config

# ---------------------------------------------------------------------------
# Theme — pull from TKO design tokens so slides stay on-brand
# ---------------------------------------------------------------------------

_C = config.TKO["colors"]
_S = config.SERIES
_RAG = config.RAG_COLORS

_BG_PAPER = _C["void"]
_BG_AXES  = _C["graphite"]
_GRID     = _C["cold_slate"]
_TEXT     = _C["ice_white"]
_TEXT_DIM = _C["ice_dim"]
_MUTED    = _C["mist"]


def _new_figure(width_px: int, height_px: int) -> tuple[Figure, "axes"]:
    """Build a TKO-themed Figure + single Axes. DPI 100 keeps inches simple."""
    dpi = 100
    fig = Figure(figsize=(width_px / dpi, height_px / dpi), dpi=dpi)
    fig.patch.set_facecolor(_BG_PAPER)
    ax = fig.add_subplot(111)
    ax.set_facecolor(_BG_AXES)
    for spine in ax.spines.values():
        spine.set_color(_GRID)
    ax.tick_params(colors=_TEXT_DIM, which="both")
    ax.xaxis.label.set_color(_TEXT_DIM)
    ax.yaxis.label.set_color(_TEXT_DIM)
    ax.title.set_color(_TEXT)
    ax.grid(True, axis="y", color=_GRID, alpha=0.4, linewidth=0.6)
    return fig, ax


def _to_png(fig: Figure) -> bytes:
    buf = BytesIO()
    canvas = FigureCanvasAgg(fig)
    canvas.print_png(buf)
    return buf.getvalue()


def _empty_chart(title: str, width_px: int, height_px: int) -> bytes:
    fig, ax = _new_figure(width_px, height_px)
    ax.text(
        0.5, 0.5, "No data in window",
        ha="center", va="center", color=_MUTED, fontsize=14,
        transform=ax.transAxes,
    )
    ax.set_xticks([])
    ax.set_yticks([])
    ax.set_title(title, fontsize=14, weight="bold", pad=10)
    fig.tight_layout()
    return _to_png(fig)


# ---------------------------------------------------------------------------
# Public entry
# ---------------------------------------------------------------------------


def render(
    kind: str,
    df: pd.DataFrame | None,
    *,
    title: str,
    width_px: int = 900,
    height_px: int = 540,
) -> bytes | None:
    """Render ``kind`` with data ``df`` to PNG bytes, or ``None`` on failure.

    Failure is logged via the calling pptx_builder which falls back further to
    a placeholder shape — but matplotlib is in-process and synchronous, so
    failures here are rare and indicate a coding bug rather than environment.
    """
    fn = _DISPATCH.get(kind)
    if fn is None:
        return None
    try:
        return fn(df, title=title, width_px=width_px, height_px=height_px)
    except Exception:  # noqa: BLE001 — defensive; never break the export
        return None


# ---------------------------------------------------------------------------
# Demand: Volume Bridge — 4-stage waterfall
# ---------------------------------------------------------------------------


def _render_demand_bridge(df, *, title, width_px, height_px) -> bytes:
    if df is None or df.empty or df["value"].abs().sum() == 0:
        return _empty_chart(title, width_px, height_px)
    fig, ax = _new_figure(width_px, height_px)

    stages = df["stage"].tolist()
    values = df["value"].astype(float).tolist()
    # First and last bars sit on the floor (absolute); middle two are deltas
    # from a running total, like a classic waterfall.
    running = values[0]
    bottoms = [0.0, running, running + (values[1] - values[0]), 0.0]
    heights = [values[0], values[1] - values[0], values[2] - values[1], values[3]]
    colors = [
        _S["budget"],
        _S["statistical_fcst"] if heights[1] < 0 else _S["consensus_fcst"],
        _S["statistical_fcst"] if heights[2] < 0 else _S["consensus_fcst"],
        _S["signal"],
    ]
    bars = ax.bar(stages, heights, bottom=bottoms, color=colors,
                  edgecolor=_GRID, linewidth=0.5)
    for bar, raw in zip(bars, values):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_y() + bar.get_height() + max(values) * 0.01,
            f"{raw:,.0f}", ha="center", va="bottom",
            color=_TEXT_DIM, fontsize=9,
        )
    ax.set_ylabel("Cases", color=_TEXT_DIM)
    ax.set_title(title, fontsize=14, weight="bold", pad=10)
    for label in ax.get_xticklabels():
        label.set_rotation(15)
        label.set_ha("right")
    fig.tight_layout()
    return _to_png(fig)


# ---------------------------------------------------------------------------
# Demand: Forecast vs Actuals — bars + 2 lines
# ---------------------------------------------------------------------------


def _render_demand_fcst_vs_actuals(df, *, title, width_px, height_px) -> bytes:
    if df is None or df.empty:
        return _empty_chart(title, width_px, height_px)
    fig, ax = _new_figure(width_px, height_px)
    x = pd.to_datetime(df["period_date"])
    ax.bar(x, df["actuals"], color=_S["actuals"], alpha=0.7,
           width=20, label="Actuals", edgecolor="none")
    ax.plot(x, df["consensus_fcst"], color=_S["consensus_fcst"],
            marker="o", linewidth=2.5, label="Consensus")
    ax.plot(x, df["statistical_fcst"], color=_S["statistical_fcst"],
            marker="s", linewidth=1.8, linestyle="--", label="Statistical")
    ax.set_ylabel("Cases", color=_TEXT_DIM)
    ax.set_title(title, fontsize=14, weight="bold", pad=10)
    leg = ax.legend(loc="upper left", facecolor=_BG_AXES, edgecolor=_GRID,
                    labelcolor=_TEXT, fontsize=9, framealpha=0.9)
    for txt in leg.get_texts():
        txt.set_color(_TEXT)
    fig.autofmt_xdate(rotation=20)
    fig.tight_layout()
    return _to_png(fig)


# ---------------------------------------------------------------------------
# Supply: DOS by SKU — horizontal bars, RAG colored, target vline
# ---------------------------------------------------------------------------


def _render_supply_dos(df, *, title, width_px, height_px) -> bytes:
    if df is None or df.empty or df["dos"].isna().all():
        return _empty_chart(title, width_px, height_px)
    d = df.dropna(subset=["dos"]).sort_values("dos").tail(15)
    fig, ax = _new_figure(width_px, height_px)
    colors = [_RAG.get(r, _RAG["red"]) for r in d["rag"]]
    ax.barh(d["sku_name"], d["dos"], color=colors, edgecolor=_GRID, linewidth=0.5)
    target = config.RAG["dos"]["target_days"]
    ax.axvline(target, color=_C["acid"], linewidth=2, linestyle=":")
    ax.text(
        target, len(d) - 0.5, f" target {target}d",
        color=_C["acid"], fontsize=9, va="top",
    )
    ax.set_xlabel("Days of Supply", color=_TEXT_DIM)
    ax.set_title(title, fontsize=14, weight="bold", pad=10)
    ax.tick_params(axis="y", labelsize=8)
    fig.tight_layout()
    return _to_png(fig)


# ---------------------------------------------------------------------------
# Supply: Production Adherence — grouped bars (plan vs actual) by plant
# ---------------------------------------------------------------------------


def _render_supply_adherence(df, *, title, width_px, height_px) -> bytes:
    if df is None or df.empty:
        return _empty_chart(title, width_px, height_px)
    fig, ax = _new_figure(width_px, height_px)
    import numpy as np
    n = len(df)
    x = np.arange(n)
    width = 0.4
    ax.bar(x - width / 2, df["plan"], width,
           label="Plan", color=_S["budget"], edgecolor="none")
    ax.bar(x + width / 2, df["actual"], width,
           label="Actual", color=_S["consensus_fcst"], edgecolor="none")
    ax.set_xticks(x)
    ax.set_xticklabels(df["plant_name"], rotation=15, ha="right")
    ax.set_ylabel("Cases", color=_TEXT_DIM)
    ax.set_title(title, fontsize=14, weight="bold", pad=10)
    leg = ax.legend(facecolor=_BG_AXES, edgecolor=_GRID, fontsize=9, framealpha=0.9)
    for txt in leg.get_texts():
        txt.set_color(_TEXT)
    fig.tight_layout()
    return _to_png(fig)


# ---------------------------------------------------------------------------
# Financial: Revenue Waterfall — Budget → Δ to LE → Δ to Actual → Actual
# ---------------------------------------------------------------------------


def _render_financial_waterfall(df, *, title, width_px, height_px) -> bytes:
    if df is None or df.empty or df["value"].abs().sum() == 0:
        return _empty_chart(title, width_px, height_px)
    fig, ax = _new_figure(width_px, height_px)
    budget = float(df.loc[df["stage"] == "Budget", "value"].iloc[0])
    le     = float(df.loc[df["stage"] == "LE", "value"].iloc[0])
    actual = float(df.loc[df["stage"] == "Actuals", "value"].iloc[0])
    delta_le  = le - budget
    delta_act = actual - le

    labels = ["Budget", "Δ to LE", "Δ to Actual", "Actual"]
    heights = [budget, delta_le, delta_act, actual]
    bottoms = [0.0, budget, le, 0.0]
    def _col(h: float, total_bar: bool) -> str:
        if total_bar:
            return _S["signal"]
        return _S["consensus_fcst"] if h >= 0 else _S["statistical_fcst"]
    colors = [
        _S["budget"],
        _col(delta_le, False),
        _col(delta_act, False),
        _col(actual, True),
    ]
    bars = ax.bar(labels, heights, bottom=bottoms, color=colors,
                  edgecolor=_GRID, linewidth=0.5)
    texts = [
        f"{budget:,.0f}",
        f"{delta_le:+,.0f}",
        f"{delta_act:+,.0f}",
        f"{actual:,.0f}",
    ]
    ymax = max(budget, le, actual)
    for bar, txt in zip(bars, texts):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_y() + bar.get_height() + ymax * 0.01,
            txt, ha="center", va="bottom", color=_TEXT_DIM, fontsize=9,
        )
    ax.set_ylabel(f"Revenue ({config.CURRENCY})", color=_TEXT_DIM)
    ax.set_title(title, fontsize=14, weight="bold", pad=10)
    fig.tight_layout()
    return _to_png(fig)


# ---------------------------------------------------------------------------
# Financial: Revenue by Channel — horizontal bars
# (Substituted for the on-screen treemap because matplotlib has no native
#  treemap, and horizontal bars are clearer in a slide-sized image anyway.)
# ---------------------------------------------------------------------------


def _render_financial_by_channel(df, *, title, width_px, height_px) -> bytes:
    if df is None or df.empty or df["revenue_actual"].abs().sum() == 0:
        return _empty_chart(title, width_px, height_px)
    d = df.sort_values("revenue_actual", ascending=True)
    fig, ax = _new_figure(width_px, height_px)
    ax.barh(
        d["channel_name"], d["revenue_actual"],
        color=_S["consensus_fcst"], edgecolor=_GRID, linewidth=0.5,
    )
    total = float(d["revenue_actual"].sum())
    for i, v in enumerate(d["revenue_actual"]):
        pct = (v / total * 100) if total else 0
        ax.text(
            v, i, f"  {v:,.0f}  ({pct:.1f}%)",
            va="center", color=_TEXT_DIM, fontsize=9,
        )
    ax.set_xlabel(f"Revenue ({config.CURRENCY})", color=_TEXT_DIM)
    ax.set_title(title, fontsize=14, weight="bold", pad=10)
    fig.tight_layout()
    return _to_png(fig)


# ---------------------------------------------------------------------------
# Scenario: Revenue per Scenario — colored bars (Base / Upside / Downside)
# ---------------------------------------------------------------------------


def _render_scenario_revenue(df, *, title, width_px, height_px) -> bytes:
    if df is None or df.empty:
        return _empty_chart(title, width_px, height_px)
    fig, ax = _new_figure(width_px, height_px)
    palette = {
        "Base":     _S["consensus_fcst"],
        "Upside":   _RAG["green"],
        "Downside": _RAG["red"],
    }
    colors = [palette.get(s, _S["consensus_fcst"]) for s in df["scenario"]]
    bars = ax.bar(df["scenario"], df["revenue"], color=colors,
                  edgecolor=_GRID, linewidth=0.5)
    ymax = float(df["revenue"].max() or 0)
    for bar, v in zip(bars, df["revenue"]):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_y() + bar.get_height() + ymax * 0.01,
            f"{v:,.0f}", ha="center", va="bottom",
            color=_TEXT_DIM, fontsize=10,
        )
    ax.set_ylabel(f"Revenue ({config.CURRENCY})", color=_TEXT_DIM)
    ax.set_title(title, fontsize=14, weight="bold", pad=10)
    fig.tight_layout()
    return _to_png(fig)


_DISPATCH = {
    "demand_bridge":           _render_demand_bridge,
    "demand_fcst_vs_actuals":  _render_demand_fcst_vs_actuals,
    "supply_dos":              _render_supply_dos,
    "supply_adherence":        _render_supply_adherence,
    "financial_waterfall":     _render_financial_waterfall,
    "financial_by_channel":    _render_financial_by_channel,
    "scenario_revenue":        _render_scenario_revenue,
}
