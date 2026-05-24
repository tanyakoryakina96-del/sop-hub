"""All app-wide constants live here.

Pages and exporters read from this module directly. Renaming a top-level key is
a contract change — see CONTRACTS.md §5.5.

Side effect on import: registers the Plotly template `"tko"` and sets it as
default, so importing this module is enough to brand every chart in the app.
"""

import os

import plotly.graph_objects as go
import plotly.io as pio

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
STORAGE_DIR = os.path.join(PROJECT_ROOT, "storage")
RAW_DIR = os.path.join(STORAGE_DIR, "raw")
TEMPLATES_DIR = os.path.join(STORAGE_DIR, "templates")
EXPORTS_DIR = os.path.join(STORAGE_DIR, "exports")
SCHEMA_SQL_PATH = os.path.join(PROJECT_ROOT, "data", "schema.sql")
# DuckDB file path is no longer a single constant — see data/session_db.py
# (each Streamlit session opens its own DB so visitors don't collide).

_TKO_TEMPLATE = os.path.join(TEMPLATES_DIR, "tko_template.pptx")
TKO_TEMPLATE_PATH = _TKO_TEMPLATE if os.path.exists(_TKO_TEMPLATE) else None

# PNG only — python-pptx cannot embed SVG directly (SPEC §7.4 step 5).
TKO_LOGO_PATH = os.path.join(PROJECT_ROOT, "assets", "tko", "assets", "tko-wordmark.png")

for _d in (STORAGE_DIR, RAW_DIR, TEMPLATES_DIR, EXPORTS_DIR):
    os.makedirs(_d, exist_ok=True)

# ---------------------------------------------------------------------------
# TKO design tokens (extracted from colors_and_type.css)
# ---------------------------------------------------------------------------

TKO = {
    "colors": {
        "void":        "#0A0C14",  # slide background
        "graphite":    "#13162A",  # card surfaces
        "cold_slate":  "#1E2235",  # borders
        "plasma":      "#6D28D9",  # primary accent (violet)
        "neon_indigo": "#818CF8",  # labels / links
        "acid":        "#C8E000",  # signal — KPI numbers, highlights
        "ice_white":   "#E8EEFF",  # primary text
        "ice_dim":     "#C8CCEE",  # secondary text
        "mist":        "#7880A0",  # body text
        "fog":         "#4A506E",  # captions
    },
    "fonts": {
        "display": "Barlow Condensed",  # uppercase titles
        "body":    "Inter",             # body text
    },
    "slide": {
        "width_emu":  9144000,   # 10 inches (16:9)
        "height_emu": 5143500,   # 5.625 inches
        "bg_color":   "#0A0C14",
    },
}

# Semantic chart-series colors — see SPEC §10.3
SERIES = {
    "actuals":          TKO["colors"]["ice_white"],
    "consensus_fcst":   TKO["colors"]["plasma"],
    "statistical_fcst": TKO["colors"]["neon_indigo"],
    "budget":           TKO["colors"]["mist"],
    "le":               TKO["colors"]["ice_dim"],
    "signal":           TKO["colors"]["acid"],
}

RAG_COLORS = {
    "green": TKO["colors"]["acid"],  # acid green-on-dark reads as positive
    "amber": "#F59E0B",
    "red":   "#EF4444",
}

# ---------------------------------------------------------------------------
# RAG thresholds — SPEC §5.4
# ---------------------------------------------------------------------------

RAG = {
    "fa":              {"green": 80, "amber": 70},          # FA% (100 - MAPE)
    "bias":            {"green_abs": 5, "amber_abs": 10},
    "dos":             {"target_days": 30, "green_band": 5, "amber_band": 10},
    "fill_rate":       {"green": 98, "amber": 95},
    "rev_vs_budget":   {"green": 98, "amber": 95},
    "gm_vs_budget_pp": {"green": 0, "amber": -1},
}

# ---------------------------------------------------------------------------
# Single-currency MVP (rule §3.10)
# ---------------------------------------------------------------------------

CURRENCY = "EUR"

# Months subtracted from period_to in default_filters() — yields a 13-month
# inclusive R13M window (SPEC §4.4).
DEFAULT_LOOKBACK_M = 12

# ---------------------------------------------------------------------------
# Plotly template — register & set default on import
# ---------------------------------------------------------------------------

_c = TKO["colors"]

_AXIS = dict(
    gridcolor=_c["cold_slate"],
    zerolinecolor=_c["cold_slate"],
    linecolor=_c["cold_slate"],
    tickcolor=_c["cold_slate"],
    # Axis tick + title text uses muted greys per CSS hierarchy:
    # tick labels = mist (body), axis titles = ice_dim (near-primary).
    tickfont=dict(family=TKO["fonts"]["body"], color=_c["mist"], size=11),
    title=dict(font=dict(family=TKO["fonts"]["body"], color=_c["ice_dim"], size=12)),
)

# `acid` is the SIGNAL color (SPEC §10.3 / CSS comment: "numbers, highlights
# only") — keep it out of the default colorway so it never accidentally paints
# a structural series. Pages that want acid pull it explicitly via
# `config.SERIES["signal"]`.
_tko_template = go.layout.Template(
    layout=go.Layout(
        paper_bgcolor=_c["void"],
        plot_bgcolor=_c["graphite"],
        font=dict(family=TKO["fonts"]["body"], color=_c["ice_white"], size=13),
        title=dict(font=dict(family=TKO["fonts"]["display"], color=_c["ice_white"], size=20)),
        colorway=[
            _c["plasma"], _c["neon_indigo"], _c["ice_dim"],
            _c["mist"], _c["fog"], _c["acid"],
        ],
        xaxis=_AXIS,
        yaxis=_AXIS,
        legend=dict(
            bgcolor=_c["graphite"],
            bordercolor=_c["cold_slate"],
            borderwidth=1,
            font=dict(family=TKO["fonts"]["body"], color=_c["ice_dim"], size=11),
        ),
        hoverlabel=dict(
            bgcolor=_c["graphite"],
            bordercolor=_c["cold_slate"],
            font=dict(family=TKO["fonts"]["body"], color=_c["ice_white"], size=12),
        ),
        margin=dict(l=56, r=24, t=64, b=48),
    )
)

pio.templates["tko"] = _tko_template
pio.templates.default = "tko"
