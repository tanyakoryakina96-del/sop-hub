# S&OP Hub — Module Contracts

**Purpose:** Pinned interface surface between modules. Every chunk implements *against* this file. If a chunk needs to change a signature, update CONTRACTS.md *first*, then write the code — never the reverse.

**Source of truth:** SPEC.md is canonical for business rules, formulas, RAG thresholds, and rationale. This file extracts only the load-bearing technical contracts so a new session can resume without re-reading 1,100 lines.

---

## 1. Module Map

```
app.py                  Entry point. Title + data-version chip. No business logic.
config.py               All constants: paths, RAG thresholds, TKO tokens, Plotly template.
.streamlit/config.toml  Port 3000, theme.

data/
  schema.sql            DuckDB DDL — single source of truth for table shapes.
  ingest.py             CSV → DuckDB. Owns: detect_schema, parse, validate, load, list_uploads.
  query.py              DuckDB → DataFrame/dataclass. Owns: Filters, current_data_version,
                        per-chart query funcs, dimension list helpers. All @st.cache_data.

pages/
  1_Upload.py           Uploader + column mapper + validation UI.
  2_Demand.py           Demand Review dashboard.
  3_Supply.py           Supply Review dashboard.
  4_Financial.py        Financial Review dashboard.
  5_Scorecard.py        KPI Scorecard.
  6_Scenario.py         Scenario Planner.
  7_Exec_Summary.py     Executive Summary (locked view + bundled export).

exports/
  pptx_builder.py       python-pptx + kaleido (primary) / matplotlib (fallback). Owns:
                        build(modules, filters, scenario_adjustments, template_path).
  xlsx_builder.py       openpyxl. Owns: build(modules, filters, scenario_adjustments).
  chart_fallback.py     matplotlib (Agg). In-process PNG renderer used when
                        kaleido times out. Owns: render(kind, df, *, title, width_px, height_px).

storage/                Runtime — gitignored. sessions/ (per-Streamlit-session DuckDB files),
                        sop.duckdb (fallback for non-Streamlit invocations),
                        raw/, templates/, exports/.
assets/tko/             Unzipped TKO design system — gitignored.
```

**Rule:** No new file appears unless this map is updated first. Target ≤ 12 Python files.

`chart_fallback.py` was added in Phase 5 because `kaleido==0.2.1` on Windows 11
hangs on subprocess IPC. It renders the same 6 chart kinds the PPTX needs,
using matplotlib's Agg backend (no chromium). See memory `kaleido-windows-hang`
for the diagnostic trail.

---

## 2. Data Contracts (DuckDB)

The DDL in `data/schema.sql` is authoritative. CONTRACTS.md just lists the **invariants** that callers depend on:

- All `period_date` values are **first-of-month** (normalized in `ingest.load`).
- All fact tables include `data_version: INTEGER` + `loaded_at: TIMESTAMP`.
- Fact PKs (composite, include `data_version`):
  - `fact_demand`    → `(period_date, sku_code, channel_code, region_code, data_version)`
  - `fact_supply`    → `(period_date, sku_code, plant_code, data_version)`
  - `fact_financial` → `(period_date, sku_code, channel_code, data_version)`
- `upload_log.status ∈ {'active', 'superseded', 'error'}`. Exactly one `'active'` row per domain at a time.
- DOS is **not** stored — computed at query time joining `fact_supply` to `fact_demand`.

**Canonical field names** (what `mapping` dict in `ingest.validate/load` targets) — see SPEC §4.1. Treat that table as the contract for what the column mapper exposes.

---

## 3. Cross-cutting Rules

These apply across all modules; breaking them is a contract violation.

1. **`data_version` filter is mandatory on every fact read.** Every query that touches a fact table must include `WHERE data_version = current_data_version('<domain>')` (the helper returns the active `int`, or `None` if no upload exists yet — in which case the query returns an empty result without hitting the table). No raw `SELECT * FROM fact_*` without this filter.
2. **Caching is at the query layer only.** Every public function in `query.py` is `@st.cache_data(ttl=600)`; dim helpers `ttl=3600`. Plotly figures are never cached. `current_data_version()` is never cached.
3. **`ingest.load` ordering — DB-transactional, then raw-file copy.** The DuckDB transaction wraps three writes atomically: fact-table inserts, `data_version` bump, `upload_log` row. After the transaction commits, the original CSV is copied to `storage/raw/`. If the file copy fails, log a warning to `upload_log.error_msg` but **do not roll back** — the DB state is consistent and the original CSV is still in the user's session memory. On a successful `load()`, the caller (`1_Upload.py`) calls `st.cache_data.clear()`.
4. **Exports return `bytes`, not `BytesIO`.** Callers pipe them straight into `st.download_button`.
5. **Filters dataclass is the only argument shape across query functions.** No ad-hoc kwargs.
6. **No I/O outside `data/` and `exports/`.** Pages render; they do not read files or open DuckDB directly.
7. **Empty list in `Filters` dim fields means "no filter on that dimension."** Query builders skip the `WHERE col IN (...)` clause entirely when the list is empty. Never coerce to `[None]` or `['ALL']`.
8. **Demand is the planning anchor.** When a date or period default is needed app-wide, it derives from `fact_demand` (latest active version), with today's first-of-month as fallback when no demand exists yet.
9. **UOM normalization rules** (per SPEC §4.3). The fact tables store volumes in each SKU's native UOM (cases, kg, units, tonnes…); `dim_sku.uom_to_cases` is the conversion factor. Three patterns apply:
   - **Cross-SKU aggregation (always normalize):** Any `SUM()` of a volume column across multiple SKUs **must** join `dim_sku` and multiply by `uom_to_cases` before summing. Examples: `demand_kpis.volume_total`, `supply_kpis.dos_avg / fill_rate / production_adherence`, `supply_capacity_utilization`, `supply_production_adherence` rows, scenario revenue aggregation, `scenario_price_proxy.avg_price` denominator.
   - **Per-SKU comparison columns (normalize to cases):** Columns shown side-by-side across SKU rows in a table or sort key **should** be normalized to cases so users can compare meaningfully. Examples: `demand_mape_by_sku.volume` (sort key for ranking).
   - **Per-SKU intrinsic values (keep native UOM):** Single-row metrics that already carry their own meaningful unit and aren't compared cross-row stay native. Examples: `supply_dos_by_sku.dos` (days — UOM-independent), `supply_gaps.shortfall` (in SKU's UOM for ordering decisions).

   When in doubt, follow the principle: any value the user might `SUM` or `compare across rows` is cases; any value with an intrinsic unit (days, %) or that represents a single planning quantity stays native.
10. **Single currency for MVP.** All `revenue_*`, `gm_*`, and `promo_spend_*` values are assumed to be in `config.CURRENCY` (default 'EUR'). `fact_financial.currency_code` is stored but not converted; rows with a non-default currency are flagged in validation but loaded as-is. Multi-currency FX is post-MVP.

---

## 4. Session State Keys

The *entire* cross-page state contract. Adding a key requires updating this table.

| Key | Type | Set by | Read by |
|-----|------|--------|---------|
| `filters` | `Filters` | Each dashboard page constructs a `Filters` from individual widget values (widgets bind to their own primitive `st.session_state` keys via `key=`) and writes the assembled dataclass here | All dashboard pages, export buttons |
| `data_versions` | `dict[str, int]` keyed by domain — e.g. `{'demand': 3, 'supply': 2, 'financial': 1}` | `1_Upload.py` after successful `load()`; updates the key for the domain just loaded | `app.py` (header chip — shows the demand version, or `"—"` if absent), all dashboard pages |
| `last_upload_status` | `str` | `1_Upload.py` | `1_Upload.py` (confirmation banner) |
| `column_mapping_template` | `dict[str, dict[str, str]]` keyed by domain | `1_Upload.py` on mapping confirm | `1_Upload.py` on next upload |
| `scenario_adjustments` | `dict` — `{'base': float, 'upside': float, 'downside': float}` (% vs consensus) | `6_Scenario.py` sliders | `6_Scenario.py`, `7_Exec_Summary.py`. **Not** read implicitly by exporters — pages pass it explicitly via the `scenario_adjustments` parameter on `pptx_builder.build` (§5.3) and `xlsx_builder.build` (§5.4). |
| `exec_filters_frozen` | `bool` | `7_Exec_Summary.py` (Freeze toggle) | `7_Exec_Summary.py` |

**Note:** The header chip on `app.py` shows the demand version specifically (treated as the primary clock for the S&OP cycle). Other domain versions are visible on the Upload page's history expander.

---

## 5. Python Function Contracts

### 5.1 `data/ingest.py`

```python
@dataclass
class DetectedColumn:
    name: str
    detected_type: str            # 'date' | 'int' | 'float' | 'text' | 'bool'
    sample_values: list[str]
    suggested_field: str | None   # canonical field name (§4.1) or None

@dataclass
class SchemaReport:
    row_count: int
    columns: list[DetectedColumn]
    encoding_used: str            # 'utf-8' | 'utf-16' | 'latin-1'

@dataclass
class ValidationIssue:
    severity: str                 # 'warning' | 'error'
    row: int | None
    column: str | None
    message: str

@dataclass
class ValidationReport:
    issues: list[ValidationIssue]
    # Derived: .errors, .warnings, .can_load (bool: no errors)


def detect_schema(file_bytes: bytes) -> SchemaReport
def parse(file_bytes: bytes, encoding: str) -> pd.DataFrame
def validate(df: pd.DataFrame, domain: str, mapping: dict[str, str]) -> ValidationReport
def load(
    df: pd.DataFrame,
    domain: str,
    mapping: dict[str, str],
    mode: str,
    filename: str,
    file_bytes: bytes,
) -> int
    # mode ∈ {'replace', 'append'}. Returns new data_version. Transactional. See SPEC §9.1.
    # `filename` is recorded in upload_log; `file_bytes` is copied to storage/raw/
    # after the DB transaction commits (per rule §3.3).
def list_uploads() -> pd.DataFrame
    # cols: filename, domain, uploaded_at, row_count, data_version, status, error_msg
    # Ordered by uploaded_at DESC. Mirrors upload_log columns.
```

`domain` ∈ `{'demand', 'supply', 'financial', 'sku_master', 'channel_master', 'region_master', 'plant_master'}`.

### 5.2 `data/query.py`

```python
@dataclass
class Filters:
    period_from: date
    period_to: date
    brands: list[str] = field(default_factory=list)       # empty = "all" (no filter)
    categories: list[str] = field(default_factory=list)
    channels: list[str] = field(default_factory=list)
    regions: list[str] = field(default_factory=list)
    skus: list[str] = field(default_factory=list)

# Default factory — used by pages when st.session_state lacks "filters":
def default_filters() -> Filters
    # period_to  = first-of-month of MAX(period_date) in active fact_demand,
    #              or first-of-current-month if no demand uploaded yet.
    # period_from = period_to - config.DEFAULT_LOOKBACK_M months
    #               (default 12 → 13-month inclusive R13M window, SPEC §4.4)
    # All dim lists empty.
    # NOT cached (depends on current_data_version output).

@dataclass
class DemandKPIs:
    mape: float           # %
    fa_pct: float         # 100 - mape
    bias: float           # %
    volume_total: float

@dataclass
class SupplyKPIs:
    dos_avg: float
    fill_rate: float                # %
    production_adherence: float     # %

@dataclass
class FinancialKPIs:
    revenue_actual: float
    revenue_vs_budget_pct: float
    revenue_vs_le_pct: float
    gm_pct: float


# Helper — NOT cached
def current_data_version(domain: str) -> int | None

# Demand   (all @st.cache_data, ttl=600)
def demand_kpis(f: Filters) -> DemandKPIs
def demand_volume_bridge(f: Filters) -> pd.DataFrame      # cols: stage, value
def demand_fcst_vs_actuals(f: Filters) -> pd.DataFrame    # cols: period_date, statistical_fcst, consensus_fcst, actuals
def demand_mape_by_sku(f: Filters) -> pd.DataFrame        # cols: sku_code, sku_name, brand, mape, fa_pct, bias, volume, rag
    # `rag` ∈ {'green','amber','red'} computed against config.RAG['fa'].
    # `volume` is normalized to cases for cross-SKU comparison (rule §3.9, comparison-column pattern).
def demand_channel_mix(f: Filters) -> pd.DataFrame        # cols: period_date, channel_code, channel_name, volume_share

# Supply
# DOS join semantics: fact_supply LEFT JOIN fact_demand on (period_date, sku_code).
# Same-period consensus_fcst is the divisor. If no matching demand row → DOS = NULL,
# excluded from aggregations. (Forward-looking demand window is post-MVP.)
def supply_kpis(f: Filters) -> SupplyKPIs
    # Window semantics:
    #   dos_avg              — POINT-IN-TIME at the latest period in f's window.
    #                          Demand-weighted, UOM-normalized to cases (rule §3.9):
    #                          SUM(inventory_qty * uom_to_cases) / NULLIF(SUM(consensus_fcst * uom_to_cases) / 30.44, 0)
    #                          NOT a simple mean of per-SKU DOS.
    #   fill_rate            — AGGREGATED over the full filter window:
    #                          SUM(orders_delivered * uom_to_cases) / NULLIF(SUM(orders_requested * uom_to_cases), 0) * 100
    #   production_adherence — AGGREGATED over the full filter window:
    #                          SUM(production_actual * uom_to_cases) / NULLIF(SUM(production_plan * uom_to_cases), 0) * 100
def supply_dos_by_sku(f: Filters) -> pd.DataFrame         # cols: sku_code, sku_name, dos, rag
    # Latest period in f's window. DOS in days (no UOM mixing — single SKU per row).
def supply_production_adherence(f: Filters) -> pd.DataFrame  # cols: plant_code, plant_name, plan, actual, adherence_pct
    # Aggregated across the filter window. plan/actual in cases (UOM-normalized).
def supply_capacity_utilization(f: Filters) -> pd.DataFrame  # cols: plant_code, plant_name, actual, capacity, utilization_pct
    # SUM(production_actual) / NULLIF(SUM(capacity_plan), 0) * 100 across filter window, per plant.
    # Both numerator and denominator UOM-normalized to cases. Plants with no capacity_plan rows are omitted.
def supply_inventory_heatmap(f: Filters) -> pd.DataFrame  # cols: sku_code, period_date, dos, rag
    # Per (sku, period) cell. No UOM aggregation needed (single SKU per row).
def supply_gaps(f: Filters) -> pd.DataFrame               # cols: sku_code, sku_name, period_date, shortfall
    # Projected stockouts in next 3 months. `shortfall` in the SKU's native UOM.
def supply_fill_rate_trend(f: Filters) -> pd.DataFrame    # cols: period_date, fill_rate
    # Monthly fill rate across the filter window, for the KPI card sparkline (SPEC §5.2).
    # Per period: SUM(orders_delivered * uom_to_cases) / NULLIF(SUM(orders_requested * uom_to_cases), 0) * 100

# Financial
# All monetary values are in `config.CURRENCY` (single-currency MVP, rule §3.10).
def financial_kpis(f: Filters) -> FinancialKPIs
def financial_revenue_waterfall(f: Filters) -> pd.DataFrame  # cols: stage (Budget/LE/Actuals), value
def financial_by_channel(f: Filters) -> pd.DataFrame         # cols: channel_code, channel_name, revenue_actual
def financial_pnl_summary(f: Filters) -> pd.DataFrame        # cols: brand, revenue, gm, promo_spend, net_revenue
def financial_ytd_progress(f: Filters) -> pd.DataFrame       # cols: brand, ytd_actual, full_year_budget, ytd_pct
    # YTD = periods in fiscal year up to f.period_to (fiscal year = calendar year for MVP).
    # full_year_budget = SUM(revenue_budget) over all 12 months of f.period_to's year.
    # ytd_pct = ytd_actual / NULLIF(full_year_budget, 0) * 100. Brand dimension joined from dim_sku.
def financial_gm_trend(f: Filters) -> pd.DataFrame           # cols: period_date, gm_pct
    # Monthly GM% across the filter window, for the KPI card sparkline (SPEC §5.3).
    # Per period: SUM(gm_actual) / NULLIF(SUM(revenue_actual), 0) * 100

# Scorecard
def scorecard_all(f: Filters) -> pd.DataFrame
    # cols: kpi_name, value, unit, rag, threshold_explainer
    #   kpi_name             — 'Forecast Accuracy' | 'Forecast Bias' | 'Days of Supply' |
    #                          'Fill Rate' | 'Revenue vs Budget' | 'GM% vs Budget'
    #   value                — the computed KPI value (float)
    #   unit                 — '%' | 'days' | 'pp' (percentage points)
    #   rag                  — 'green' | 'amber' | 'red', computed per the KPI's threshold semantics
    #   threshold_explainer  — human-readable string for tooltips/cells, e.g. '≥ 80% green; ≥ 70% amber'
    #                          or 'target 30 days ±5 green; ±10 amber'.
    # Heterogeneous threshold semantics (simple ≥, absolute |x|, band-around-target, pp-delta) live
    # inside this function; callers only render `value`, `rag`, and the `threshold_explainer` string.
    # DOS row uses the demand-weighted aggregate (see supply_kpis), compared to
    # config.RAG['dos']['target_days'] ± green_band / amber_band.

# Scenario (Phase 7)
def scenario_baseline(f: Filters) -> pd.DataFrame
    # cols: sku_code, channel_code, consensus_volume_cases
    # Per-(SKU, channel) consensus volume aggregated over the FORWARD portion of f's window —
    # i.e. period_dates strictly after the latest active fact_demand actuals period. This is what
    # scenario % adjustments are applied to.
    # consensus_volume_cases = SUM(consensus_fcst * uom_to_cases)  (UOM-normalized, rule §3.9)
    # joined to dim_sku for uom_to_cases. Used by 6_Scenario.py to compute scenario revenue:
    #   scenario_revenue = SUM(consensus_volume_cases * (1 + adjustment%) * avg_price)
    # joining this result with scenario_price_proxy on (sku_code, channel_code).
def scenario_price_proxy(f: Filters) -> pd.DataFrame
    # cols: sku_code, channel_code, avg_price
    # Window: the 3 calendar months ending at the latest period with actuals ≤ f.period_to —
    # i.e. anchored to the actuals frontier, not the calendar period_to. This keeps the
    # proxy meaningful when f.period_to extends past actuals into the forward plan.
    #   avg_price = SUM(revenue_actual) / NULLIF(SUM(actuals * uom_to_cases), 0)
    # joining fact_demand to fact_financial on (period_date, sku_code, channel_code),
    # then joining dim_sku for uom_to_cases. avg_price unit = currency per case.
    # (sku, channel) pairs with no actuals in the window are omitted; the page surfaces
    # the count as an "X SKU-channel pairs excluded (no recent revenue)" note.
    # Returns empty if no actuals exist ≤ f.period_to.

# Dimension lookups (ttl=3600)
def list_brands() -> list[str]
def list_categories() -> list[str]
def list_channels() -> list[str]
def list_regions() -> list[str]
def list_skus() -> pd.DataFrame                            # cols: sku_code, sku_name
```

**RAG values returned in DataFrames:** strings `'green' | 'amber' | 'red'`. Pages map to colors via `config.RAG_COLORS`.

### 5.3 `exports/pptx_builder.py`

```python
def build(
    modules: list[str],                         # subset of {'demand','supply','financial','scorecard','scenario'}
    filters: Filters,
    scenario_adjustments: dict | None = None,   # required iff 'scenario' in modules; shape per §4 session state
    template_path: str | None = None,           # see "template handling" note below
) -> tuple[bytes, str]:                         # (file_bytes, suggested_filename)
```

**Template handling.** The TKO MVP path is **programmatic** — slides are built from `config.TKO` design tokens. `template_path` is an optional escape hatch for the future:
- If `template_path` is `None` and `config.TKO_TEMPLATE_PATH` exists on disk → use it as a `python-pptx` slide master.
- If `template_path` is provided and exists → use that file as the slide master.
- Otherwise → fully programmatic build from `config.TKO`. No template file needed.

This resolves the SPEC §6.1 vs §7.4 tension: §7.4 (programmatic) is the default; §6.1 (template) is honored when a file is present.

**Filename format.** `S&OP_Report_YYYY-MM.pptx` (matches SPEC §6.2 — modern browsers handle `&` in Content-Disposition correctly). `YYYY-MM` resolution order:
1. `MAX(period_date)` in the active `fact_demand` version, if present
2. `filters.period_to`, formatted as YYYY-MM
3. Today's first-of-month, formatted as YYYY-MM

### 5.4 `exports/xlsx_builder.py`

```python
def build(
    modules: list[str],
    filters: Filters,
    scenario_adjustments: dict | None = None,   # required iff 'scenario' in modules
) -> tuple[bytes, str]:
```

Filename format: `S&OP_Report_YYYY-MM.xlsx` (matches SPEC §6.2). `YYYY-MM` resolved the same way as `pptx_builder` (demand max period → `filters.period_to` → today's first-of-month).

### 5.5 `config.py` — exported constants

Pages and exports read these directly. Treat as a contract; renaming a key breaks callers.

```python
STORAGE_DIR         : str        # storage/
RAW_DIR             : str        # storage/raw
TEMPLATES_DIR       : str        # storage/templates
SCHEMA_SQL_PATH     : str        # data/schema.sql
TKO_TEMPLATE_PATH   : str | None # storage/templates/tko_template.pptx if present
TKO_LOGO_PATH       : str        # assets/tko/assets/tko-wordmark.png — PNG only.
                                 # The source SVG is manually converted once (SPEC §7.4 step 5);
                                 # python-pptx cannot embed SVGs directly.

TKO                 : dict       # see SPEC §7.4 — colors{}, fonts{}, slide{}
SERIES              : dict       # semantic color mapping for chart series — see SPEC §10.3
RAG_COLORS          : dict       # {'green': ..., 'amber': ..., 'red': ...}
RAG                 : dict       # threshold config — see SPEC §5.4

CURRENCY            : str        # e.g. 'EUR' — single-currency MVP
DEFAULT_LOOKBACK_M  : int        # 12 — months subtracted from period_to in default_filters()
                                 # (yields a 13-month inclusive R13M window — SPEC §4.4)
```

**Side effect on import:** `config.py` registers the Plotly template `"tko"` via `plotly.io.templates`, and sets it as default. Importing `config` once is enough to brand every chart.

### 5.6 `data/session_db.py` — per-session DuckDB management

The deployed demo is multi-tenant: each Streamlit session gets its own DuckDB file under `storage/sessions/`, so two visitors uploading CSVs at the same time never overwrite each other. **All DuckDB connections must open against `session_db.get_db_path()` — never against a hard-coded path.**

```python
SESSIONS_DIR        : str        # storage/sessions/

get_db_path() -> str
    # Per-Streamlit-session DuckDB file path.
    # Caches a UUID in st.session_state["session_db_id"] on first call.
    # Falls back to storage/sop.duckdb when called outside a Streamlit session
    # (unit tests, ad-hoc scripts).

seed_if_empty() -> None
    # Idempotent. On first call inside a Streamlit session with no prior
    # uploads, ingests test_data/*.csv via ingest.load() so dashboards render
    # immediately. Guarded by st.session_state["session_db_seeded"].
    # Also populates st.session_state["data_versions"] for the home-page chip.

cleanup_stale_sessions(max_age_hours: float = 6) -> None
    # Best-effort prune of orphaned files in SESSIONS_DIR.
    # Runs at most once per Python process. Streamlit Cloud's ephemeral disk
    # makes this mostly a local-dev hygiene measure.
```

**Session-state keys owned here** (extends CONTRACTS §4):

| Key | Type | Set by | Read by |
|-----|------|--------|---------|
| `session_db_id` | `str` (hex UUID) | `get_db_path()` on first access | `get_db_path()` |
| `session_db_seeded` | `bool` | `seed_if_empty()` | `seed_if_empty()` |

---

## 6. Phase Contract Map

What each phase **introduces** (writes new contracts) vs **consumes** (calls existing).

| Phase | Introduces | Consumes |
|-------|------------|----------|
| 1 — Skeleton + Upload | `config.py` constants, `schema.sql`, all of `ingest.py`, session keys `data_versions` / `last_upload_status` / `column_mapping_template` | — |
| 2 — Demand Dashboard | `Filters`, `default_filters`, `DemandKPIs`, `current_data_version`, all `demand_*` + `list_*` funcs, session key `filters` | `ingest.list_uploads`, config constants |
| 3 — Excel Export | `xlsx_builder.build` with `modules=['demand']` only — Cover sheet + Demand sheet. (Scorecard sheet deferred to Phase 5 when `scorecard_all` exists.) | `Filters`, demand queries, `config.TKO` |
| 4 — Supply + Financial | `SupplyKPIs`, `FinancialKPIs`, all `supply_*` + `financial_*` funcs (including `supply_capacity_utilization`, `supply_fill_rate_trend`, `financial_ytd_progress`, `financial_gm_trend`); extends `xlsx_builder.build` to accept `'supply'` and `'financial'` modules | `Filters`, `current_data_version` |
| 5 — Scorecard + PPTX | `scorecard_all`, `pptx_builder.build` accepting modules ⊆ {demand, supply, financial, scorecard}; retrofits Scorecard sheet into `xlsx_builder.build`. Both exporters raise `NotImplementedError` on `'scenario'` until Phase 7. | All query funcs, `config.RAG`, `config.TKO` |
| 6 — TKO Polish | (none — refinement) | `pptx_builder`, `config.TKO`, `assets/tko/` |
| 7 — Scenario + Exec | `scenario_baseline`, `scenario_price_proxy`; session keys `scenario_adjustments` / `exec_filters_frozen`; implements `'scenario'` handling in both `pptx_builder.build` and `xlsx_builder.build` (signatures already declared in Phase 5) | All queries, both exporters |

**Read this column when starting a chunk:** "Introduces" tells you what you're allowed to add. "Consumes" tells you what should already exist and you must not modify.

---

## 7. Resolved Contract Decisions

All four ambiguities from the initial draft are resolved below. Each decision is now load-bearing — changing one requires a CONTRACTS.md update, not an inline code change.

### 7.1 Filters defaults (Phase 2)
- **`period_to`** = first-of-month of `MAX(period_date)` in active `fact_demand`. Fallback: first-of-current-month when no demand has been uploaded.
- **`period_from`** = `period_to` − 12 months. Yields 13 months inclusive (R13M, SPEC §4.4).
- **Dim list fields** default empty; empty = "no filter" per cross-cutting rule §3.7.
- Implemented as `query.default_filters() -> Filters` (not cached — depends on `current_data_version('demand')`).
- **Why demand anchors:** SPEC frames the platform around the 5-gate cycle, and Demand Review is gate 1 — its data is always present before Supply/Financial. Anchoring elsewhere creates a chicken-and-egg on first run.

### 7.2 DOS demand join (Phase 4)
- **Same-period join:** `fact_supply LEFT JOIN fact_demand ON (period_date, sku_code)`. Same-period `consensus_fcst` is the divisor in `inventory_qty / (consensus_fcst / 30.44)`.
- If no matching demand row → DOS = NULL → excluded from aggregates (no zero-division, no fake values).
- **Why same-period and not rolling-forward:** matches SPEC §8.2 formula literally; the rolling-forward refinement (typically next-3-month avg demand) is industry standard but introduces a window-config parameter and edge cases at horizon boundaries — defer until requested.

### 7.3 Scorecard DOS aggregation (Phase 5)
- **Demand-weighted aggregate** at the filter scope, latest period in window, with UOM normalization (rule §3.9):
  ```
  dos_agg = SUM(inventory_qty * uom_to_cases)
            / NULLIF(SUM(consensus_fcst * uom_to_cases) / 30.44, 0)
  ```
- **Not** a simple mean across SKUs (over-weights small SKUs) and **not** worst-case min (hides systemic position from execs).
- RAG comparison: `config.RAG['dos']['target_days']` ± `green_band` / `amber_band`.
- Worst-N SKUs already surface on the Supply page via `supply_dos_by_sku` — execs drill there for detail.

### 7.4 Scenario price proxy (Phase 7)
- **Window:** the 3 calendar months **ending at the latest actuals period ≤ `Filters.period_to`** — i.e. anchored to the actuals frontier, not the calendar `period_to`. If the user's window extends past the actuals frontier into forward planning months, the proxy still uses the 3 actuals months immediately preceding (and including) that frontier.
- Returns **empty** if no actuals exist at or before `period_to` in the active demand version.
- **Grain:** per `(sku_code, channel_code)` — matches `fact_financial` PK.
- **Formula** (with UOM normalization per rule §3.9):
  ```
  avg_price = SUM(revenue_actual) / NULLIF(SUM(actuals * uom_to_cases), 0)   -- unit: currency per case
  ```
  joining `fact_demand` ↔ `fact_financial` on `(period_date, sku_code, channel_code)`, then `dim_sku` for `uom_to_cases`.
- `(sku, channel)` pairs with no actuals in the window are **omitted** from the result; the Scenario page renders a banner: *"N SKU-channel pairs excluded — no revenue history in trailing 3 months."*
- Scenario revenue per scenario = `SUM(scenario_volume_in_cases × avg_price)` joined back on `(sku_code, channel_code)`.
- **Why trailing 3M, not 12M or latest:** 12M smooths through too much seasonality and goes stale; latest-period alone is volatile (promo months distort). Three months is the standard CPG demand-planning recency window.
- **Why anchor to actuals frontier, not literal `period_to`:** users routinely set `period_to` to the planning horizon end (forward months) to populate `scenario_baseline`. A literal-`period_to` proxy window would land entirely in forward months, find no revenue, and report every pair excluded. Anchoring to the actuals frontier keeps the proxy a function of recent history regardless of how far forward `period_to` reaches.

---

*Last updated: 2026-05-21 (Phase 7 — Scenario Planner + Exec Summary: `query.scenario_baseline` + `query.scenario_price_proxy` implemented per §5.2 / §7.4; `xlsx_builder.build` now writes a Scenario sheet (volume × trailing-3M avg price per scenario) and `pptx_builder.build` adds Slide 14 (3-scenario table + revenue bar chart) — both accept `scenario_adjustments` and default to `{base:0, upside:+5, downside:-5}` when omitted; `chart_fallback._render_scenario_revenue` added as the matplotlib kaleido fallback; `pages/6_Scenario.py` exposes three sliders writing to `st.session_state["scenario_adjustments"]`; `pages/7_Exec_Summary.py` provides the locked exec view with `exec_filters_frozen` session key and the bundled full-report export. **Phase 7 follow-up (same session):** to make the test-data Scenario flow exercise real numbers, `generate_test_data.py` extended to 16 months (Apr 2025 → Jul 2026) with May–Jul 2026 as a demand-only forward planning horizon; `ingest._SCHEMA['demand']` moved `actuals` from required → optional (consensus-only rows are a valid planning shape); `scenario_price_proxy` window re-anchored from "trailing 3 months ending at `f.period_to`" to "trailing 3 months ending at the latest actuals period ≤ `f.period_to`" so the proxy keeps working when users widen `period_to` into the forward horizon. Phase 6 footer prior: TKO polish — `assets/tko/` unzipped, `build_wordmark.py` regenerates `tko-wordmark.png` (1040×320), cover slide embeds wordmark via `pptx_builder._add_tko_wordmark`, Plotly `tko` template tuned — `acid` signal-only, axis ticks on `mist`, titles on `ice_dim`, hoverlabel on `graphite`.*
