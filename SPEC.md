# S&OP Reporting Platform — E2E Specification

**Role context:** This platform is designed by and for an S&OP Manager / Business Analyst in a CPG company.
**Primary goal:** Turn raw CSV data exports (from ERP, TMS, demand planning tools) into interactive S&OP dashboards with one-click PPTX and Excel exports.
**Build approach:** Vibe-coded with Claude — AI-maintainable as a top architectural constraint.
**App runs on:** `http://localhost:8501` (Streamlit 1.36 has 8501 hardcoded in its frontend bundle — earlier drafts said 3000 but that port doesn't work with the pinned stack)
**PPTX export template:** TKO brand template (to be uploaded; see §7.4)

---

## Table of Contents

0. [Maintainability Principles (Vibe Coding)](#0-maintainability-principles-vibe-coding)
1. [Business Context](#1-business-context)
2. [S&OP Process & Meeting Cadence](#2-sop-process--meeting-cadence)
3. [Functional Requirements](#3-functional-requirements)
4. [Data Requirements](#4-data-requirements)
5. [Dashboard Modules](#5-dashboard-modules)
6. [Export Requirements](#6-export-requirements)
7. [Architecture](#7-architecture)
8. [Data Model](#8-data-model)
9. [Internal API (Python function contracts)](#9-internal-api-python-function-contracts)
10. [UI/UX Specification](#10-uiux-specification)
11. [Non-Functional Requirements](#11-non-functional-requirements)
12. [Build Phases](#12-build-phases)
13. [Glossary](#13-glossary)

---

## 0. Maintainability Principles (Vibe Coding)

This solution is built and maintained iteratively with Claude. Every design choice in this spec follows these rules — if a future change violates them, push back before implementing.

### Rule 1 — One language only
**Python only.** No JavaScript, no TypeScript, no separate frontend codebase. A vibe-coded app split across two languages is twice as hard to keep coherent in one Claude session.

### Rule 2 — One process, one repo
No microservices, no separate API server. One Python app that serves UI, queries data, and generates exports. No Docker required to run.

### Rule 3 — Boring, well-known libraries
Prefer libraries Claude has seen thousands of times: `pandas`, `streamlit`, `duckdb`, `plotly`, `python-pptx`, `openpyxl`. Skip anything that needs custom build pipelines or rare bindings.

### Rule 4 — Flat file structure, < 15 files at first
Keep modules small and few. A single `app.py` can grow into a `pages/` directory only when navigation actually splits. No premature folder hierarchies.

### Rule 5 — No background jobs, no async, no queues
Exports run synchronously. If a 30-second wait is acceptable, do not add Celery, Redis, BackgroundTasks, threads, or asyncio. They are 10× more code and 100× more failure modes.

### Rule 6 — No auth, no user accounts (MVP)
This is a desktop/local tool for the S&OP manager running it. One user, no login, no roles. Skip JWT, sessions, RBAC, audit logs until there is a real multi-user requirement.

### Rule 7 — State lives in two places only
- **Disk** (uploaded CSVs and DuckDB file) for durable state.
- **`st.session_state`** for ephemeral UI state (filters, current page, last upload).
No Redis, no separate state store, no cookies, no localStorage hacks.

### Rule 8 — Inline > abstract
Two pages with similar charts? Repeat the code. Don't extract a chart factory until you have 3 real usages. Vibe coding rewards explicit, scannable code over DRY cleverness.

### Rule 9 — Configuration in one file
All thresholds, paths, brand colors, RAG cutoffs in a single `config.py`. No `.env` files, no YAML, no Settings classes. Just module-level constants.

### Rule 10 — Pinned dependencies
`requirements.txt` with `==` pins, not `>=`. A vibe-coded app should run identically in 6 months. A breaking minor version update is a maintenance disaster waiting to happen.

### What this means for the stack
- **Streamlit** replaces React + FastAPI (one process, Python only, all UI primitives built-in).
- **Plotly** replaces Recharts + Nivo (one chart library, exports cleanly to PNG for PPTX).
- **DuckDB** stays (in-process, no separate server, SQL on CSVs is its specialty).
- **No JWT, no Zustand, no Tailwind, no shadcn, no Celery, no Vite.**

### What I keep from the original spec
- All business requirements (§1–§6) — unchanged.
- All data model fields, KPI formulas, RAG thresholds — unchanged.
- All export contents (PPTX slide structure, Excel sheet structure) — unchanged.
- Only the *delivery mechanism* shrinks.

---

## 1. Business Context

### 1.1 Industry
Consumer Packaged Goods (CPG). Companies in this space manage large SKU portfolios, multi-channel distribution (retail, e-commerce, foodservice), and highly seasonal demand patterns.

### 1.2 The Problem
S&OP reporting in CPG is typically done manually in Excel. Data is extracted from multiple systems (SAP, Kinaxis, Nielsen, Circana, etc.) into CSVs, then pasted into Excel templates, then copy-pasted into PowerPoint decks — a process that takes 2–3 days per month and is error-prone.

### 1.3 The Solution
A web application that:
- Accepts CSV uploads from any source system
- Automatically structures data into S&OP-relevant views
- Renders interactive dashboards for each S&OP review gate
- Exports professional PowerPoint decks and Excel reports in one click

### 1.4 Primary Users

| Role | Usage |
|------|-------|
| S&OP Manager / BA | Uploads data, configures dashboards, exports decks |
| Demand Planner | Reviews demand accuracy, adjusts forecasts |
| Supply Planner | Reviews inventory and supply constraints |
| Finance BP | Reviews revenue and margin performance |
| S&OP Executive | Views consensus plan, approves scenarios |

---

## 2. S&OP Process & Meeting Cadence

The platform supports the standard 5-gate monthly S&OP cycle:

```
Week 1         Week 2         Week 3         Week 4
   │              │              │              │
[Data Close]  [Demand       [Supply        [Pre-S&OP]    [Exec S&OP]
[Upload CSVs]  Review]        Review]
```

### 2.1 Gate 1 — Demand Review
- **Who:** Demand Planning team + Commercial leads
- **Dashboard:** Demand Review Module
- **Key output:** Validated statistical + consensus forecast by SKU/channel/region

### 2.2 Gate 2 — Supply Review
- **Who:** Supply Chain, Procurement, Manufacturing
- **Dashboard:** Supply Review Module
- **Key output:** Constrained supply plan, highlighted gaps

### 2.3 Gate 3 — Financial Review
- **Who:** Finance BPs, S&OP Manager
- **Dashboard:** Financial Review Module
- **Key output:** Revenue/margin projection vs budget

### 2.4 Gate 4 — Pre-S&OP
- **Who:** Cross-functional leads
- **Dashboard:** Scenario Planner + KPI Scorecard
- **Key output:** 3-scenario recommendation for executive decision

### 2.5 Gate 5 — Executive S&OP
- **Who:** C-Suite, VP level
- **Dashboard:** Executive Summary (read-only, locked view)
- **Key output:** Approved consensus plan → triggers PPTX export

---

## 3. Functional Requirements

### 3.1 Data Ingestion

#### FR-ING-01: CSV Upload
- User can upload one or multiple CSV files via drag-and-drop or file picker
- Accepted encodings: UTF-8, UTF-16, Latin-1 (auto-detected)
- Max file size: 500 MB per file
- Max files per session upload: 20
- Files can be uploaded at any time; they do not reset the existing dataset unless the user explicitly replaces

#### FR-ING-02: Schema Auto-Detection
- System reads the first 1,000 rows to infer:
  - Column names (from header row)
  - Data types: date, number (integer/float), text, boolean
  - Date formats: auto-parsed (ISO 8601, DD/MM/YYYY, MM/DD/YYYY, YYYYMM)
- Detection result shown to user before commit

#### FR-ING-03: Column Mapping
- User maps detected columns to S&OP semantic fields (e.g., "Fcst_Vol" → `statistical_fcst`)
- **One saved mapping per domain** (demand, supply, financial, …) is persisted in `st.session_state["column_mapping_template"]` and applied automatically on the next upload to the same domain
- User can override the auto-applied mapping before confirming
- A full "mapping template library" with named templates (e.g. "SAP", "Nielsen") is out of MVP scope — one default per domain is enough for a single-user tool

#### FR-ING-04: Data Validation
- Validation runs after mapping, before ingestion
- Checks performed:
  - Required fields (per §4.1) are not null
  - Date values are parseable
  - Numeric fields contain no text (report exceptions, allow user to ignore or block)
  - SKU codes exist in `dim_sku` (warn if new, block if blank)
  - Channel/region/plant codes exist in their dim tables (warn if new)
  - **Duplicate rows on the domain's primary key** — the PK is domain-specific:
    - demand → `(period_date, sku_code, channel_code, region_code)`
    - supply → `(period_date, sku_code, plant_code)`
    - financial → `(period_date, sku_code, channel_code)`
    - dimensions → the dim code (e.g. `sku_code`)
    - On duplicate: flag and deduplicate (keep last occurrence in the file)
- Validation report shown as a summary with row-level drill-down
- User can: fix-and-re-upload, continue-ignoring-warnings (errors always block), or cancel

#### FR-ING-05: Versioning
- Each upload creates a new `data_version` integer (auto-incremented) and a row in `upload_log` with `status = 'active'`.
- Previous active version for the same domain is automatically marked `status = 'superseded'` — but the rows remain in the fact table for audit.
- The current active version is shown in the dashboard header chip.
- A read-only "Upload history" expander on the Upload page lists past uploads (filename, domain, rows, timestamp, status).
- **Manual rollback UI is deferred.** Until built, recovery from a bad upload is: re-upload the previous good CSV (creates a new version with the old data). This is acceptable because the original CSVs are preserved in `storage/raw/`.

#### FR-ING-06: Append vs Replace
- On upload, user selects:
  - **Replace:** overwrites the existing dataset for that domain
  - **Append:** adds new rows; deduplicates on primary key

### 3.2 Dashboard Interactions

#### FR-DASH-01: Global Filters
- Filter bar always visible at top of every dashboard page
- Filter dimensions: Time Period (month range), Brand, Category, SKU, Channel, Region
- Filters apply across all charts on the active page simultaneously
- Filter state is preserved when navigating between pages in the same session

#### FR-DASH-02: Chart Interactions
- Hover tooltip on all charts showing exact values (Plotly default)
- Plotly modebar (top-right of each chart) provides built-in PNG download
- For per-chart raw-data export, each chart has a small `st.download_button` underneath labeled "⬇ Data (CSV)"
- Charts are responsive to window resize (Streamlit + `use_container_width=True`)

#### FR-DASH-03: Data Freshness Indicator
- Every dashboard page shows in the header: "Data as of [upload timestamp] | Version [n]"

#### Deferred (post-MVP)
- **Saved Views** — named filter+page bookmarks. Requires a UI for save/load and an extra DuckDB table. Streamlit session_state already preserves filters within a session, which covers 80% of the need.
- **Cross-filtering** — click a bar to filter other charts. Streamlit's rerun model makes this awkward (would need callback wiring). Defer until a concrete user request.

### 3.3 Export

#### FR-EXP-01: PPTX Export
- "Export PPTX" button at the bottom of each dashboard page exports **that page only** (one or two slides)
- "Export full report" button on the Exec Summary page exports **all modules** as one bundled deck
- Each dashboard page maps to 1–2 slides: KPI summary slide + chart-grid slide
- Charts are inserted as **PNG images** rendered server-side from Plotly via kaleido (no native PPT chart objects — see §6.1)
- Tables are inserted as native PPT table objects (small enough that python-pptx handles them well)
- Slides are constructed programmatically using **TKO design tokens** (colors, fonts) defined in `config.py` — there is no `.pptx` master file (see §7.4)
- Cover slide auto-populated: report title, S&OP cycle month, export date, active filter summary
- Section divider slides separate gate sections (Demand / Supply / Financial / Scenario)
- Export is **synchronous** (no job queue) — typical wall time 10–20 s; `st.spinner` shows progress; file streams back to the browser via `st.download_button`

#### FR-EXP-02: Excel Export
- "Export XLSX" button at the bottom of each dashboard page exports **that page** as a workbook
- "Export full report" on the Exec Summary page produces one workbook covering all modules
- Per-chart raw-data download is offered as a **CSV** button beneath each chart (see §3.2 FR-DASH-02) — not per-chart Excel
- Workbook structure: one sheet per module, with summary at top + flat data below in the same sheet (see §6.2)
- Active filters applied to the exported data
- Numeric formatting preserved (thousands separator, % for ratios, currency symbol)

### 3.4 User Management

**Out of MVP scope.** Single-user local app. No login, no roles, no audit log. Revisit only if the tool is hosted for multiple users.

---

## 4. Data Requirements

### 4.1 Supported Input Domains

Each domain corresponds to a CSV type the user uploads. Multiple domains can be uploaded separately.

> **Note on column names:** The CSV header names are arbitrary — users map them to canonical fields in the Upload UI (see §3.1 FR-ING-03). The canonical field names below (`sku_code`, `channel_code`, `plant_code`, etc.) are what the column mapper targets.

| Domain | Description | Minimum Required Columns (canonical names) | Optional Columns |
|--------|-------------|---------------------------------------------|------------------|
| `demand` | Forecast vs actuals by SKU/channel/period | `period_date`, `sku_code`, `channel_code`, `statistical_fcst`, `consensus_fcst`, `actuals` | `region_code` |
| `supply` | Inventory and production data | `period_date`, `sku_code`, `plant_code`, `inventory_qty`, `production_plan`, `production_actual` | `capacity_plan`, `orders_requested`, `orders_delivered` |
| `financial` | Revenue and margin data | `period_date`, `sku_code`, `channel_code`, `revenue_actual`, `revenue_budget`, `revenue_le`, `gm_actual` | `gm_budget`, `promo_spend_actual`, `promo_spend_budget`, `currency_code` |
| `sku_master` | SKU dimension table | `sku_code`, `sku_name`, `brand`, `category`, `uom` | `subcategory`, `uom_to_cases` |
| `channel_master` | Channel dimension | `channel_code`, `channel_name` | `channel_type` |
| `region_master` | Region / geography dimension | `region_code`, `region_name`, `country` | `cluster` |
| `plant_master` | Plant / manufacturing site dimension | `plant_code`, `plant_name` | `country` |

> **Charts that depend on optional columns** are rendered only when those columns are populated. Specifically: Fill Rate (needs `orders_*`), Capacity Utilization (needs `capacity_plan`), Promo / Net Revenue (needs `promo_spend_actual`). The dashboard shows an inline "no data" hint where applicable.

### 4.2 Date Granularity
- All fact data: **monthly** granularity only (MVP)
- Date column must be parseable to year-month (YYYYMM, YYYY-MM-01, etc.)
- System normalizes all dates to first-of-month internally
- Weekly granularity is deferred — would require adding a `granularity` field to `Filters` and parallel weekly fact tables

### 4.3 Units of Measure
- Volume: cases, units, kg, tonnes — stored in the original UOM provided by the CSV
- Conversion to cases happens at **query time** using `dim_sku.uom_to_cases` (no duplicate `*_cases` columns in fact tables)
- Revenue: any currency; currency code stored per row; dashboard shows single currency (user selects FX basis in `config.py` for MVP — multi-currency FX conversion deferred)
- Forecasts and actuals must arrive in the same UOM (validation enforces this)

### 4.4 Horizon
- Historical data: up to 36 months back
- Forward-looking horizon: up to 18 months
- Rolling 13-month view is the default dashboard display window

---

## 5. Dashboard Modules

### 5.1 Demand Review Dashboard

**Purpose:** Validate the forecast and understand volume trends by SKU, channel, region.

#### Charts & KPIs

| Visual | Type | Metric |
|--------|------|--------|
| Forecast Accuracy | KPI card + trend sparkline | MAPE = mean(\|Actual - Forecast\| / Actual) × 100 |
| Forecast Bias | KPI card + trend sparkline | Bias = mean((Actual - Forecast) / Actual) × 100 |
| Volume Bridge | Waterfall chart | Prior Period Actuals → Statistical Fcst → Consensus Fcst → Current Actuals |
| Forecast vs Actuals | Line + bar combo | Monthly actuals (bar) overlaid with statistical + consensus forecast (lines) |
| SKU-level Accuracy | Table with heatmap | MAPE per SKU, sortable, color-coded RAG |
| Channel Mix | 100% stacked bar | Volume share by channel over time |
| Over/Under Forecast | Scatter plot | SKUs plotted by MAPE vs volume; quadrant view |

#### Calculations
```
MAPE = AVG(ABS(actuals - consensus_fcst) / actuals) * 100
Bias = AVG((actuals - consensus_fcst) / actuals) * 100
FA% = 100 - MAPE  (Forecast Accuracy)
Statistical vs Consensus Delta = consensus_fcst - statistical_fcst
```

#### Default view
- Period: Rolling 13 months (R13M)
- Grouping: Brand > Category > SKU
- (Prior-year benchmark line on time-series charts is a candidate enhancement — not in MVP.)

---

### 5.2 Supply Review Dashboard

**Purpose:** Identify supply constraints and inventory risks.

#### Charts & KPIs

| Visual | Type | Metric |
|--------|------|--------|
| Days of Supply (DOS) | KPI card + gauge | inventory_qty / avg_daily_demand |
| Fill Rate | KPI card + trend | orders_delivered / orders_requested |
| Inventory by SKU | Horizontal bar | Current DOS vs target DOS (colored by risk) |
| Production Adherence | Clustered bar | production_actual vs production_plan by plant |
| Capacity Utilization | Bullet chart | production_actual vs capacity_plan (planned capacity) per plant |
| Inventory Cover Map | Heat map | SKU × Month: green (safe) / yellow (risk) / red (stockout) |
| Supply Gaps | Table | SKUs with projected stockout in next 3 months |

#### Calculations
```
DOS = inventory_qty / (consensus_fcst / 30.44)        -- computed at query time, joining fact_supply to fact_demand
Fill Rate = SUM(orders_delivered) / SUM(orders_requested)
Production Adherence = SUM(production_actual) / SUM(production_plan)
Capacity Utilization = SUM(production_actual) / SUM(capacity_plan)
```

---

### 5.3 Financial Review Dashboard

**Purpose:** Track revenue and margin vs budget and latest estimate (LE).

#### Charts & KPIs

| Visual | Type | Metric |
|--------|------|--------|
| Revenue Actuals vs Budget | KPI card with delta % | revenue_actual vs revenue_budget |
| Revenue vs LE | KPI card with delta % | revenue_actual vs revenue_le |
| GM% | KPI card + trend | gross_margin / revenue * 100 |
| Revenue Waterfall | Waterfall | Budget → LE → Actuals (delta per stage, in absolute revenue) |
| Revenue by Channel | Treemap | Revenue share and absolute by channel |
| P&L Summary | Table | Revenue / GM / Promo spend / Net revenue by brand |
| YTD vs Full-Year Budget | Progress bar set | One bar per brand |

#### Calculations
```
GM%                     = SUM(gm_actual) / NULLIF(SUM(revenue_actual), 0) * 100
Revenue vs Budget Δ%    = (SUM(revenue_actual) - SUM(revenue_budget)) / NULLIF(SUM(revenue_budget), 0) * 100
Revenue vs LE Δ%        = (SUM(revenue_actual) - SUM(revenue_le))     / NULLIF(SUM(revenue_le), 0)     * 100
Net Revenue             = revenue_actual - promo_spend_actual
```

> **Note:** A price/volume/mix decomposition of the revenue waterfall would require joining `fact_demand` (volume) to `fact_financial` (revenue) and inferring unit prices. This is non-trivial and is deferred to post-MVP. The Budget → LE → Actuals waterfall is the MVP version.

---

### 5.4 KPI Scorecard

**Purpose:** Single-page executive view of all tier-1 KPIs with RAG status.

| KPI | Frequency | Green | Amber | Red |
|-----|-----------|-------|-------|-----|
| Forecast Accuracy | Monthly | ≥ 80% | 70–79% | < 70% |
| Forecast Bias | Monthly | ±5% | ±5–10% | > ±10% |
| Days of Supply | Monthly | Target ± 5 days | Target ± 5–10 days | > Target ± 10 days |
| Fill Rate | Monthly | ≥ 98% | 95–97% | < 95% |
| Revenue vs Budget | Monthly | ≥ 98% of budget | 95–97% | < 95% |
| GM% vs Budget | Monthly | ≥ budget GM% | within 1pp below | > 1pp below |

**Thresholds and the DOS target live in `config.py`** as a single source of truth — no per-SKU targets in MVP:

```python
RAG = {
    "fa":  {"green": 80, "amber": 70},          # MAPE-based FA%
    "bias": {"green_abs": 5, "amber_abs": 10},
    "dos": {"target_days": 30, "green_band": 5, "amber_band": 10},
    "fill_rate": {"green": 98, "amber": 95},
    "rev_vs_budget": {"green": 98, "amber": 95},
    "gm_vs_budget_pp": {"green": 0, "amber": -1},
}
```

A per-SKU `dos_target_days` column on `dim_sku` is a candidate post-MVP enhancement.

---

### 5.5 Scenario Planner

**Purpose:** Side-by-side comparison of Base / Upside / Downside demand scenarios.

- User defines three fixed scenarios — **Base**, **Upside**, **Downside** — by entering a single `% volume adjustment vs consensus` for each, applied uniformly (Phase 7 MVP). Per-brand or per-SKU group adjustments are deferred.
- Dashboard shows volume, revenue (volume × current avg price proxy), and DOS implications per scenario, side-by-side.
- Scenario inputs live in `st.session_state["scenario_adjustments"]` — **in-session only, not persisted** across restarts. Persisted scenarios would require a `dim_scenario` table and is deferred until concretely needed.
- Scenario outputs can be exported as a dedicated section in the PPTX deck.

---

### 5.6 Executive Summary

**Purpose:** One-page view for the Exec S&OP meeting.

- Curated single-page: **4 KPI cards** (Forecast Accuracy, Fill Rate, Revenue vs Budget, GM%) + a volume trend chart + a revenue-vs-budget chart + the 3-scenario comparison table.
- **Filter freeze:** A "🔒 Freeze filters for Exec" button locks the current filter selection and hides the filter bar (replaced by a read-only badge showing the frozen scope). Clicking "Unfreeze" restores filter controls. The frozen state lives in `st.session_state["exec_filters_frozen"]`.
- No role gating — the freeze is an honor-system convenience, not a security boundary (single-user app, see §0 Rule 6).
- Large, high-contrast visuals optimized for projector display.
- "Export full report" button on this page produces the bundled PPTX (all modules).

---

## 6. Export Requirements

### 6.1 PPTX Export — Detailed Spec

#### Slide Structure (default)
```
Slide 1  — Cover (title, cycle month, date, logo)
Slide 2  — Agenda / Table of Contents
Slide 3  — KPI Scorecard summary
Slide 4  — [Divider] Demand Review
Slides 5–7 — Demand Review charts (2 charts per slide)
Slide 8  — [Divider] Supply Review
Slides 9–10 — Supply Review charts
Slide 11 — [Divider] Financial Review
Slides 12–13 — Financial Review charts
Slide 14 — Scenario Comparison
Slide 15 — Appendix / Data notes
```

#### Technical notes
- Library: `python-pptx`
- Charts: each Plotly figure is rendered to PNG server-side via `kaleido` (`fig.to_image(format="png", width=1200, height=600, scale=2)`) and inserted as a picture on the slide. The underlying DataFrame is dumped to the slide's notes pane as a CSV string for traceability.
- Native PPT chart objects are explicitly **not** used — PNG-only keeps the builder code short and reliable.
- Tables: native PPT table objects (small enough that python-pptx handles them well).
- Template file: `.pptx` file used as slide master; `python-pptx` applies the template's theme, fonts, colors via layouts.
- TKO template: stored at `storage/templates/tko_template.pptx`; applied to all exports when present, fallback to a built-in neutral template otherwise.

### 6.2 Excel Export — Detailed Spec

#### File Structure

Sheets are created conditionally based on which modules are requested + which data is loaded. Each module sheet contains both a styled summary section (top) and the pivot-ready flat data (below) — no duplicate "Data_*" sheets.

```
Workbook: S&OP_Report_YYYY-MM.xlsx

Always present:
  Sheet: Cover         — title, cycle month, export date, active filter summary, data version
  Sheet: Scorecard     — KPI table with RAG conditional formatting

Conditional on requested modules + loaded data:
  Sheet: Demand        — top: KPI summary + chart-source tables (styled, frozen header)
                       — below (separated by blank rows): flat fact_demand rows (pivot-ready)
  Sheet: Supply        — same pattern: summary + flat data
  Sheet: Financial     — same pattern: summary + flat data
  Sheet: Scenario      — only included from Phase 7 onward, when scenario_adjustments is set
```

#### Formatting rules
- Header rows: bold, TKO `--cold-slate` (#1E2235) background, `--ice-white` (#E8EEFF) text
- Numbers: thousands separator, 1 decimal place for volumes, 2 for financials
- Percentages: `0.0%` format
- RAG cells: conditional formatting using TKO semantic mapping — `acid` (#C8E000) for green, `#F59E0B` for amber, `#EF4444` for red (note: TKO has no native amber/red, so we fall back to neutral RAG tones for those)
- Frozen panes on the summary header row (row 1) + column A
- Library: `openpyxl`

---

## 7. Architecture

### 7.1 Technology Stack (Vibe-Coded, Python-Only)

| Layer | Technology | Pinned Version | Purpose |
|-------|-----------|----------------|---------|
| UI framework | Streamlit | 1.36.0 | Pages, widgets, file upload, layout — all built-in |
| Data engine | DuckDB | 1.0.0 | In-process SQL on CSV/Parquet |
| DataFrames | pandas | 2.2.2 | CSV parsing, schema detection, transforms |
| Charts | Plotly | 5.22.0 | Interactive charts, exports cleanly to PNG |
| PPTX export | python-pptx | 0.6.23 | PowerPoint generation |
| Excel export | openpyxl | 3.1.4 | Excel generation with formatting |
| Image export (for PPTX) | kaleido | 0.2.1 | Plotly → static PNG renderer |
| Language | Python | 3.11 | Single language for the whole app |

**Not used:** React, TypeScript, FastAPI, Vite, Zustand, Tailwind, shadcn, Celery, Redis, Docker, Nginx, JWT.

### 7.2 Why Streamlit (and not React + FastAPI)

| Concern | React + FastAPI | Streamlit |
|---------|-----------------|-----------|
| Languages to maintain | 2 (Python + TS) | 1 (Python) |
| Files for a new page | ~5 (route, component, hook, API endpoint, type) | 1 |
| Lines to add a chart | ~50 | ~5 |
| Build step | Yes (Vite) | No |
| Cross-origin headaches | Yes | No |
| AI context cost to understand the app | High | Low |
| Polish ceiling | Higher | Sufficient for S&OP review meetings |

For an internal analytics tool maintained by one person + Claude, Streamlit wins decisively.

### 7.3 File Structure (target for MVP)

```
S&OP Dashboard/
├── app.py                 # Streamlit entry point — title + data-version chip
├── config.py              # All constants: paths, RAG thresholds, brand colors
├── .streamlit/
│   └── config.toml        # Sets port = 3000 + theme defaults
├── data/
│   ├── ingest.py          # CSV upload, schema detect, validation, load to DuckDB
│   ├── query.py           # Reusable DuckDB query functions per module + Filters dataclass
│   └── schema.sql         # DuckDB DDL (all fact + dim tables)
├── pages/
│   ├── 1_Upload.py        # Upload UI + column mapper + validation report
│   ├── 2_Demand.py        # Demand Review dashboard
│   ├── 3_Supply.py        # Supply Review dashboard
│   ├── 4_Financial.py     # Financial Review dashboard
│   ├── 5_Scorecard.py     # KPI Scorecard
│   ├── 6_Scenario.py      # Scenario Planner
│   └── 7_Exec_Summary.py  # Executive Summary (locked view)
├── exports/
│   ├── pptx_builder.py    # python-pptx slide assembly (PNG via kaleido)
│   └── xlsx_builder.py    # openpyxl workbook assembly
├── storage/                # Created at first run; gitignored
│   ├── sop.duckdb          # The DuckDB database file
│   ├── raw/                # Original uploaded CSVs (audit trail)
│   ├── templates/          # tko_template.pptx (manually copied in)
│   └── exports/            # Generated PPTX and XLSX files (optional persistence)
├── requirements.txt        # Pinned dependencies
├── README.md               # How to run
└── SPEC.md                 # This file
```

**File count target: 12 Python files for the full MVP.** If a new feature needs a 13th file, ask whether it can go in an existing one.

### 7.4 TKO Brand Integration

**Reality check:** "TKO" is the personal-brand **design system** of the project owner (Tatiana Koriakina · "AI in Supply Chain"). It is delivered as a zip (`TKO Design System.zip`) containing CSS design tokens, Google fonts references, SVG logos, and HTML slide examples — **not** a PowerPoint master file. This means `pptx_builder.py` constructs slides programmatically using these tokens, rather than opening a `.pptx` template.

#### What ships in the project

```
TKO Design System.zip               (kept at project root; do not commit unzipped if large)
  ├── colors_and_type.css           — CSS variables for colors, fonts, type scale
  ├── assets/
  │   ├── tko-monogram.svg          — square logo
  │   ├── tko-wordmark.svg          — horizontal logo
  │   └── signal-token.svg          — accent decoration
  └── slides/slides.css             — reference layouts (cover / stat / bullets / quote / compare)
```

#### How the app uses it

1. **One-time extraction** (manual or first-run script): unzip into `assets/tko/` (gitignored). The `pptx_builder.py` reads `assets/tko/assets/tko-wordmark.svg` for the cover logo.
2. **Design tokens are hardcoded in `config.py`** — not parsed from CSS at runtime. A constant block:

   ```python
   # config.py — TKO design tokens (extracted from colors_and_type.css)
   TKO = {
       "colors": {
           "void":        "#0A0C14",   # slide background
           "graphite":    "#13162A",   # card surfaces
           "cold_slate":  "#1E2235",   # borders
           "plasma":      "#6D28D9",   # primary accent (violet)
           "neon_indigo": "#818CF8",   # labels / links
           "acid":        "#C8E000",   # signal — KPI numbers, highlights
           "ice_white":   "#E8EEFF",   # primary text
           "ice_dim":     "#C8CCEE",   # secondary text
           "mist":        "#7880A0",   # body text
           "fog":         "#4A506E",   # captions
       },
       "fonts": {
           "display": "Barlow Condensed",   # uppercase titles
           "body":    "Inter",              # body text
       },
       "slide": {
           "width_emu":  9144000,    # 10 inches (16:9 default)
           "height_emu": 5143500,    # 5.625 inches
           "bg_color":   "#0A0C14",
       },
   }
   TKO_LOGO_PATH = "assets/tko/assets/tko-wordmark.svg"
   ```

3. **Plotly figures** use a `tko` template (registered in `config.py` via `plotly.io.templates`) so every chart picks up TKO colors and fonts automatically — set once, applies everywhere.
4. **PPTX slides** are built by `pptx_builder.py` using python-pptx primitives: solid-fill slide backgrounds, two stacked rectangles for the plasma→acid accent bar, text boxes in Barlow Condensed / Inter, PNG-converted charts inserted as pictures.
5. **SVG → PNG for logos:** SVG insertion is awkward in python-pptx. We convert the wordmark SVG to PNG once (manually, committed to `assets/tko/`) so the builder just inserts a picture.
6. **Fonts:** Barlow Condensed and Inter are Google Fonts. The README documents installing them on the build machine; if missing, PowerPoint substitutes (deck still readable, slightly off-brand).
7. **Fallback:** If `assets/tko/` is missing, `pptx_builder` uses neutral defaults (Calibri, dark grey background) and shows an `st.warning`.

### 7.5 Data Flow (Simplified)

```
CSV Upload (Streamlit file_uploader)
  → pandas reads file → schema detection (dtypes, date parsing on first 1000 rows)
  → user confirms column mapping in UI form
  → validation function returns (warnings[], errors[])
  → on confirm: pandas DataFrame → duckdb.execute("INSERT INTO ...")
  → original CSV copied to storage/raw/ for audit
  → success message; session_state["data_version"] incremented

Dashboard Page Render
  → page reads filter state from st.session_state
  → calls query.py functions (e.g. demand_mape_by_sku(filters))
  → DuckDB returns DataFrame
  → Plotly chart rendered with st.plotly_chart(fig)

Export
  → user clicks "Export PPTX" or "Export Excel" button on a page
  → wrapped in `with st.spinner("Building export…"):` for user feedback
  → synchronous call to pptx_builder.build(modules, filters) or xlsx_builder.build(...)
  → returns (bytes, filename) → st.download_button serves the file
  → typical wall time: 10–20 seconds for PPTX, 3–10 for XLSX
```

No HTTP layer, no JSON serialization, no API contracts to keep in sync. DuckDB returns DataFrames, Streamlit renders them, exports consume them directly.

### 7.6 Caching Strategy (mandatory)

Streamlit reruns the entire page script on every widget interaction. Without caching, a 5-chart page on a 1M-row fact table will not hit the 2-second filter-change NFR. Caching is therefore not optional — it's a core part of the architecture.

#### Rules

- **Cache at the query layer, never at the chart layer.** Every function in `data/query.py` is decorated with `@st.cache_data(ttl=600)` (10-minute TTL). Plotly figure creation is *not* cached — it's cheap, and caching figures interacts badly with Streamlit's serialization.
- **Cache key includes `data_version`.** Each cached query function takes `filters: Filters` and reads `current_data_version(domain)` inside the function. Streamlit hashes the `Filters` dataclass for the cache key; the `data_version` is read fresh on every call so a new upload invalidates results automatically (because the query reads from a different version).
- **`current_data_version()` is NOT cached** — it's cheap and must always reflect the latest upload.
- **Clear cache on successful upload:** the Upload page calls `st.cache_data.clear()` after a successful `load()` to drop any stale chart results that pre-date the new version.
- **Dimension lookups (`list_brands` etc.) cached with longer TTL** (1 hour) — they change rarely.

This single decorator pattern is the only "infrastructure" concession Streamlit demands. Without it, the app feels sluggish; with it, even a 5M-row DuckDB scan feels instant on second view.

---

## 8. Data Model

### 8.1 DuckDB Tables

```sql
-- Dimension: SKUs
CREATE TABLE dim_sku (
  sku_code        VARCHAR PRIMARY KEY,
  sku_name        VARCHAR,
  brand           VARCHAR,
  category        VARCHAR,
  subcategory     VARCHAR,
  uom             VARCHAR,   -- e.g. 'cases', 'kg'
  uom_to_cases    FLOAT,     -- conversion factor
  is_active       BOOLEAN DEFAULT TRUE,
  loaded_at       TIMESTAMP
);

-- Dimension: Channels
CREATE TABLE dim_channel (
  channel_code    VARCHAR PRIMARY KEY,
  channel_name    VARCHAR,
  channel_type    VARCHAR,    -- 'retail', 'ecomm', 'foodservice', 'export'
  loaded_at       TIMESTAMP
);

-- Dimension: Regions
CREATE TABLE dim_region (
  region_code     VARCHAR PRIMARY KEY,
  region_name     VARCHAR,
  country         VARCHAR,
  cluster         VARCHAR,
  loaded_at       TIMESTAMP
);

-- Dimension: Plants (manufacturing/DC sites referenced by fact_supply)
CREATE TABLE dim_plant (
  plant_code      VARCHAR PRIMARY KEY,
  plant_name      VARCHAR,
  country         VARCHAR,
  loaded_at       TIMESTAMP
);

-- Fact: Demand
CREATE TABLE fact_demand (
  period_date         DATE,          -- always first day of month
  sku_code            VARCHAR,
  channel_code        VARCHAR,
  region_code         VARCHAR,
  statistical_fcst    FLOAT,         -- volume in original UOM
  consensus_fcst      FLOAT,
  actuals             FLOAT,
  data_version        INTEGER,
  loaded_at           TIMESTAMP,
  PRIMARY KEY (period_date, sku_code, channel_code, region_code, data_version)
);

-- Fact: Supply
-- Note: DOS is NOT stored — it's computed at query time by joining to fact_demand,
-- because it depends on consensus_fcst which may update independently of supply data.
CREATE TABLE fact_supply (
  period_date           DATE,
  sku_code              VARCHAR,
  plant_code            VARCHAR,
  inventory_qty         FLOAT,
  production_plan       FLOAT,
  production_actual     FLOAT,
  capacity_plan         FLOAT,
  orders_requested      FLOAT,
  orders_delivered      FLOAT,
  data_version          INTEGER,
  loaded_at             TIMESTAMP,
  PRIMARY KEY (period_date, sku_code, plant_code, data_version)
);

-- Fact: Financial
CREATE TABLE fact_financial (
  period_date         DATE,
  sku_code            VARCHAR,
  channel_code        VARCHAR,
  revenue_actual      FLOAT,
  revenue_budget      FLOAT,
  revenue_le          FLOAT,         -- latest estimate
  gm_actual           FLOAT,
  gm_budget           FLOAT,
  promo_spend_actual  FLOAT,
  promo_spend_budget  FLOAT,
  currency_code       VARCHAR DEFAULT 'EUR',
  data_version        INTEGER,
  loaded_at           TIMESTAMP,
  PRIMARY KEY (period_date, sku_code, channel_code, data_version)
);

-- Metadata: Upload history (single-user; no `uploaded_by`)
CREATE TABLE upload_log (
  id              INTEGER PRIMARY KEY,
  filename        VARCHAR,
  domain          VARCHAR,    -- 'demand', 'supply', 'financial', 'sku_master', 'channel_master', 'region_master', 'plant_master'
  uploaded_at     TIMESTAMP,
  row_count       INTEGER,
  data_version    INTEGER,
  status          VARCHAR,    -- 'active', 'superseded', 'error'
  error_msg       VARCHAR
);
```

> **Removed from earlier drafts:** `saved_views` table (feature deferred), `uploaded_by` / `created_by` / `is_shared` columns (no multi-user model), and `fact_supply.dos` (computed at query time).

### 8.1.1 Active data version

Each fact table keeps multiple versions for rollback. All read queries must filter on the currently active version:

```sql
WHERE data_version = (SELECT MAX(data_version)
                      FROM upload_log
                      WHERE domain = '<domain>' AND status = 'active')
```

This is encapsulated in `data/query.py::current_data_version(domain: str) -> int` so individual chart queries don't repeat the subquery.

### 8.2 Calculated Metrics (computed at query time)

| Metric | Formula |
|--------|---------|
| MAPE | `AVG(ABS(actuals - consensus_fcst) / NULLIF(actuals, 0)) * 100` |
| Bias | `AVG((actuals - consensus_fcst) / NULLIF(actuals, 0)) * 100` |
| FA% | `100 - MAPE` |
| DOS | `inventory_qty / NULLIF(consensus_fcst / 30.44, 0)` |
| Fill Rate | `SUM(orders_delivered) / NULLIF(SUM(orders_requested), 0) * 100` |
| Production Adherence | `SUM(production_actual) / NULLIF(SUM(production_plan), 0) * 100` |
| Revenue vs Budget Δ% | `(revenue_actual - revenue_budget) / NULLIF(revenue_budget, 0) * 100` |
| GM% | `gm_actual / NULLIF(revenue_actual, 0) * 100` |

---

## 9. Internal API (Python function contracts)

There is no HTTP API. All "API" is regular Python function calls inside the Streamlit process. This section documents the function signatures that other modules depend on, so refactors don't silently break callers.

### 9.1 `data/ingest.py`

```python
from dataclasses import dataclass, field
from datetime import date
import pandas as pd

@dataclass
class DetectedColumn:
    name: str
    detected_type: str            # 'date' | 'int' | 'float' | 'text' | 'bool'
    sample_values: list[str]
    suggested_field: str | None   # e.g. 'period_date', 'consensus_fcst', or None

@dataclass
class SchemaReport:
    row_count: int                # total rows in the uploaded file
    columns: list[DetectedColumn]
    encoding_used: str            # 'utf-8' | 'utf-16' | 'latin-1'

@dataclass
class ValidationIssue:
    severity: str                 # 'warning' | 'error'
    row: int | None               # None if file-level (e.g. missing required column)
    column: str | None
    message: str

@dataclass
class ValidationReport:
    issues: list[ValidationIssue] = field(default_factory=list)

    @property
    def errors(self) -> list[ValidationIssue]:
        return [i for i in self.issues if i.severity == "error"]

    @property
    def warnings(self) -> list[ValidationIssue]:
        return [i for i in self.issues if i.severity == "warning"]

    @property
    def can_load(self) -> bool:
        return len(self.errors) == 0


def detect_schema(file_bytes: bytes) -> SchemaReport:
    """Read first 1000 rows; return detected columns, types, sample values.
    Auto-detects encoding (utf-8 / utf-16 / latin-1) and date formats."""

def parse(file_bytes: bytes, encoding: str) -> pd.DataFrame:
    """Parse the full CSV into a DataFrame. Called after the user confirms the
    column mapping but before validate(). Keeps original (un-renamed) column headers."""

def validate(df: pd.DataFrame, domain: str, mapping: dict[str, str]) -> ValidationReport:
    """Run validation rules per §3.1 FR-ING-04. Returns ValidationReport.
    `mapping` maps CSV column names → canonical field names from §4.1.
    Does not modify the DataFrame and does not load."""

def load(df: pd.DataFrame, domain: str, mapping: dict[str, str], mode: str) -> int:
    """Insert into DuckDB inside a single transaction.
    1. Renames df columns from CSV names → canonical names using `mapping`.
    2. Normalizes period_date to first-of-month (`.dt.to_period('M').dt.to_timestamp()`).
    3. Assigns a new data_version (max existing + 1).
    4. mode='replace' → marks prior active versions for this domain as 'superseded';
       mode='append'  → keeps prior active versions active, deduplicates on the domain PK.
    5. Writes upload_log row with status='active' and copies the raw file to storage/raw/.
    Returns the new data_version. On any failure, the transaction rolls back."""

def list_uploads() -> pd.DataFrame:
    """Return upload history from upload_log table, newest first."""
```

### 9.2 `data/query.py`

One function per chart on a dashboard. Names match the chart name. All take a `Filters` dataclass and return either a DataFrame or a typed bundle.

```python
from dataclasses import dataclass, field
from datetime import date
import pandas as pd

@dataclass
class Filters:
    period_from: date
    period_to: date
    brands: list[str] = field(default_factory=list)
    categories: list[str] = field(default_factory=list)
    channels: list[str] = field(default_factory=list)
    regions: list[str] = field(default_factory=list)
    skus: list[str] = field(default_factory=list)

@dataclass
class DemandKPIs:
    mape: float           # %
    fa_pct: float         # 100 - mape
    bias: float           # %
    volume_total: float   # in original UOM, summed

@dataclass
class SupplyKPIs:
    dos_avg: float
    fill_rate: float          # %
    production_adherence: float  # %

@dataclass
class FinancialKPIs:
    revenue_actual: float
    revenue_vs_budget_pct: float
    revenue_vs_le_pct: float
    gm_pct: float

# ---- Helpers ----
def current_data_version(domain: str) -> int | None:
    """Return the active data_version for a given domain, or None if no upload exists yet.
    Callers must handle None — typically by returning an empty DataFrame and letting the page
    render an `st.info("Upload <domain> data to see this dashboard")` placeholder.
    Not cached — must reflect uploads immediately."""

# ---- Demand ----   (all @st.cache_data; see §7.6)
def demand_kpis(f: Filters) -> DemandKPIs
def demand_volume_bridge(f: Filters) -> pd.DataFrame      # cols: stage, value
def demand_fcst_vs_actuals(f: Filters) -> pd.DataFrame    # cols: period_date, statistical_fcst, consensus_fcst, actuals
def demand_mape_by_sku(f: Filters) -> pd.DataFrame        # cols: sku_code, sku_name, brand, mape, fa_pct, bias, volume
def demand_channel_mix(f: Filters) -> pd.DataFrame        # cols: period_date, channel_code, channel_name, volume_share

# ---- Supply ----
def supply_kpis(f: Filters) -> SupplyKPIs
def supply_dos_by_sku(f: Filters) -> pd.DataFrame         # cols: sku_code, sku_name, dos, rag (target is config.RAG.dos.target_days, global)
def supply_production_adherence(f: Filters) -> pd.DataFrame  # cols: plant_code, plant_name, plan, actual, adherence_pct
def supply_inventory_heatmap(f: Filters) -> pd.DataFrame  # cols: sku_code, period_date, dos, rag
def supply_gaps(f: Filters) -> pd.DataFrame               # cols: sku_code, sku_name, period_date, shortfall

# ---- Financial ----
def financial_kpis(f: Filters) -> FinancialKPIs
def financial_revenue_waterfall(f: Filters) -> pd.DataFrame   # cols: stage (Budget/LE/Actuals), value
def financial_by_channel(f: Filters) -> pd.DataFrame          # cols: channel_code, channel_name, revenue_actual
def financial_pnl_summary(f: Filters) -> pd.DataFrame         # cols: brand, revenue, gm, promo_spend, net_revenue

# ---- Scorecard ----
def scorecard_all(f: Filters) -> pd.DataFrame
    # cols: kpi_name, value, green_threshold, amber_threshold, rag
    # (single row per KPI; thresholds come from config.RAG; "target" is implicit in the thresholds)

# ---- Dimension lookups (populate filter dropdowns; cached with ttl=3600) ----
def list_brands() -> list[str]
def list_categories() -> list[str]
def list_channels() -> list[str]
def list_regions() -> list[str]
def list_skus() -> pd.DataFrame    # cols: sku_code, sku_name (for searchable dropdown)
```

### 9.3 `exports/pptx_builder.py`

```python
def build(
    modules: list[str],              # subset of ['demand', 'supply', 'financial', 'scorecard', 'scenario']
    filters: Filters,
    template_path: str | None = None,  # defaults to config.TKO_TEMPLATE_PATH if present
) -> tuple[bytes, str]:
    """Build the PPTX synchronously. Returns (file_bytes, suggested_filename)."""
```

### 9.4 `exports/xlsx_builder.py`

```python
def build(
    modules: list[str],
    filters: Filters,
) -> tuple[bytes, str]:
    """Build the XLSX workbook synchronously. Returns (file_bytes, suggested_filename)."""
```

Both return `bytes` (not `BytesIO`) so callers can pass them straight to `st.download_button(data=..., file_name=...)`. The suggested filename includes the cycle month, e.g. `S&OP_Report_2026-05.pptx`.

### 9.5 Session state contract

Streamlit `st.session_state` keys used across pages — keep this list short and documented.

| Key | Type | Set by | Read by |
|-----|------|--------|---------|
| `filters` | `Filters` dataclass | All dashboard pages (filter bar) | All dashboard pages, export buttons |
| `data_version` | int | Upload page after successful load | All pages (shown in header) |
| `last_upload_status` | str | Upload page | Upload page (for confirmation banner) |
| `column_mapping_template` | dict | Upload page when user saves a template | Upload page on next upload |

That's the entire contract. No more keys without updating this table.

---

## 10. UI/UX Specification

Streamlit's layout primitives drive the design. We adapt to its model rather than fight it.

### 10.1 Layout (Streamlit native)

```
┌─────────────────────────────────────────────────────┐
│  st.title  "S&OP Reporting"  |  data version chip   │
├──────────┬──────────────────────────────────────────┤
│ SIDEBAR  │  FILTER BAR (st.columns inside expander) │
│ (st.     │  [Period from-to] [Brand] [Category]     │
│ sidebar) │  [Channel] [Region] [SKU search]         │
│          ├──────────────────────────────────────────┤
│ Pages    │                                          │
│ (auto    │  KPI ROW (4 × st.metric)                 │
│ from     │                                          │
│ pages/)  │  CHART GRID (st.columns(2) → plotly)     │
│          │                                          │
│ Upload   │  TABLE SECTION (st.dataframe)            │
│ Demand   │                                          │
│ Supply   │  EXPORT BUTTONS at bottom of each page:  │
│ Financial│   [⬇ Export this page as PPTX]           │
│ Scorecard│   [⬇ Export this page as XLSX]           │
│ Scenario │                                          │
│ Exec     │                                          │
│ Summary  │  (sidebar labels are auto-derived from   │
│          │   filenames: 1_Upload.py → "Upload",     │
│          │   5_Scorecard.py → "Scorecard", etc.)    │
└──────────┴──────────────────────────────────────────┘
```

### 10.2 Filter Bar
- Implemented as an `st.expander("🔍 Filters", expanded=True)` at top of each page.
- Layout: **two rows** of `st.columns(...)`:
  - Row 1: `st.columns([2, 2, 2, 2, 2, 1])` →
    `[Period from] [Period to] [Brand multiselect] [Category multiselect] [Channel multiselect] [Reset]`
  - Row 2: `st.columns(2)` →
    `[Region multiselect] [SKU search/multiselect]`
- Each input reads/writes `st.session_state["filters"]` via Streamlit's widget `key=` mechanism (no manual sync code).
- The "Reset" button in row 1 clears all filters to defaults (rolling 13 months, all dimensions selected).

### 10.3 Color Conventions (Plotly figures)

Charts use the TKO design tokens from `config.TKO["colors"]` via a registered Plotly template named `"tko"`. Semantic mapping for S&OP charts:

```python
# In config.py, derived from TKO tokens
SERIES = {
    "actuals":         TKO["colors"]["ice_white"],    # #E8EEFF — primary
    "consensus_fcst":  TKO["colors"]["plasma"],       # #6D28D9 — accent
    "statistical_fcst":TKO["colors"]["neon_indigo"],  # #818CF8 — secondary
    "budget":          TKO["colors"]["mist"],         # #7880A0 — neutral reference
    "le":              TKO["colors"]["ice_dim"],      # #C8CCEE — neutral secondary
    "signal":          TKO["colors"]["acid"],         # #C8E000 — KPI highlight
}

RAG_COLORS = {
    "green": TKO["colors"]["acid"],   # #C8E000 — green-on-dark reads as positive
    "amber": "#F59E0B",
    "red":   "#EF4444",
}
```

Plotly figures inherit `paper_bgcolor = TKO["colors"]["void"]`, `plot_bgcolor = TKO["colors"]["graphite"]`, and font `Inter / #E8EEFF` via the registered template — no per-chart styling needed.

### 10.4 Chart Sizing
- KPI metrics: `st.columns(4)` with `st.metric()` (number + delta).
- Full-width chart: pass `use_container_width=True` to `st.plotly_chart`.
- Two-column: `st.columns(2)` with one chart in each.
- Default Plotly figure height: 380 px.

### 10.5 Export Buttons (per page)
- Two `st.download_button` calls at the bottom of each dashboard page, wrapped in `with st.spinner("Building export…"):` so the user sees feedback during the build (typically 5–10 s for single-page; up to 20 s for a full-report PPTX).
- File downloads via standard browser save dialog.
- A separate "Export full report" button on the Exec Summary page bundles all modules into one PPTX.

---

## 11. Non-Functional Requirements

| Category | Requirement | Target |
|----------|-------------|--------|
| Performance | Page initial load | < 3 seconds |
| Performance | Chart re-render on filter change | < 2 seconds (Streamlit reruns the script) |
| Performance | CSV ingestion (100k rows) | < 10 seconds |
| Performance | PPTX export generation (full report) | 10–20 seconds typical, < 30 seconds worst case |
| Performance | Excel export generation (full report) | 3–10 seconds typical, < 15 seconds worst case |
| Scalability | Max rows in DuckDB | 10 million rows (single machine) |
| Scalability | Concurrent users | 1 (single-user local app) |
| Reliability | Data integrity on upload error | No partial writes — ingest in a DuckDB transaction |
| Usability | Click-count to export from any page | ≤ 2 clicks |
| Browser | Supported browsers | Chrome 120+, Edge 120+, Firefox 121+ |
| Platform | OS | Windows 10/11, macOS 13+ |
| Port | App URL | `localhost:8501` — set in `.streamlit/config.toml` (`[server] port = 8501`). Streamlit 1.36 hardcodes 8501 in its frontend bundle, so this is the only port the pinned stack works on without patching JS. |
| Maintainability | Total Python LOC for MVP | ≤ 1,500 lines |
| Maintainability | External dependencies | ≤ 10 top-level packages in requirements.txt |

---

## 12. Build Phases

Phases are scoped to be vibe-codeable in a single Claude session each. If a phase grows beyond ~500 LOC or ~5 files, split it.

### Phase 1 — Skeleton + Upload (1 session)
- [ ] `requirements.txt` with pinned versions
- [ ] `.streamlit/config.toml` — port = 3000, optional dark theme matching TKO
- [ ] `app.py` shell with title + data-version chip
- [ ] `config.py` with paths, RAG thresholds, TKO design tokens (colors + fonts), Plotly `tko` template registration
- [ ] `data/schema.sql` — full DuckDB DDL (all fact + dim tables, including `dim_plant`)
- [ ] `data/ingest.py` — `detect_schema`, `parse`, `validate`, `load`, `list_uploads` (works for **all domains**)
- [ ] `pages/1_Upload.py` — file uploader, domain selector, column mapper, validation report, upload-history expander
- [ ] Verify: upload sample CSVs for demand + sku_master + channel_master + region_master + plant_master; rows visible in DuckDB; an upload history row appears; re-uploading the same domain with mode='replace' supersedes the previous version

### Phase 2 — Demand Review Dashboard (1 session)
- [ ] `data/query.py` — `Filters`, `DemandKPIs`, `current_data_version`, all `demand_*` functions, dimension list helpers
- [ ] **All query functions decorated with `@st.cache_data(ttl=600)`; dimension helpers with `ttl=3600`** (see §7.6)
- [ ] `pages/2_Demand.py` — filter bar (per §10.2) + KPI row + 5 charts + SKU MAPE table
- [ ] Per-chart "⬇ Data (CSV)" download buttons (§3.2 FR-DASH-02)
- [ ] Cache cleared on successful upload (call added in `pages/1_Upload.py`)
- [ ] Verify: filter changes update all charts within 2 s (cold path may be slower; second filter change should be < 1 s)

### Phase 3 — Excel Export (1 session)
- [ ] `exports/xlsx_builder.py` — Cover sheet + Scorecard sheet + Demand sheet (summary + flat data in one sheet)
- [ ] Download button on Demand page (and Scorecard page once Phase 5 lands)
- [ ] Verify: exported XLSX opens in Excel, formatting intact, RAG conditional fills render, frozen header works

### Phase 4 — Supply + Financial Dashboards (1 session)
- [ ] Domain-specific validation rules in `ingest.py` for supply + financial
- [ ] All `supply_*` and `financial_*` query functions in `query.py` (with caching)
- [ ] `pages/3_Supply.py` and `pages/4_Financial.py`
- [ ] Extend `xlsx_builder.py` to add Supply + Financial sheets
- [ ] Verify: charts that depend on optional columns gracefully show "no data" when those columns are absent (§4.1)

### Phase 5 — KPI Scorecard + PPTX Export (1 session)
- [ ] `scorecard_all` query function applying `config.RAG` thresholds
- [ ] `pages/5_Scorecard.py` with RAG-colored cards
- [ ] `exports/pptx_builder.py` — programmatic slide construction using TKO tokens from `config.py` (no .pptx master file)
  - Cover slide: void bg + radial-gradient overlay + huge Barlow Condensed title + acid accent rule
  - Section dividers: eyebrow label + h1 title + accent bar
  - Chart slides: 1–2 PNGs per slide (rendered via `fig.to_image(format="png", width=1200, height=600, scale=2)`)
  - Table slides: native PPT table objects styled with TKO colors
- [ ] PPTX download button on Demand, Supply, Financial, Scorecard pages
- [ ] Verify: exported PPTX opens in PowerPoint, slides match TKO brand (dark bg, plasma+acid accents, Barlow Condensed where available), charts are readable at projector size

### Phase 6 — TKO Polish & Asset Wiring (1 session)
- [ ] Unzip `TKO Design System.zip` into `assets/tko/` (one-time, gitignored)
- [ ] Convert `assets/tko/assets/tko-wordmark.svg` to PNG once (manual, committed) for embedding on cover slide
- [ ] Add TKO wordmark to cover slide via `slide.shapes.add_picture(...)`
- [ ] Tune `tko` Plotly template colorway, font, paper/plot backgrounds to exactly match the CSS tokens
- [ ] Verify: side-by-side compare exported cover slide vs `assets/tko/slides/index.html` rendered in browser — visually consistent

### Phase 7 — Scenario Planner + Exec Summary (1 session)
- [ ] `pages/6_Scenario.py` — three sliders (Base / Upside / Downside, % adjustment to consensus) writing to `st.session_state["scenario_adjustments"]`; comparison table + chart
- [ ] `pages/7_Exec_Summary.py` — 4 KPI cards (FA, Fill Rate, Revenue vs Budget, GM%) + volume trend + revenue-vs-budget chart + scenario comparison table; "🔒 Freeze filters" toggle
- [ ] "Export full report" button on Exec Summary calling `pptx_builder.build(['demand','supply','financial','scorecard','scenario'], filters)`
- [ ] Verify: full-report PPTX is < 30 s build, opens cleanly, scenario slide reflects current session adjustments

### Out of MVP scope (deferred until there is a real need)
- User accounts, RBAC, audit log, `uploaded_by` tracking
- Saved views / shareable URLs (and the `saved_views` table)
- Cross-filtering (Streamlit reruns make this awkward — defer until justified)
- Custom KPI threshold UI (edit `config.py` for now)
- Column mapping template library beyond a single saved mapping per domain
- Per-SKU `dos_target_days` (single global target in `config.py` for now)
- Weekly granularity (monthly only in MVP)
- Multi-currency FX conversion (single currency in `config.py`)
- Native PowerPoint chart objects (PNG-via-kaleido only)
- Price/volume/mix decomposition in revenue waterfall
- In-app TKO template upload UI (manual file copy for now)
- SSO / SAML

---

## 13. Glossary

| Term | Definition |
|------|-----------|
| S&OP | Sales & Operations Planning — monthly cross-functional process to balance demand and supply |
| CPG | Consumer Packaged Goods — companies selling branded consumer products through retail channels |
| MAPE | Mean Absolute Percentage Error — primary forecast accuracy metric |
| FA% | Forecast Accuracy % = 100 - MAPE |
| Bias | Systematic over- or under-forecasting |
| DOS | Days of Supply — how many days of demand can be covered by current inventory |
| LE | Latest Estimate — most recent financial projection within the fiscal year |
| SKU | Stock Keeping Unit — a unique product variant |
| RAG | Red / Amber / Green — traffic-light KPI status system |
| Consensus Forecast | The agreed demand plan after human adjustments on the statistical forecast |
| Statistical Forecast | System-generated forecast from historical data and algorithms |
| UOM | Unit of Measure (cases, kg, units, etc.) |
| DuckDB | In-process analytical SQL database; handles large CSV/Parquet queries without a server |
| TKO | The branding/template framework used for PPTX exports in this project |
| Cycle Month | The S&OP planning month (e.g., "May 2026 S&OP cycle") |

---

*Last updated: 2026-05-20 | Author: S&OP Platform Project*
