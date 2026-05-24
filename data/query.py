"""DuckDB → DataFrame/dataclass read layer for dashboard pages.

Implements the contracts in CONTRACTS.md §5.2. Every public query function:

* takes a :class:`Filters` dataclass (rule §3.5),
* filters by the active ``data_version`` for its domain (rule §3.1),
* UOM-normalizes cross-SKU volume sums via ``dim_sku.uom_to_cases`` (rule §3.9),
* and is wrapped with ``@st.cache_data`` (rule §3.2).

Pages must go through this module to read DuckDB — no raw connections in
``pages/`` (rule §3.6).
"""

from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import date

import duckdb
import pandas as pd
import streamlit as st

import config

# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------


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
    volume_total: float   # cases (UOM-normalized)


@dataclass
class SupplyKPIs:
    dos_avg: float                # days — point-in-time at latest period
    fill_rate: float              # % — aggregated over window
    production_adherence: float   # % — aggregated over window


@dataclass
class FinancialKPIs:
    revenue_actual: float
    revenue_vs_budget_pct: float
    revenue_vs_le_pct: float
    gm_pct: float


# ---------------------------------------------------------------------------
# Connection management
# ---------------------------------------------------------------------------

# Schema is run once per process. Streamlit reruns the script per interaction,
# but modules are imported once, so this flag persists across reruns.
_schema_initialized = False


def _ensure_schema() -> None:
    global _schema_initialized
    if _schema_initialized:
        return
    c = duckdb.connect(config.DUCKDB_PATH)
    try:
        with open(config.SCHEMA_SQL_PATH, "r", encoding="utf-8") as f:
            c.execute(f.read())
    finally:
        c.close()
    _schema_initialized = True


@contextmanager
def _conn():
    """Open a short-lived read-only connection. Schema is ensured first."""
    _ensure_schema()
    c = duckdb.connect(config.DUCKDB_PATH, read_only=True)
    try:
        yield c
    finally:
        c.close()


# ---------------------------------------------------------------------------
# Version + defaults — NOT cached (rule §3.2)
# ---------------------------------------------------------------------------


def current_data_version(domain: str) -> int | None:
    """Active data_version for ``domain``, or None if no active upload exists."""
    with _conn() as c:
        row = c.execute(
            "SELECT data_version FROM upload_log "
            "WHERE domain = ? AND status = 'active' "
            "ORDER BY data_version DESC LIMIT 1",
            [domain],
        ).fetchone()
    return int(row[0]) if row else None


def default_filters() -> Filters:
    """Default per CONTRACTS §7.1: period_to = MAX(period_date) in active
    fact_demand (first-of-month), or today's first-of-month if no demand yet;
    period_from = period_to − DEFAULT_LOOKBACK_M months. Dim lists empty.
    """
    today_fom = pd.Timestamp(date.today()).to_period("M").to_timestamp().date()
    period_to = today_fom
    dv = current_data_version("demand")
    if dv is not None:
        with _conn() as c:
            row = c.execute(
                "SELECT MAX(period_date) FROM fact_demand WHERE data_version = ?",
                [dv],
            ).fetchone()
        if row and row[0] is not None:
            period_to = (
                pd.Timestamp(row[0]).to_period("M").to_timestamp().date()
            )
    period_from = (
        pd.Timestamp(period_to)
        - pd.DateOffset(months=config.DEFAULT_LOOKBACK_M)
    ).date().replace(day=1)
    return Filters(period_from=period_from, period_to=period_to)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _demand_where(f: Filters, dv: int) -> tuple[str, list]:
    """Compose WHERE clause + params for a fact_demand (alias ``f``) join with
    dim_sku (alias ``s``). Empty dim lists mean "no filter" (rule §3.7).
    """
    clauses = [
        "f.data_version = ?",
        "f.period_date BETWEEN ? AND ?",
    ]
    params: list = [dv, f.period_from, f.period_to]
    if f.brands:
        clauses.append(
            "s.brand IN (" + ",".join(["?"] * len(f.brands)) + ")"
        )
        params.extend(f.brands)
    if f.categories:
        clauses.append(
            "s.category IN (" + ",".join(["?"] * len(f.categories)) + ")"
        )
        params.extend(f.categories)
    if f.channels:
        clauses.append(
            "f.channel_code IN (" + ",".join(["?"] * len(f.channels)) + ")"
        )
        params.extend(f.channels)
    if f.regions:
        clauses.append(
            "f.region_code IN (" + ",".join(["?"] * len(f.regions)) + ")"
        )
        params.extend(f.regions)
    if f.skus:
        clauses.append(
            "f.sku_code IN (" + ",".join(["?"] * len(f.skus)) + ")"
        )
        params.extend(f.skus)
    return " AND ".join(clauses), params


def _rag_for_fa(fa_pct: float | None) -> str:
    if fa_pct is None or pd.isna(fa_pct):
        return "red"
    if fa_pct >= config.RAG["fa"]["green"]:
        return "green"
    if fa_pct >= config.RAG["fa"]["amber"]:
        return "amber"
    return "red"


def _rag_for_fill_rate(fr: float | None) -> str:
    if fr is None or pd.isna(fr):
        return "red"
    if fr >= config.RAG["fill_rate"]["green"]:
        return "green"
    if fr >= config.RAG["fill_rate"]["amber"]:
        return "amber"
    return "red"


def _rag_for_dos(dos: float | None) -> str:
    if dos is None or pd.isna(dos):
        return "red"
    cfg = config.RAG["dos"]
    delta = abs(float(dos) - cfg["target_days"])
    if delta <= cfg["green_band"]:
        return "green"
    if delta <= cfg["amber_band"]:
        return "amber"
    return "red"


def _rag_for_bias(bias: float | None) -> str:
    if bias is None or pd.isna(bias):
        return "red"
    cfg = config.RAG["bias"]
    mag = abs(float(bias))
    if mag <= cfg["green_abs"]:
        return "green"
    if mag <= cfg["amber_abs"]:
        return "amber"
    return "red"


def _rag_for_rev_vs_budget(ratio_pct: float | None) -> str:
    """``ratio_pct`` = actual / budget × 100 (achievement %)."""
    if ratio_pct is None or pd.isna(ratio_pct):
        return "red"
    cfg = config.RAG["rev_vs_budget"]
    if ratio_pct >= cfg["green"]:
        return "green"
    if ratio_pct >= cfg["amber"]:
        return "amber"
    return "red"


def _rag_for_gm_vs_budget(pp_delta: float | None) -> str:
    """``pp_delta`` = GM% actual − GM% budget, in percentage points."""
    if pp_delta is None or pd.isna(pp_delta):
        return "red"
    cfg = config.RAG["gm_vs_budget_pp"]
    if pp_delta >= cfg["green"]:
        return "green"
    if pp_delta >= cfg["amber"]:
        return "amber"
    return "red"


def _supply_where(f: Filters, dv: int) -> tuple[str, list]:
    """WHERE clause for fact_supply (alias ``f``) joined to dim_sku (alias ``s``).

    Channel/region filters are intentionally ignored — fact_supply is plant-keyed
    and has no channel_code/region_code columns (a plant ships into many of both).
    Brand / category / sku filters propagate via dim_sku.
    """
    clauses = [
        "f.data_version = ?",
        "f.period_date BETWEEN ? AND ?",
    ]
    params: list = [dv, f.period_from, f.period_to]
    if f.brands:
        clauses.append(
            "s.brand IN (" + ",".join(["?"] * len(f.brands)) + ")"
        )
        params.extend(f.brands)
    if f.categories:
        clauses.append(
            "s.category IN (" + ",".join(["?"] * len(f.categories)) + ")"
        )
        params.extend(f.categories)
    if f.skus:
        clauses.append(
            "f.sku_code IN (" + ",".join(["?"] * len(f.skus)) + ")"
        )
        params.extend(f.skus)
    return " AND ".join(clauses), params


def _financial_where(f: Filters, dv: int) -> tuple[str, list]:
    """WHERE clause for fact_financial (alias ``f``) joined to dim_sku (``s``).

    Region filter is intentionally ignored — fact_financial has no region_code.
    """
    clauses = [
        "f.data_version = ?",
        "f.period_date BETWEEN ? AND ?",
    ]
    params: list = [dv, f.period_from, f.period_to]
    if f.brands:
        clauses.append(
            "s.brand IN (" + ",".join(["?"] * len(f.brands)) + ")"
        )
        params.extend(f.brands)
    if f.categories:
        clauses.append(
            "s.category IN (" + ",".join(["?"] * len(f.categories)) + ")"
        )
        params.extend(f.categories)
    if f.channels:
        clauses.append(
            "f.channel_code IN (" + ",".join(["?"] * len(f.channels)) + ")"
        )
        params.extend(f.channels)
    if f.skus:
        clauses.append(
            "f.sku_code IN (" + ",".join(["?"] * len(f.skus)) + ")"
        )
        params.extend(f.skus)
    return " AND ".join(clauses), params


def _empty(cols: list[str]) -> pd.DataFrame:
    return pd.DataFrame({c: pd.Series(dtype="object") for c in cols})


# ---------------------------------------------------------------------------
# Demand queries (ttl=600 — rule §3.2)
# ---------------------------------------------------------------------------


@st.cache_data(ttl=600, show_spinner=False)
def demand_kpis(f: Filters) -> DemandKPIs:
    dv = current_data_version("demand")
    _nan = float("nan")
    if dv is None:
        return DemandKPIs(mape=_nan, fa_pct=_nan, bias=_nan, volume_total=0.0)
    where, params = _demand_where(f, dv)
    sql = f"""
        SELECT
            AVG(ABS(f.actuals - f.consensus_fcst) / NULLIF(f.actuals, 0)) * 100 AS mape,
            AVG((f.actuals - f.consensus_fcst) / NULLIF(f.actuals, 0)) * 100 AS bias,
            SUM(f.actuals * COALESCE(s.uom_to_cases, 1)) AS volume_total
        FROM fact_demand f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where}
    """
    with _conn() as c:
        row = c.execute(sql, params).fetchone()
    # SUM over an empty result set is NULL — distinguishes "no rows in window"
    # from "rows exist, volume sums to 0". Surface NaN so callers can render
    # "—" rather than a misleading FA=100% on empty data.
    if not row or row[2] is None:
        return DemandKPIs(mape=_nan, fa_pct=_nan, bias=_nan, volume_total=0.0)
    mape = float(row[0]) if row[0] is not None else _nan
    bias = float(row[1]) if row[1] is not None else _nan
    fa = _nan if pd.isna(mape) else 100.0 - mape
    return DemandKPIs(mape=mape, fa_pct=fa, bias=bias, volume_total=float(row[2]))


@st.cache_data(ttl=600, show_spinner=False)
def demand_volume_bridge(f: Filters) -> pd.DataFrame:
    """Stages: Prior Period Actuals → Statistical Fcst → Consensus Fcst → Current Actuals.

    "Prior" = the same-length window immediately preceding ``f``'s window
    (e.g. an R13M window's prior is the 13 months before that).
    """
    dv = current_data_version("demand")
    if dv is None:
        return _empty(["stage", "value"])

    span_months = (f.period_to.year - f.period_from.year) * 12 + (
        f.period_to.month - f.period_from.month
    )
    prior_to = (
        pd.Timestamp(f.period_from) - pd.DateOffset(months=1)
    ).date().replace(day=1)
    prior_from = (
        pd.Timestamp(prior_to) - pd.DateOffset(months=span_months)
    ).date().replace(day=1)

    where_curr, params_curr = _demand_where(f, dv)
    prior = Filters(
        period_from=prior_from,
        period_to=prior_to,
        brands=f.brands,
        categories=f.categories,
        channels=f.channels,
        regions=f.regions,
        skus=f.skus,
    )
    where_prior, params_prior = _demand_where(prior, dv)

    sql_curr = f"""
        SELECT
            SUM(f.statistical_fcst * COALESCE(s.uom_to_cases, 1)) AS stat,
            SUM(f.consensus_fcst   * COALESCE(s.uom_to_cases, 1)) AS cons,
            SUM(f.actuals          * COALESCE(s.uom_to_cases, 1)) AS act
        FROM fact_demand f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where_curr}
    """
    sql_prior = f"""
        SELECT SUM(f.actuals * COALESCE(s.uom_to_cases, 1)) AS prior_act
        FROM fact_demand f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where_prior}
    """
    with _conn() as c:
        stat, cons, act = c.execute(sql_curr, params_curr).fetchone()
        (prior_act,) = c.execute(sql_prior, params_prior).fetchone()

    def _n(v):
        return float(v) if v is not None else 0.0

    return pd.DataFrame({
        "stage": ["Prior Period Actuals", "Statistical Fcst",
                  "Consensus Fcst", "Current Actuals"],
        "value": [_n(prior_act), _n(stat), _n(cons), _n(act)],
    })


@st.cache_data(ttl=600, show_spinner=False)
def demand_fcst_vs_actuals(f: Filters) -> pd.DataFrame:
    dv = current_data_version("demand")
    if dv is None:
        return _empty(["period_date", "statistical_fcst",
                       "consensus_fcst", "actuals"])
    where, params = _demand_where(f, dv)
    sql = f"""
        SELECT
            f.period_date,
            SUM(f.statistical_fcst * COALESCE(s.uom_to_cases, 1)) AS statistical_fcst,
            SUM(f.consensus_fcst   * COALESCE(s.uom_to_cases, 1)) AS consensus_fcst,
            SUM(f.actuals          * COALESCE(s.uom_to_cases, 1)) AS actuals
        FROM fact_demand f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where}
        GROUP BY f.period_date
        ORDER BY f.period_date
    """
    with _conn() as c:
        return c.execute(sql, params).fetchdf()


@st.cache_data(ttl=600, show_spinner=False)
def demand_mape_by_sku(f: Filters) -> pd.DataFrame:
    dv = current_data_version("demand")
    if dv is None:
        return _empty(["sku_code", "sku_name", "brand",
                       "mape", "fa_pct", "bias", "volume", "rag"])
    where, params = _demand_where(f, dv)
    sql = f"""
        SELECT
            f.sku_code,
            COALESCE(s.sku_name, f.sku_code) AS sku_name,
            s.brand AS brand,
            AVG(ABS(f.actuals - f.consensus_fcst) / NULLIF(f.actuals, 0)) * 100 AS mape,
            AVG((f.actuals - f.consensus_fcst) / NULLIF(f.actuals, 0)) * 100 AS bias,
            SUM(f.actuals * COALESCE(s.uom_to_cases, 1)) AS volume
        FROM fact_demand f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where}
        GROUP BY f.sku_code, s.sku_name, s.brand
        ORDER BY volume DESC NULLS LAST
    """
    with _conn() as c:
        df = c.execute(sql, params).fetchdf()
    if df.empty:
        df["fa_pct"] = pd.Series(dtype="float")
        df["rag"] = pd.Series(dtype="object")
        return df[["sku_code", "sku_name", "brand",
                   "mape", "fa_pct", "bias", "volume", "rag"]]
    df["fa_pct"] = 100.0 - df["mape"]
    df["rag"] = df["fa_pct"].apply(_rag_for_fa)
    return df[["sku_code", "sku_name", "brand",
               "mape", "fa_pct", "bias", "volume", "rag"]]


@st.cache_data(ttl=600, show_spinner=False)
def demand_channel_mix(f: Filters) -> pd.DataFrame:
    dv = current_data_version("demand")
    if dv is None:
        return _empty(["period_date", "channel_code",
                       "channel_name", "volume_share"])
    where, params = _demand_where(f, dv)
    sql = f"""
        WITH per_channel AS (
            SELECT
                f.period_date,
                f.channel_code,
                COALESCE(ch.channel_name, f.channel_code) AS channel_name,
                SUM(f.actuals * COALESCE(s.uom_to_cases, 1)) AS volume
            FROM fact_demand f
            LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
            LEFT JOIN dim_channel ch ON f.channel_code = ch.channel_code
            WHERE {where}
            GROUP BY f.period_date, f.channel_code, ch.channel_name
        ),
        per_period AS (
            SELECT period_date, SUM(volume) AS total
            FROM per_channel
            GROUP BY period_date
        )
        SELECT
            pc.period_date,
            pc.channel_code,
            pc.channel_name,
            CASE WHEN pp.total > 0 THEN pc.volume / pp.total * 100 ELSE 0 END
                AS volume_share
        FROM per_channel pc
        JOIN per_period pp USING (period_date)
        ORDER BY pc.period_date, pc.channel_code
    """
    with _conn() as c:
        return c.execute(sql, params).fetchdf()


# ---------------------------------------------------------------------------
# Supply queries (ttl=600 — rule §3.2)
# ---------------------------------------------------------------------------
#
# DOS uses fact_supply LEFT JOIN fact_demand on (period_date, sku_code), with
# fact_demand pre-aggregated to that grain (fact_demand's PK includes channel
# and region, so without aggregation we'd over-count demand). Same-period
# divisor, NULLs excluded — see CONTRACTS §7.2.


@st.cache_data(ttl=600, show_spinner=False)
def supply_kpis(f: Filters) -> SupplyKPIs:
    sv = current_data_version("supply")
    dv = current_data_version("demand")
    nan = float("nan")
    if sv is None:
        return SupplyKPIs(dos_avg=nan, fill_rate=nan, production_adherence=nan)

    where_s, params_s = _supply_where(f, sv)

    # Fill rate + production adherence: aggregated over the full window,
    # both UOM-normalized to cases (rule §3.9).
    sql_aggs = f"""
        SELECT
            SUM(f.orders_delivered  * COALESCE(s.uom_to_cases, 1)) AS od,
            SUM(f.orders_requested  * COALESCE(s.uom_to_cases, 1)) AS orq,
            SUM(f.production_actual * COALESCE(s.uom_to_cases, 1)) AS pa,
            SUM(f.production_plan   * COALESCE(s.uom_to_cases, 1)) AS pp
        FROM fact_supply f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where_s}
    """
    with _conn() as c:
        row = c.execute(sql_aggs, params_s).fetchone()

    def _safe_pct(num, den):
        if num is None or den is None or den == 0:
            return nan
        return float(num) / float(den) * 100.0

    fill_rate = _safe_pct(row[0] if row else None, row[1] if row else None)
    prod_adh  = _safe_pct(row[2] if row else None, row[3] if row else None)

    # DOS — point-in-time at latest period in the window; demand-weighted aggregate
    # (CONTRACTS §7.3). Skip if demand has no active upload.
    dos_avg = nan
    if dv is not None:
        sql_dos = f"""
            WITH supply_window AS (
                SELECT f.period_date, f.sku_code, f.inventory_qty,
                       COALESCE(s.uom_to_cases, 1) AS u2c
                FROM fact_supply f
                LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
                WHERE {where_s}
            ),
            latest AS (SELECT MAX(period_date) AS d FROM supply_window),
            demand_agg AS (
                SELECT period_date, sku_code, SUM(consensus_fcst) AS cf
                FROM fact_demand
                WHERE data_version = ?
                  AND period_date = (SELECT d FROM latest)
                GROUP BY period_date, sku_code
            )
            -- CONTRACTS §7.2: rows with no matching demand are excluded from
            -- DOS aggregations. FILTER (...) keeps numerator and denominator
            -- symmetric — without it, a SKU with supply but no demand would
            -- inflate inv_cases while contributing 0 to dem_cases.
            SELECT
                SUM(sw.inventory_qty * sw.u2c) FILTER (WHERE da.cf IS NOT NULL)
                    AS inv_cases,
                SUM(da.cf            * sw.u2c) FILTER (WHERE da.cf IS NOT NULL)
                    AS dem_cases
            FROM supply_window sw
            JOIN latest l ON sw.period_date = l.d
            LEFT JOIN demand_agg da
                ON da.period_date = sw.period_date
               AND da.sku_code    = sw.sku_code
        """
        with _conn() as c:
            inv_cases, dem_cases = c.execute(sql_dos, params_s + [dv]).fetchone()
        if inv_cases is not None and dem_cases not in (None, 0):
            dos_avg = float(inv_cases) / (float(dem_cases) / 30.44)

    return SupplyKPIs(
        dos_avg=dos_avg, fill_rate=fill_rate, production_adherence=prod_adh
    )


@st.cache_data(ttl=600, show_spinner=False)
def supply_dos_by_sku(f: Filters) -> pd.DataFrame:
    """Per-SKU DOS at the latest period in f's window. DOS in days — UOM-independent
    (rule §3.9: intrinsic values stay native, single SKU per row)."""
    sv = current_data_version("supply")
    dv = current_data_version("demand")
    if sv is None:
        return _empty(["sku_code", "sku_name", "dos", "rag"])
    where_s, params_s = _supply_where(f, sv)

    if dv is None:
        # Without demand, DOS is undefined — surface the SKU list with NaN dos.
        sql = f"""
            WITH supply_window AS (
                SELECT f.period_date, f.sku_code, f.inventory_qty,
                       COALESCE(s.sku_name, f.sku_code) AS sku_name
                FROM fact_supply f
                LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
                WHERE {where_s}
            ),
            latest AS (SELECT MAX(period_date) AS d FROM supply_window)
            SELECT sw.sku_code, sw.sku_name, NULL::DOUBLE AS dos
            FROM supply_window sw JOIN latest l ON sw.period_date = l.d
            GROUP BY sw.sku_code, sw.sku_name
            ORDER BY sw.sku_code
        """
        with _conn() as c:
            df = c.execute(sql, params_s).fetchdf()
        df["rag"] = df["dos"].apply(_rag_for_dos)
        return df[["sku_code", "sku_name", "dos", "rag"]]

    sql = f"""
        WITH supply_window AS (
            SELECT f.period_date, f.sku_code, f.inventory_qty,
                   COALESCE(s.sku_name, f.sku_code) AS sku_name
            FROM fact_supply f
            LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
            WHERE {where_s}
        ),
        latest AS (SELECT MAX(period_date) AS d FROM supply_window),
        demand_agg AS (
            SELECT period_date, sku_code, SUM(consensus_fcst) AS cf
            FROM fact_demand
            WHERE data_version = ?
              AND period_date = (SELECT d FROM latest)
            GROUP BY period_date, sku_code
        )
        SELECT
            sw.sku_code,
            sw.sku_name,
            SUM(sw.inventory_qty) AS inv,
            MAX(da.cf) AS cf
        FROM supply_window sw
        JOIN latest l ON sw.period_date = l.d
        LEFT JOIN demand_agg da
            ON da.period_date = sw.period_date AND da.sku_code = sw.sku_code
        GROUP BY sw.sku_code, sw.sku_name
        ORDER BY sw.sku_code
    """
    with _conn() as c:
        df = c.execute(sql, params_s + [dv]).fetchdf()
    if df.empty:
        return _empty(["sku_code", "sku_name", "dos", "rag"])
    daily = df["cf"] / 30.44
    df["dos"] = df["inv"].astype(float).where(daily > 0) / daily.where(daily > 0)
    df["rag"] = df["dos"].apply(_rag_for_dos)
    return df[["sku_code", "sku_name", "dos", "rag"]]


@st.cache_data(ttl=600, show_spinner=False)
def supply_production_adherence(f: Filters) -> pd.DataFrame:
    """Per-plant plan vs actual across the filter window. Plan/actual in cases."""
    sv = current_data_version("supply")
    if sv is None:
        return _empty(["plant_code", "plant_name", "plan", "actual", "adherence_pct"])
    where_s, params_s = _supply_where(f, sv)
    sql = f"""
        SELECT
            f.plant_code,
            COALESCE(p.plant_name, f.plant_code) AS plant_name,
            SUM(f.production_plan   * COALESCE(s.uom_to_cases, 1)) AS plan,
            SUM(f.production_actual * COALESCE(s.uom_to_cases, 1)) AS actual,
            SUM(f.production_actual * COALESCE(s.uom_to_cases, 1))
              / NULLIF(SUM(f.production_plan * COALESCE(s.uom_to_cases, 1)), 0) * 100
                AS adherence_pct
        FROM fact_supply f
        LEFT JOIN dim_sku   s ON f.sku_code   = s.sku_code
        LEFT JOIN dim_plant p ON f.plant_code = p.plant_code
        WHERE {where_s}
        GROUP BY f.plant_code, p.plant_name
        ORDER BY f.plant_code
    """
    with _conn() as c:
        return c.execute(sql, params_s).fetchdf()


@st.cache_data(ttl=600, show_spinner=False)
def supply_capacity_utilization(f: Filters) -> pd.DataFrame:
    """Per-plant capacity utilization across the filter window.
    Plants with no capacity_plan rows are omitted (HAVING)."""
    sv = current_data_version("supply")
    if sv is None:
        return _empty(["plant_code", "plant_name", "actual", "capacity", "utilization_pct"])
    where_s, params_s = _supply_where(f, sv)
    sql = f"""
        SELECT
            f.plant_code,
            COALESCE(p.plant_name, f.plant_code) AS plant_name,
            SUM(f.production_actual * COALESCE(s.uom_to_cases, 1)) AS actual,
            SUM(f.capacity_plan     * COALESCE(s.uom_to_cases, 1)) AS capacity,
            SUM(f.production_actual * COALESCE(s.uom_to_cases, 1))
              / NULLIF(SUM(f.capacity_plan * COALESCE(s.uom_to_cases, 1)), 0) * 100
                AS utilization_pct
        FROM fact_supply f
        LEFT JOIN dim_sku   s ON f.sku_code   = s.sku_code
        LEFT JOIN dim_plant p ON f.plant_code = p.plant_code
        WHERE {where_s}
        GROUP BY f.plant_code, p.plant_name
        HAVING SUM(f.capacity_plan) IS NOT NULL AND SUM(f.capacity_plan) > 0
        ORDER BY f.plant_code
    """
    with _conn() as c:
        return c.execute(sql, params_s).fetchdf()


@st.cache_data(ttl=600, show_spinner=False)
def supply_inventory_heatmap(f: Filters) -> pd.DataFrame:
    """Per (sku, period) DOS for the heatmap. Single SKU per cell — no UOM
    aggregation needed (rule §3.9, per-SKU intrinsic value)."""
    sv = current_data_version("supply")
    dv = current_data_version("demand")
    if sv is None or dv is None:
        return _empty(["sku_code", "period_date", "dos", "rag"])
    where_s, params_s = _supply_where(f, sv)
    sql = f"""
        WITH supply_window AS (
            SELECT f.period_date, f.sku_code, SUM(f.inventory_qty) AS inv
            FROM fact_supply f
            LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
            WHERE {where_s}
            GROUP BY f.period_date, f.sku_code
        ),
        demand_agg AS (
            SELECT period_date, sku_code, SUM(consensus_fcst) AS cf
            FROM fact_demand WHERE data_version = ?
            GROUP BY period_date, sku_code
        )
        SELECT
            sw.sku_code, sw.period_date,
            sw.inv / NULLIF(da.cf / 30.44, 0) AS dos
        FROM supply_window sw
        LEFT JOIN demand_agg da
            ON da.period_date = sw.period_date AND da.sku_code = sw.sku_code
        ORDER BY sw.sku_code, sw.period_date
    """
    with _conn() as c:
        df = c.execute(sql, params_s + [dv]).fetchdf()
    if df.empty:
        return _empty(["sku_code", "period_date", "dos", "rag"])
    df["rag"] = df["dos"].apply(_rag_for_dos)
    return df[["sku_code", "period_date", "dos", "rag"]]


@st.cache_data(ttl=600, show_spinner=False)
def supply_gaps(f: Filters) -> pd.DataFrame:
    """Projected stockouts in the 3 months immediately after f.period_to.

    A "gap" row is (sku, period) where consensus_fcst > inventory_qty +
    production_plan. shortfall in the SKU's native UOM (rule §3.9 — single
    planning quantity, kept native for ordering decisions).
    """
    sv = current_data_version("supply")
    dv = current_data_version("demand")
    if sv is None or dv is None:
        return _empty(["sku_code", "sku_name", "period_date", "shortfall"])

    forward_from = (
        pd.Timestamp(f.period_to) + pd.DateOffset(months=1)
    ).date().replace(day=1)
    forward_to = (
        pd.Timestamp(f.period_to) + pd.DateOffset(months=3)
    ).date().replace(day=1)

    # Build supply-side WHERE for the forward window using the same dim filters.
    forward_filter = Filters(
        period_from=forward_from,
        period_to=forward_to,
        brands=f.brands, categories=f.categories,
        channels=f.channels, regions=f.regions, skus=f.skus,
    )
    where_s, params_s = _supply_where(forward_filter, sv)

    sql = f"""
        WITH supply_fwd AS (
            SELECT f.period_date, f.sku_code,
                   COALESCE(s.sku_name, f.sku_code) AS sku_name,
                   SUM(f.inventory_qty)   AS inv,
                   SUM(f.production_plan) AS prod_plan
            FROM fact_supply f
            LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
            WHERE {where_s}
            GROUP BY f.period_date, f.sku_code, s.sku_name
        ),
        demand_agg AS (
            SELECT period_date, sku_code, SUM(consensus_fcst) AS cf
            FROM fact_demand WHERE data_version = ?
            GROUP BY period_date, sku_code
        )
        SELECT
            sf.sku_code, sf.sku_name, sf.period_date,
            da.cf - (COALESCE(sf.inv, 0) + COALESCE(sf.prod_plan, 0)) AS shortfall
        FROM supply_fwd sf
        LEFT JOIN demand_agg da
            ON da.period_date = sf.period_date AND da.sku_code = sf.sku_code
        WHERE da.cf IS NOT NULL
          AND da.cf > COALESCE(sf.inv, 0) + COALESCE(sf.prod_plan, 0)
        ORDER BY sf.period_date, shortfall DESC
    """
    with _conn() as c:
        return c.execute(sql, params_s + [dv]).fetchdf()


@st.cache_data(ttl=600, show_spinner=False)
def supply_fill_rate_trend(f: Filters) -> pd.DataFrame:
    """Monthly fill rate across the filter window, for the KPI sparkline (SPEC §5.2)."""
    sv = current_data_version("supply")
    if sv is None:
        return _empty(["period_date", "fill_rate"])
    where_s, params_s = _supply_where(f, sv)
    sql = f"""
        SELECT
            f.period_date,
            SUM(f.orders_delivered * COALESCE(s.uom_to_cases, 1))
              / NULLIF(SUM(f.orders_requested * COALESCE(s.uom_to_cases, 1)), 0) * 100
                AS fill_rate
        FROM fact_supply f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where_s}
        GROUP BY f.period_date
        ORDER BY f.period_date
    """
    with _conn() as c:
        return c.execute(sql, params_s).fetchdf()


# ---------------------------------------------------------------------------
# Financial queries (ttl=600 — rule §3.2)
# ---------------------------------------------------------------------------
# Single-currency MVP — values are assumed to be in config.CURRENCY (rule §3.10).
# No UOM normalization on monetary columns; volumes do not enter these queries
# except via the price-proxy work in Phase 7.


@st.cache_data(ttl=600, show_spinner=False)
def financial_kpis(f: Filters) -> FinancialKPIs:
    fv = current_data_version("financial")
    nan = float("nan")
    if fv is None:
        return FinancialKPIs(
            revenue_actual=0.0,
            revenue_vs_budget_pct=nan,
            revenue_vs_le_pct=nan,
            gm_pct=nan,
        )
    where_f, params_f = _financial_where(f, fv)
    sql = f"""
        SELECT
            SUM(f.revenue_actual) AS rev_act,
            SUM(f.revenue_budget) AS rev_bud,
            SUM(f.revenue_le)     AS rev_le,
            SUM(f.gm_actual)      AS gm_act
        FROM fact_financial f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where_f}
    """
    with _conn() as c:
        row = c.execute(sql, params_f).fetchone()
    if not row or row[0] is None:
        return FinancialKPIs(
            revenue_actual=0.0,
            revenue_vs_budget_pct=nan,
            revenue_vs_le_pct=nan,
            gm_pct=nan,
        )
    rev_act = float(row[0])
    rev_bud = float(row[1]) if row[1] is not None else None
    rev_le  = float(row[2]) if row[2] is not None else None
    gm_act  = float(row[3]) if row[3] is not None else None

    def _delta_pct(act, base):
        if base is None or base == 0:
            return nan
        return (act - base) / base * 100.0

    return FinancialKPIs(
        revenue_actual=rev_act,
        revenue_vs_budget_pct=_delta_pct(rev_act, rev_bud),
        revenue_vs_le_pct=_delta_pct(rev_act, rev_le),
        gm_pct=(gm_act / rev_act * 100.0) if (gm_act is not None and rev_act) else nan,
    )


@st.cache_data(ttl=600, show_spinner=False)
def financial_revenue_waterfall(f: Filters) -> pd.DataFrame:
    """Budget → LE → Actuals stages (SPEC §5.3 — MVP version)."""
    fv = current_data_version("financial")
    if fv is None:
        return _empty(["stage", "value"])
    where_f, params_f = _financial_where(f, fv)
    sql = f"""
        SELECT
            SUM(f.revenue_budget) AS rev_bud,
            SUM(f.revenue_le)     AS rev_le,
            SUM(f.revenue_actual) AS rev_act
        FROM fact_financial f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where_f}
    """
    with _conn() as c:
        row = c.execute(sql, params_f).fetchone()

    def _n(v):
        return float(v) if v is not None else 0.0

    return pd.DataFrame({
        "stage": ["Budget", "LE", "Actuals"],
        "value": [_n(row[0]), _n(row[1]), _n(row[2])],
    })


@st.cache_data(ttl=600, show_spinner=False)
def financial_by_channel(f: Filters) -> pd.DataFrame:
    fv = current_data_version("financial")
    if fv is None:
        return _empty(["channel_code", "channel_name", "revenue_actual"])
    where_f, params_f = _financial_where(f, fv)
    sql = f"""
        SELECT
            f.channel_code,
            COALESCE(ch.channel_name, f.channel_code) AS channel_name,
            SUM(f.revenue_actual) AS revenue_actual
        FROM fact_financial f
        LEFT JOIN dim_sku     s  ON f.sku_code     = s.sku_code
        LEFT JOIN dim_channel ch ON f.channel_code = ch.channel_code
        WHERE {where_f}
        GROUP BY f.channel_code, ch.channel_name
        ORDER BY revenue_actual DESC NULLS LAST
    """
    with _conn() as c:
        return c.execute(sql, params_f).fetchdf()


@st.cache_data(ttl=600, show_spinner=False)
def financial_pnl_summary(f: Filters) -> pd.DataFrame:
    fv = current_data_version("financial")
    if fv is None:
        return _empty(["brand", "revenue", "gm", "promo_spend", "net_revenue"])
    where_f, params_f = _financial_where(f, fv)
    sql = f"""
        SELECT
            s.brand,
            SUM(f.revenue_actual)        AS revenue,
            SUM(f.gm_actual)             AS gm,
            SUM(f.promo_spend_actual)    AS promo_spend,
            SUM(f.revenue_actual - COALESCE(f.promo_spend_actual, 0)) AS net_revenue
        FROM fact_financial f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where_f}
        GROUP BY s.brand
        ORDER BY revenue DESC NULLS LAST
    """
    with _conn() as c:
        return c.execute(sql, params_f).fetchdf()


@st.cache_data(ttl=600, show_spinner=False)
def financial_ytd_progress(f: Filters) -> pd.DataFrame:
    """YTD actual vs full-year budget per brand. Fiscal year = calendar year (MVP).

    Note: this query uses f.period_to to anchor the fiscal year — f.period_from
    is intentionally ignored. Dim filters (brand/category/channel/sku) still apply.
    """
    fv = current_data_version("financial")
    if fv is None:
        return _empty(["brand", "ytd_actual", "full_year_budget", "ytd_pct"])

    year = f.period_to.year
    year_start = date(year, 1, 1)
    year_end   = date(year, 12, 1)  # first-of-month grain; covers all 12 rows

    # Reuse _financial_where via temp Filters with explicit periods.
    ytd_filter = Filters(
        period_from=year_start, period_to=f.period_to,
        brands=f.brands, categories=f.categories,
        channels=f.channels, regions=f.regions, skus=f.skus,
    )
    fy_filter = Filters(
        period_from=year_start, period_to=year_end,
        brands=f.brands, categories=f.categories,
        channels=f.channels, regions=f.regions, skus=f.skus,
    )
    where_ytd, params_ytd = _financial_where(ytd_filter, fv)
    where_fy,  params_fy  = _financial_where(fy_filter,  fv)

    sql = f"""
        WITH ytd AS (
            SELECT s.brand AS brand,
                   SUM(f.revenue_actual) AS ytd_actual
            FROM fact_financial f
            LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
            WHERE {where_ytd}
            GROUP BY s.brand
        ),
        fy AS (
            SELECT s.brand AS brand,
                   SUM(f.revenue_budget) AS full_year_budget
            FROM fact_financial f
            LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
            WHERE {where_fy}
            GROUP BY s.brand
        )
        SELECT
            COALESCE(y.brand, fy.brand)                    AS brand,
            COALESCE(y.ytd_actual, 0)                      AS ytd_actual,
            COALESCE(fy.full_year_budget, 0)               AS full_year_budget,
            CASE WHEN fy.full_year_budget IS NULL OR fy.full_year_budget = 0
                 THEN NULL
                 ELSE COALESCE(y.ytd_actual, 0) / fy.full_year_budget * 100
            END                                            AS ytd_pct
        FROM ytd y FULL OUTER JOIN fy ON y.brand = fy.brand
        ORDER BY full_year_budget DESC NULLS LAST
    """
    with _conn() as c:
        return c.execute(sql, params_ytd + params_fy).fetchdf()


@st.cache_data(ttl=600, show_spinner=False)
def financial_gm_trend(f: Filters) -> pd.DataFrame:
    """Monthly GM% across the filter window, for the KPI sparkline (SPEC §5.3)."""
    fv = current_data_version("financial")
    if fv is None:
        return _empty(["period_date", "gm_pct"])
    where_f, params_f = _financial_where(f, fv)
    sql = f"""
        SELECT
            f.period_date,
            SUM(f.gm_actual) / NULLIF(SUM(f.revenue_actual), 0) * 100 AS gm_pct
        FROM fact_financial f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where_f}
        GROUP BY f.period_date
        ORDER BY f.period_date
    """
    with _conn() as c:
        return c.execute(sql, params_f).fetchdf()


# ---------------------------------------------------------------------------
# Scorecard (ttl=600 — rule §3.2)
# ---------------------------------------------------------------------------
# Six tier-1 KPIs with heterogeneous threshold semantics (CONTRACTS §5.2,
# SPEC §5.4). All threshold-comparison logic lives inside this function; pages
# only render `value`, `rag`, and the `threshold_explainer` string.


@st.cache_data(ttl=600, show_spinner=False)
def scorecard_all(f: Filters) -> pd.DataFrame:
    """Tier-1 KPI scorecard — one row per KPI (SPEC §5.4).

    DOS reuses the demand-weighted aggregate from ``supply_kpis``
    (CONTRACTS §7.3). Revenue vs Budget is shown as achievement %
    (actual / budget × 100) so the existing ``RAG['rev_vs_budget']``
    thresholds (98 / 95) apply directly. GM% vs Budget is a pp delta.
    """
    nan = float("nan")
    rows: list[dict] = []

    # 1 & 2 — Forecast Accuracy and Forecast Bias (demand)
    dem = demand_kpis(f)
    rows.append({
        "kpi_name": "Forecast Accuracy",
        "value": dem.fa_pct,
        "unit": "%",
        "rag": _rag_for_fa(dem.fa_pct),
        "threshold_explainer": (
            f"≥ {config.RAG['fa']['green']}% green · "
            f"≥ {config.RAG['fa']['amber']}% amber"
        ),
    })
    rows.append({
        "kpi_name": "Forecast Bias",
        "value": dem.bias,
        "unit": "%",
        "rag": _rag_for_bias(dem.bias),
        "threshold_explainer": (
            f"|bias| ≤ {config.RAG['bias']['green_abs']}% green · "
            f"≤ {config.RAG['bias']['amber_abs']}% amber"
        ),
    })

    # 3 & 4 — Days of Supply and Fill Rate (supply)
    sup = supply_kpis(f)
    rows.append({
        "kpi_name": "Days of Supply",
        "value": sup.dos_avg,
        "unit": "days",
        "rag": _rag_for_dos(sup.dos_avg),
        "threshold_explainer": (
            f"target {config.RAG['dos']['target_days']}d "
            f"± {config.RAG['dos']['green_band']} green · "
            f"± {config.RAG['dos']['amber_band']} amber"
        ),
    })
    rows.append({
        "kpi_name": "Fill Rate",
        "value": sup.fill_rate,
        "unit": "%",
        "rag": _rag_for_fill_rate(sup.fill_rate),
        "threshold_explainer": (
            f"≥ {config.RAG['fill_rate']['green']}% green · "
            f"≥ {config.RAG['fill_rate']['amber']}% amber"
        ),
    })

    # 5 — Revenue vs Budget (achievement %)
    fin = financial_kpis(f)
    achievement = (
        nan if pd.isna(fin.revenue_vs_budget_pct)
        else 100.0 + fin.revenue_vs_budget_pct
    )
    rows.append({
        "kpi_name": "Revenue vs Budget",
        "value": achievement,
        "unit": "%",
        "rag": _rag_for_rev_vs_budget(achievement),
        "threshold_explainer": (
            f"≥ {config.RAG['rev_vs_budget']['green']}% of budget green · "
            f"≥ {config.RAG['rev_vs_budget']['amber']}% amber"
        ),
    })

    # 6 — GM% vs Budget (pp delta). Need GM%(actual) and GM%(budget); the
    # FinancialKPIs dataclass only carries GM% actual, so compute the pp
    # delta directly here from the same filter window.
    fv = current_data_version("financial")
    gm_delta_pp = nan
    if fv is not None:
        where_f, params_f = _financial_where(f, fv)
        sql = f"""
            SELECT
                SUM(f.gm_actual) / NULLIF(SUM(f.revenue_actual), 0) * 100
                  - SUM(f.gm_budget) / NULLIF(SUM(f.revenue_budget), 0) * 100
                    AS pp_delta
            FROM fact_financial f
            LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
            WHERE {where_f}
        """
        with _conn() as c:
            row = c.execute(sql, params_f).fetchone()
        if row and row[0] is not None:
            gm_delta_pp = float(row[0])
    rows.append({
        "kpi_name": "GM% vs Budget",
        "value": gm_delta_pp,
        "unit": "pp",
        "rag": _rag_for_gm_vs_budget(gm_delta_pp),
        "threshold_explainer": (
            f"≥ {config.RAG['gm_vs_budget_pp']['green']}pp green · "
            f"≥ {config.RAG['gm_vs_budget_pp']['amber']}pp amber"
        ),
    })

    return pd.DataFrame(rows, columns=[
        "kpi_name", "value", "unit", "rag", "threshold_explainer",
    ])


# ---------------------------------------------------------------------------
# Scenario (Phase 7) — see CONTRACTS §5.2 / §7.4
# ---------------------------------------------------------------------------
# Both queries support the Scenario Planner (pages/6_Scenario.py) and the
# Exec Summary's 3-scenario block. They are pure read functions; the % volume
# adjustments live in session_state and are applied by the page, not here.


@st.cache_data(ttl=600, show_spinner=False)
def scenario_baseline(f: Filters) -> pd.DataFrame:
    """Per-(sku, channel) consensus volume over the FORWARD portion of f's window.

    "Forward" = period_dates strictly after the latest period that carries
    actuals in active fact_demand (i.e. the planning horizon, not history).
    If no actuals exist yet, the full window is treated as forward.
    Volume is UOM-normalized to cases (rule §3.9).
    """
    dv = current_data_version("demand")
    if dv is None:
        return _empty(["sku_code", "channel_code", "consensus_volume_cases"])

    with _conn() as c:
        row = c.execute(
            "SELECT MAX(period_date) FROM fact_demand "
            "WHERE data_version = ? AND actuals IS NOT NULL",
            [dv],
        ).fetchone()
    latest_actuals = row[0] if row else None

    # Build a forward-only filter by lifting period_from above the actuals cut.
    if latest_actuals is None:
        forward_from = f.period_from
    else:
        forward_from = (
            pd.Timestamp(latest_actuals) + pd.DateOffset(months=1)
        ).date().replace(day=1)
        # If the forward edge is after the window already, nothing to project.
        if forward_from > f.period_to:
            return _empty(["sku_code", "channel_code", "consensus_volume_cases"])
        if forward_from < f.period_from:
            forward_from = f.period_from

    fwd = Filters(
        period_from=forward_from, period_to=f.period_to,
        brands=f.brands, categories=f.categories,
        channels=f.channels, regions=f.regions, skus=f.skus,
    )
    where, params = _demand_where(fwd, dv)
    sql = f"""
        SELECT
            f.sku_code,
            f.channel_code,
            SUM(f.consensus_fcst * COALESCE(s.uom_to_cases, 1))
                AS consensus_volume_cases
        FROM fact_demand f
        LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
        WHERE {where}
        GROUP BY f.sku_code, f.channel_code
        HAVING SUM(f.consensus_fcst) IS NOT NULL
        ORDER BY f.sku_code, f.channel_code
    """
    with _conn() as c:
        return c.execute(sql, params).fetchdf()


@st.cache_data(ttl=600, show_spinner=False)
def scenario_price_proxy(f: Filters) -> pd.DataFrame:
    """Per-(sku, channel) trailing-3M average price (currency per case).

    Window: the 3 calendar months ending at the **latest actuals period
    ≤ f.period_to** (CONTRACTS §7.4). Anchoring to actuals (not the calendar
    period_to) keeps the proxy meaningful when f.period_to extends into the
    forward planning horizon, where neither demand actuals nor financial
    revenue exist yet. (sku, channel) pairs with no actuals/revenue in the
    window are omitted; the page surfaces the count.
    """
    fv = current_data_version("financial")
    dv = current_data_version("demand")
    if fv is None or dv is None:
        return _empty(["sku_code", "channel_code", "avg_price"])

    # Find the latest actuals period ≤ f.period_to. If none exists in window
    # (e.g. all of f's window is forward-of-actuals), no proxy is possible.
    with _conn() as c:
        row = c.execute(
            "SELECT MAX(period_date) FROM fact_demand "
            "WHERE data_version = ? AND actuals IS NOT NULL "
            "AND period_date <= ?",
            [dv, f.period_to],
        ).fetchone()
    actuals_anchor = row[0] if row else None
    if actuals_anchor is None:
        return _empty(["sku_code", "channel_code", "avg_price"])

    window_to = pd.Timestamp(actuals_anchor).to_period("M").to_timestamp().date()
    window_from = (
        pd.Timestamp(window_to) - pd.DateOffset(months=2)
    ).date().replace(day=1)

    # Reuse _demand_where to pick up brand/category/channel/region/sku filters
    # against a forced 3-month window. Region filter is OK here because
    # fact_demand carries region_code.
    proxy_filter = Filters(
        period_from=window_from, period_to=window_to,
        brands=f.brands, categories=f.categories,
        channels=f.channels, regions=f.regions, skus=f.skus,
    )
    where_d, params_d = _demand_where(proxy_filter, dv)

    # Join semantics: aggregate fact_demand → (period, sku, channel) actuals in
    # cases; aggregate fact_financial → (period, sku, channel) revenue_actual.
    # Then SUM(revenue) / SUM(actuals_cases) per (sku, channel).
    sql = f"""
        WITH demand_agg AS (
            SELECT f.period_date, f.sku_code, f.channel_code,
                   SUM(f.actuals * COALESCE(s.uom_to_cases, 1)) AS actuals_cases
            FROM fact_demand f
            LEFT JOIN dim_sku s ON f.sku_code = s.sku_code
            WHERE {where_d}
            GROUP BY f.period_date, f.sku_code, f.channel_code
        ),
        financial_agg AS (
            SELECT period_date, sku_code, channel_code,
                   SUM(revenue_actual) AS rev_actual
            FROM fact_financial
            WHERE data_version = ?
              AND period_date BETWEEN ? AND ?
            GROUP BY period_date, sku_code, channel_code
        )
        SELECT
            d.sku_code,
            d.channel_code,
            SUM(fa.rev_actual) / NULLIF(SUM(d.actuals_cases), 0) AS avg_price
        FROM demand_agg d
        JOIN financial_agg fa
          ON fa.period_date  = d.period_date
         AND fa.sku_code     = d.sku_code
         AND fa.channel_code = d.channel_code
        GROUP BY d.sku_code, d.channel_code
        HAVING SUM(d.actuals_cases) IS NOT NULL
           AND SUM(d.actuals_cases) > 0
           AND SUM(fa.rev_actual) IS NOT NULL
        ORDER BY d.sku_code, d.channel_code
    """
    with _conn() as c:
        return c.execute(sql, params_d + [fv, window_from, window_to]).fetchdf()


# ---------------------------------------------------------------------------
# Dimension lookups (ttl=3600 — rule §3.2)
# ---------------------------------------------------------------------------


@st.cache_data(ttl=3600, show_spinner=False)
def list_brands() -> list[str]:
    with _conn() as c:
        rows = c.execute(
            "SELECT DISTINCT brand FROM dim_sku "
            "WHERE brand IS NOT NULL ORDER BY brand"
        ).fetchall()
    return [r[0] for r in rows]


@st.cache_data(ttl=3600, show_spinner=False)
def list_categories() -> list[str]:
    with _conn() as c:
        rows = c.execute(
            "SELECT DISTINCT category FROM dim_sku "
            "WHERE category IS NOT NULL ORDER BY category"
        ).fetchall()
    return [r[0] for r in rows]


@st.cache_data(ttl=3600, show_spinner=False)
def list_channels() -> list[str]:
    with _conn() as c:
        rows = c.execute(
            "SELECT channel_code FROM dim_channel ORDER BY channel_code"
        ).fetchall()
    return [r[0] for r in rows]


@st.cache_data(ttl=3600, show_spinner=False)
def list_regions() -> list[str]:
    with _conn() as c:
        rows = c.execute(
            "SELECT region_code FROM dim_region ORDER BY region_code"
        ).fetchall()
    return [r[0] for r in rows]


@st.cache_data(ttl=3600, show_spinner=False)
def list_skus() -> pd.DataFrame:
    with _conn() as c:
        return c.execute(
            "SELECT sku_code, sku_name FROM dim_sku ORDER BY sku_code"
        ).fetchdf()
