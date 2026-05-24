"""CSV → DuckDB ingestion.

Implements the contracts in CONTRACTS.md §5.1: detect_schema, parse, validate,
load, list_uploads. The only module that opens the DuckDB file for writes —
pages and exporters must go through query.py (read) or this module (write).
"""

from __future__ import annotations

import io
import os
import re
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import datetime

import duckdb
import pandas as pd

import config

# ---------------------------------------------------------------------------
# Dataclasses (CONTRACTS §5.1 / SPEC §9.1)
# ---------------------------------------------------------------------------


@dataclass
class DetectedColumn:
    name: str
    detected_type: str            # 'date' | 'int' | 'float' | 'text' | 'bool'
    sample_values: list[str]
    suggested_field: str | None   # canonical field name or None


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


# ---------------------------------------------------------------------------
# Per-domain schema (canonical fields, types, PK, target table)
# ---------------------------------------------------------------------------

DOMAINS = ("demand", "supply", "financial",
           "sku_master", "channel_master", "region_master", "plant_master")

# canonical field → target table column
_SCHEMA = {
    "demand": {
        "table": "fact_demand",
        "required": ["period_date", "sku_code", "channel_code",
                     "statistical_fcst", "consensus_fcst"],
        "optional": ["region_code", "actuals"],
        "numeric":  ["statistical_fcst", "consensus_fcst", "actuals"],
        "date":     ["period_date"],
        "pk":       ["period_date", "sku_code", "channel_code", "region_code"],
        "defaults": {"region_code": "ALL"},
    },
    "supply": {
        "table": "fact_supply",
        "required": ["period_date", "sku_code", "plant_code",
                     "inventory_qty", "production_plan", "production_actual"],
        "optional": ["capacity_plan", "orders_requested", "orders_delivered"],
        "numeric":  ["inventory_qty", "production_plan", "production_actual",
                     "capacity_plan", "orders_requested", "orders_delivered"],
        "date":     ["period_date"],
        "pk":       ["period_date", "sku_code", "plant_code"],
        "defaults": {},
    },
    "financial": {
        "table": "fact_financial",
        "required": ["period_date", "sku_code", "channel_code",
                     "revenue_actual", "revenue_budget", "revenue_le", "gm_actual"],
        "optional": ["gm_budget", "promo_spend_actual",
                     "promo_spend_budget", "currency_code"],
        "numeric":  ["revenue_actual", "revenue_budget", "revenue_le",
                     "gm_actual", "gm_budget",
                     "promo_spend_actual", "promo_spend_budget"],
        "date":     ["period_date"],
        "pk":       ["period_date", "sku_code", "channel_code"],
        "defaults": {"currency_code": config.CURRENCY},
    },
    "sku_master": {
        "table": "dim_sku",
        "required": ["sku_code", "sku_name", "brand", "category", "uom"],
        "optional": ["subcategory", "uom_to_cases"],
        "numeric":  ["uom_to_cases"],
        "date":     [],
        "pk":       ["sku_code"],
        "defaults": {"uom_to_cases": 1.0, "subcategory": None},
    },
    "channel_master": {
        "table": "dim_channel",
        "required": ["channel_code", "channel_name"],
        "optional": ["channel_type"],
        "numeric":  [],
        "date":     [],
        "pk":       ["channel_code"],
        "defaults": {"channel_type": None},
    },
    "region_master": {
        "table": "dim_region",
        "required": ["region_code", "region_name", "country"],
        "optional": ["cluster"],
        "numeric":  [],
        "date":     [],
        "pk":       ["region_code"],
        "defaults": {"cluster": None},
    },
    "plant_master": {
        "table": "dim_plant",
        "required": ["plant_code", "plant_name"],
        "optional": ["country"],
        "numeric":  [],
        "date":     [],
        "pk":       ["plant_code"],
        "defaults": {"country": None},
    },
}

# Header-to-canonical synonyms for suggested_field. Exact-after-normalization match.
_SYNONYMS: dict[str, str] = {
    # dates
    "period": "period_date", "month": "period_date", "date": "period_date",
    "period_month": "period_date", "year_month": "period_date", "yyyymm": "period_date",
    # dims
    "sku": "sku_code", "material": "sku_code", "item": "sku_code",
    "sku_id": "sku_code", "item_code": "sku_code",
    "channel": "channel_code", "channel_id": "channel_code",
    "region": "region_code", "region_id": "region_code", "geo": "region_code",
    "plant": "plant_code", "site": "plant_code", "factory": "plant_code",
    # forecasts
    "stat_fcst": "statistical_fcst", "statistical_forecast": "statistical_fcst",
    "fcst_stat": "statistical_fcst",
    "cons_fcst": "consensus_fcst", "consensus_forecast": "consensus_fcst",
    "fcst_cons": "consensus_fcst", "demand_plan": "consensus_fcst",
    "actual": "actuals", "actual_volume": "actuals", "sales_actual": "actuals",
    # supply
    "inventory": "inventory_qty", "stock": "inventory_qty", "on_hand": "inventory_qty",
    "prod_plan": "production_plan", "production_planned": "production_plan",
    "prod_actual": "production_actual", "production": "production_actual",
    "capacity": "capacity_plan",
    "orders_req": "orders_requested", "orders_demand": "orders_requested",
    "orders_del": "orders_delivered", "orders_shipped": "orders_delivered",
    # financial
    "revenue": "revenue_actual", "net_revenue": "revenue_actual",
    "rev_actual": "revenue_actual",
    "rev_budget": "revenue_budget", "budget": "revenue_budget",
    "rev_le": "revenue_le", "le": "revenue_le", "latest_estimate": "revenue_le",
    "gm": "gm_actual", "gross_margin": "gm_actual",
    "promo": "promo_spend_actual", "promo_spend": "promo_spend_actual",
    "currency": "currency_code",
    # sku master
    "sku_description": "sku_name", "description": "sku_name", "name": "sku_name",
    "brand_name": "brand", "category_name": "category",
    "unit_of_measure": "uom", "unit": "uom",
    "conversion_factor": "uom_to_cases", "case_factor": "uom_to_cases",
    # channel/region/plant master
    "channel_description": "channel_name",
    "region_description": "region_name",
    "plant_description": "plant_name",
    "country_code": "country", "country_name": "country",
}


def _norm(s: str) -> str:
    """Normalize header for synonym lookup: lowercase, alphanum + underscores."""
    return re.sub(r"[^a-z0-9]+", "_", str(s).strip().lower()).strip("_")


def _all_canonical_fields() -> set[str]:
    out: set[str] = set()
    for spec in _SCHEMA.values():
        out.update(spec["required"])
        out.update(spec["optional"])
    return out


# ---------------------------------------------------------------------------
# DuckDB connection (lazy; runs schema.sql on first open)
# ---------------------------------------------------------------------------


def _get_conn() -> duckdb.DuckDBPyConnection:
    conn = duckdb.connect(config.DUCKDB_PATH)
    with open(config.SCHEMA_SQL_PATH, "r", encoding="utf-8") as f:
        conn.execute(f.read())
    return conn


# ---------------------------------------------------------------------------
# detect_schema
# ---------------------------------------------------------------------------

_ENCODINGS = ("utf-8", "utf-16", "latin-1")
_DATE_FORMATS = (
    "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y",
    "%Y%m%d", "%Y-%m", "%Y/%m", "%Y%m",
)


def _detect_encoding(file_bytes: bytes) -> str:
    """Try common CSV encodings in order; return the first that fully decodes."""
    for enc in _ENCODINGS:
        try:
            file_bytes.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    return "latin-1"  # safest fallback (single-byte, always decodes)


def _infer_column_type(s: pd.Series) -> str:
    """Return one of 'date' | 'int' | 'float' | 'bool' | 'text'."""
    sample = s.dropna().astype(str).str.strip()
    sample = sample[sample != ""]
    if len(sample) == 0:
        return "text"

    # bool
    bool_set = {"true", "false", "0", "1", "yes", "no"}
    if sample.str.lower().isin(bool_set).all() and sample.str.lower().nunique() > 1:
        return "bool"

    # date — try a few common formats; >= 80% parseable counts
    for fmt in _DATE_FORMATS:
        parsed = pd.to_datetime(sample, format=fmt, errors="coerce")
        if parsed.notna().mean() >= 0.8:
            return "date"
    parsed = pd.to_datetime(sample, errors="coerce", format="mixed")
    if parsed.notna().mean() >= 0.8:
        return "date"

    # numeric
    numeric = pd.to_numeric(sample.str.replace(",", "", regex=False), errors="coerce")
    if numeric.notna().mean() >= 0.95:
        if (numeric.dropna() % 1 == 0).all():
            return "int"
        return "float"

    return "text"


def _suggest_field(col_name: str) -> str | None:
    n = _norm(col_name)
    if n in _all_canonical_fields():
        return n
    return _SYNONYMS.get(n)


def detect_schema(file_bytes: bytes) -> SchemaReport:
    encoding = _detect_encoding(file_bytes)
    buf = io.BytesIO(file_bytes)
    head = pd.read_csv(buf, encoding=encoding, nrows=1000, dtype=str, keep_default_na=False)

    # Row count: cheap full pass with chunked reading
    buf.seek(0)
    total_rows = sum(
        len(chunk) for chunk in pd.read_csv(
            buf, encoding=encoding, dtype=str, chunksize=50_000, keep_default_na=False
        )
    )

    columns: list[DetectedColumn] = []
    for col in head.columns:
        ser = head[col]
        samples = ser.head(5).tolist()
        columns.append(DetectedColumn(
            name=str(col),
            detected_type=_infer_column_type(ser),
            sample_values=[str(v) for v in samples],
            suggested_field=_suggest_field(col),
        ))
    return SchemaReport(row_count=total_rows, columns=columns, encoding_used=encoding)


# ---------------------------------------------------------------------------
# parse
# ---------------------------------------------------------------------------


def parse(file_bytes: bytes, encoding: str) -> pd.DataFrame:
    return pd.read_csv(io.BytesIO(file_bytes), encoding=encoding, dtype=str,
                       keep_default_na=False)


# ---------------------------------------------------------------------------
# validate
# ---------------------------------------------------------------------------


def _coerce_date(series: pd.Series) -> pd.Series:
    """Best-effort date parser across SPEC §3.1 accepted formats."""
    out = pd.to_datetime(series, errors="coerce", format="mixed")
    if out.isna().any():
        for fmt in _DATE_FORMATS:
            mask = out.isna()
            if not mask.any():
                break
            attempt = pd.to_datetime(series[mask], format=fmt, errors="coerce")
            out.loc[mask] = attempt
    return out


def validate(df: pd.DataFrame, domain: str, mapping: dict[str, str]) -> ValidationReport:
    if domain not in _SCHEMA:
        return ValidationReport(issues=[ValidationIssue(
            severity="error", row=None, column=None,
            message=f"Unknown domain '{domain}'",
        )])

    spec = _SCHEMA[domain]
    report = ValidationReport()
    mapped_fields = set(mapping.values())

    # Required fields present
    for fld in spec["required"]:
        if fld not in mapped_fields:
            report.issues.append(ValidationIssue(
                severity="error", row=None, column=None,
                message=f"Required field '{fld}' not mapped",
            ))
    if not report.can_load:
        return report  # fail fast — row checks would be meaningless

    # Inverse map: canonical → csv col
    canon_to_csv = {v: k for k, v in mapping.items()}

    # Required values not null
    for fld in spec["required"]:
        csv_col = canon_to_csv[fld]
        if csv_col not in df.columns:
            report.issues.append(ValidationIssue(
                severity="error", row=None, column=csv_col,
                message=f"CSV column '{csv_col}' not present in file",
            ))
            continue
        null_mask = df[csv_col].isna() | (df[csv_col].astype(str).str.strip() == "")
        for idx in df.index[null_mask][:5]:
            report.issues.append(ValidationIssue(
                severity="error", row=int(idx),
                column=csv_col,
                message=f"Required field '{fld}' is empty",
            ))
        if null_mask.sum() > 5:
            report.issues.append(ValidationIssue(
                severity="error", row=None, column=csv_col,
                message=f"... and {int(null_mask.sum()) - 5} more empty values in '{fld}'",
            ))

    # Date parseability
    for fld in spec["date"]:
        if fld in canon_to_csv:
            csv_col = canon_to_csv[fld]
            if csv_col in df.columns:
                parsed = _coerce_date(df[csv_col])
                bad = parsed.isna() & df[csv_col].astype(str).str.strip().ne("")
                for idx in df.index[bad][:5]:
                    report.issues.append(ValidationIssue(
                        severity="error", row=int(idx), column=csv_col,
                        message=f"Cannot parse '{df.at[idx, csv_col]}' as date",
                    ))

    # Numeric type
    for fld in spec["numeric"]:
        if fld in canon_to_csv:
            csv_col = canon_to_csv[fld]
            if csv_col in df.columns:
                cleaned = df[csv_col].astype(str).str.replace(",", "", regex=False)
                numeric = pd.to_numeric(cleaned, errors="coerce")
                bad = numeric.isna() & df[csv_col].astype(str).str.strip().ne("")
                for idx in df.index[bad][:3]:
                    report.issues.append(ValidationIssue(
                        severity="warning", row=int(idx), column=csv_col,
                        message=f"Non-numeric value '{df.at[idx, csv_col]}' in '{fld}'",
                    ))

    # Cross-reference dim codes (warn only — load is allowed even with new codes)
    conn = _get_conn()
    try:
        _xref(report, conn, df, canon_to_csv, "sku_code",     "dim_sku",     "sku_code")
        _xref(report, conn, df, canon_to_csv, "channel_code", "dim_channel", "channel_code")
        _xref(report, conn, df, canon_to_csv, "region_code",  "dim_region",  "region_code")
        _xref(report, conn, df, canon_to_csv, "plant_code",   "dim_plant",   "plant_code")
    finally:
        conn.close()

    # Duplicate PKs in the file
    pk_csv_cols = [canon_to_csv[c] for c in spec["pk"] if c in canon_to_csv]
    if pk_csv_cols and all(c in df.columns for c in pk_csv_cols):
        dup_mask = df.duplicated(subset=pk_csv_cols, keep=False)
        if dup_mask.any():
            report.issues.append(ValidationIssue(
                severity="warning", row=None, column=None,
                message=(f"{int(dup_mask.sum())} duplicate rows on primary key "
                         f"{spec['pk']}; last occurrence wins on load"),
            ))

    return report


def _xref(report: ValidationReport, conn, df: pd.DataFrame,
          canon_to_csv: dict[str, str], canon_field: str,
          dim_table: str, dim_col: str) -> None:
    if canon_field not in canon_to_csv:
        return
    csv_col = canon_to_csv[canon_field]
    if csv_col not in df.columns:
        return
    file_codes = set(df[csv_col].dropna().astype(str).str.strip()) - {""}
    if not file_codes:
        return
    existing = set(c[0] for c in conn.execute(
        f"SELECT {dim_col} FROM {dim_table}"
    ).fetchall())
    new_codes = file_codes - existing
    if new_codes and existing:  # don't warn on first-ever upload
        sample = sorted(new_codes)[:5]
        report.issues.append(ValidationIssue(
            severity="warning", row=None, column=csv_col,
            message=(f"{len(new_codes)} {canon_field}(s) not in {dim_table}: "
                     f"{', '.join(sample)}{'…' if len(new_codes) > 5 else ''}"),
        ))


# ---------------------------------------------------------------------------
# load (transactional)
# ---------------------------------------------------------------------------


def _prepare_df(df: pd.DataFrame, domain: str, mapping: dict[str, str]) -> pd.DataFrame:
    """Rename CSV→canonical, coerce types, normalize dates, fill defaults.

    Output columns are exactly the target table's column list (excluding
    data_version + loaded_at, which load() appends).
    """
    spec = _SCHEMA[domain]
    canon_to_csv = {v: k for k, v in mapping.items()}

    out_cols = spec["required"] + spec["optional"]
    out = pd.DataFrame(index=df.index)

    for fld in out_cols:
        if fld in canon_to_csv and canon_to_csv[fld] in df.columns:
            out[fld] = df[canon_to_csv[fld]]
        elif fld in spec["defaults"]:
            out[fld] = spec["defaults"][fld]
        else:
            out[fld] = None

    for fld in spec["date"]:
        out[fld] = _coerce_date(out[fld]).dt.to_period("M").dt.to_timestamp()

    for fld in spec["numeric"]:
        if fld in out.columns:
            out[fld] = pd.to_numeric(
                out[fld].astype(str).str.replace(",", "", regex=False),
                errors="coerce",
            )

    pk_cols = [c for c in spec["pk"] if c in out.columns]
    if pk_cols:
        out = out.drop_duplicates(subset=pk_cols, keep="last").reset_index(drop=True)

    return out


def _next_data_version(conn, domain: str) -> int:
    row = conn.execute(
        "SELECT COALESCE(MAX(data_version), 0) FROM upload_log WHERE domain = ?",
        [domain],
    ).fetchone()
    return int(row[0]) + 1


def _supersede_active(conn, domain: str) -> None:
    conn.execute(
        "UPDATE upload_log SET status = 'superseded' "
        "WHERE domain = ? AND status = 'active'",
        [domain],
    )


def _current_active_version(conn, domain: str) -> int | None:
    row = conn.execute(
        "SELECT data_version FROM upload_log "
        "WHERE domain = ? AND status = 'active' "
        "ORDER BY data_version DESC LIMIT 1",
        [domain],
    ).fetchone()
    return int(row[0]) if row else None


def _is_fact(domain: str) -> bool:
    return domain in ("demand", "supply", "financial")


def _insert_fact(conn, table: str, df: pd.DataFrame, version: int) -> None:
    df = df.copy()
    df["data_version"] = version
    df["loaded_at"] = datetime.utcnow()
    cols = list(df.columns)
    col_list = ", ".join(cols)
    conn.register("_load_df", df)
    conn.execute(f"INSERT INTO {table} ({col_list}) SELECT {col_list} FROM _load_df")
    conn.unregister("_load_df")


def _insert_dim(conn, table: str, df: pd.DataFrame, pk: list[str], mode: str) -> None:
    """Upsert rows into a dim table. For mode='replace', also delete any
    pre-existing PKs not in the new file.

    Why upsert instead of DELETE+INSERT: DuckDB's unique index used to enforce
    PRIMARY KEY does not see deletions until commit, so DELETE + INSERT in the
    same transaction throws a spurious "duplicate key" error when the new file
    overlaps the existing PKs. Upsert side-steps this.
    """
    df = df.copy()
    df["loaded_at"] = datetime.utcnow()
    cols = list(df.columns)
    col_list = ", ".join(cols)
    non_pk = [c for c in cols if c not in pk]
    pk_list = ", ".join(pk)

    conn.register("_load_df", df)
    try:
        if non_pk:
            set_clause = ", ".join(f"{c} = EXCLUDED.{c}" for c in non_pk)
            conn.execute(
                f"INSERT INTO {table} ({col_list}) SELECT {col_list} FROM _load_df "
                f"ON CONFLICT ({pk_list}) DO UPDATE SET {set_clause}"
            )
        else:
            conn.execute(
                f"INSERT INTO {table} ({col_list}) SELECT {col_list} FROM _load_df "
                f"ON CONFLICT ({pk_list}) DO NOTHING"
            )

        if mode == "replace":
            # Drop rows whose PK isn't in the new file. Single-column PK uses
            # an IN clause; composite PK uses a tuple form.
            if len(pk) == 1:
                conn.execute(
                    f"DELETE FROM {table} "
                    f"WHERE {pk[0]} NOT IN (SELECT {pk[0]} FROM _load_df)"
                )
            else:
                conn.execute(
                    f"DELETE FROM {table} "
                    f"WHERE ({pk_list}) NOT IN (SELECT {pk_list} FROM _load_df)"
                )
    finally:
        conn.unregister("_load_df")


def load(
    df: pd.DataFrame,
    domain: str,
    mapping: dict[str, str],
    mode: str,
    filename: str,
    file_bytes: bytes,
) -> int:
    if domain not in _SCHEMA:
        raise ValueError(f"Unknown domain '{domain}'")
    if mode not in ("replace", "append"):
        raise ValueError(f"Unknown mode '{mode}'")

    spec = _SCHEMA[domain]
    prepared = _prepare_df(df, domain, mapping)
    conn = _get_conn()

    try:
        conn.execute("BEGIN TRANSACTION")
        new_version = _next_data_version(conn, domain)

        if _is_fact(domain):
            prior_version = _current_active_version(conn, domain)
            if mode == "append" and prior_version is not None:
                prior_rows = conn.execute(
                    f"SELECT * EXCLUDE (data_version, loaded_at) "
                    f"FROM {spec['table']} WHERE data_version = ?",
                    [prior_version],
                ).fetchdf()
                merged = pd.concat([prior_rows, prepared], ignore_index=True)
                merged = merged.drop_duplicates(subset=spec["pk"], keep="last")
                prepared = merged.reset_index(drop=True)

            _supersede_active(conn, domain)
            _insert_fact(conn, spec["table"], prepared, new_version)
        else:
            _supersede_active(conn, domain)
            _insert_dim(conn, spec["table"], prepared, spec["pk"], mode)

        conn.execute(
            "INSERT INTO upload_log "
            "(filename, domain, uploaded_at, row_count, data_version, status, error_msg) "
            "VALUES (?, ?, ?, ?, ?, 'active', NULL)",
            [filename, domain, datetime.utcnow(), len(prepared), new_version],
        )
        conn.execute("COMMIT")
    except Exception:
        conn.execute("ROLLBACK")
        conn.close()
        raise

    # Post-commit: copy raw file. Failure here doesn't roll back (rule §3.3).
    copy_error: str | None = None
    try:
        os.makedirs(config.RAW_DIR, exist_ok=True)
        safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", filename)
        target = os.path.join(
            config.RAW_DIR,
            f"{domain}_v{new_version}_{safe_name}",
        )
        with tempfile.NamedTemporaryFile(
            delete=False, dir=config.RAW_DIR, suffix=".part"
        ) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        shutil.move(tmp_path, target)
    except Exception as e:  # noqa: BLE001 — record & continue per rule §3.3
        copy_error = f"raw-file copy failed: {e}"

    if copy_error:
        conn.execute(
            "UPDATE upload_log SET error_msg = ? "
            "WHERE domain = ? AND data_version = ?",
            [copy_error, domain, new_version],
        )

    conn.close()
    return new_version


# ---------------------------------------------------------------------------
# list_uploads
# ---------------------------------------------------------------------------


def list_uploads() -> pd.DataFrame:
    conn = _get_conn()
    try:
        return conn.execute(
            "SELECT filename, domain, uploaded_at, row_count, "
            "data_version, status, error_msg "
            "FROM upload_log ORDER BY uploaded_at DESC"
        ).fetchdf()
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Helpers exposed for the Upload page
# ---------------------------------------------------------------------------


def canonical_fields_for(domain: str) -> dict[str, list[str]]:
    """Return {'required': [...], 'optional': [...]} for the column-mapper UI."""
    spec = _SCHEMA[domain]
    return {"required": list(spec["required"]), "optional": list(spec["optional"])}
