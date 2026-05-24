"""Per-Streamlit-session DuckDB management.

Each Streamlit session gets its own DuckDB file under ``storage/sessions/`` so
visitors of the deployed demo can upload data and run scenarios in parallel
without colliding. On first access in a session, the DB is seeded from
``test_data/*.csv`` so dashboards render immediately on the home screen.

Outside a Streamlit session (unit tests, ad-hoc scripts), :func:`get_db_path`
falls back to the legacy single-file path so existing tooling still works.

Public API consumed by ``data.query``, ``data.ingest``, ``exports.*`` and
``app.py``:

* :func:`get_db_path` — replaces references to the old ``config.DUCKDB_PATH``
  constant. Always call this just before opening a DuckDB connection.
* :func:`seed_if_empty` — call once per Streamlit session (e.g. from
  ``app.py``) before the user touches any dashboard.
* :func:`cleanup_stale_sessions` — best-effort prune of orphaned session DB
  files on app boot. Streamlit Cloud's ephemeral disk reclaims everything on
  restart, but local dev would otherwise accumulate them.
"""

from __future__ import annotations

import os
import time
import uuid

import config

SESSIONS_DIR = os.path.join(config.STORAGE_DIR, "sessions")
os.makedirs(SESSIONS_DIR, exist_ok=True)

# Process-wide fallback used when get_db_path() runs outside a Streamlit
# session (CLI, tests). Same filename as the pre-refactor location so existing
# local DBs aren't orphaned.
_FALLBACK_DB_PATH = os.path.join(config.STORAGE_DIR, "sop.duckdb")

# Seed order: dims first so fact-load xref warnings stay silent on a fresh DB.
_SEED_FILES: tuple[tuple[str, str], ...] = (
    ("dim_sku.csv",        "sku_master"),
    ("dim_channel.csv",    "channel_master"),
    ("dim_region.csv",     "region_master"),
    ("dim_plant.csv",      "plant_master"),
    ("fact_demand.csv",    "demand"),
    ("fact_supply.csv",    "supply"),
    ("fact_financial.csv", "financial"),
)


def get_db_path() -> str:
    """Return the DuckDB path for the current Streamlit session.

    Inside a Streamlit script run, a per-session UUID is stored in
    ``st.session_state["session_db_id"]`` on first access, mapping to a unique
    file under :data:`SESSIONS_DIR`. Outside a Streamlit run, returns the
    process-wide fallback path.
    """
    try:
        import streamlit as st
        sid = st.session_state.get("session_db_id")
        if sid is None:
            sid = uuid.uuid4().hex
            st.session_state["session_db_id"] = sid
        return os.path.join(SESSIONS_DIR, f"sop_{sid}.duckdb")
    except (ImportError, RuntimeError, AttributeError):
        return _FALLBACK_DB_PATH


def seed_if_empty() -> None:
    """Populate the current session's DB from ``test_data/`` on first call.

    Idempotent — guarded by ``st.session_state["session_db_seeded"]``. Safe to
    call on every Streamlit script rerun. Does nothing outside a Streamlit
    session, or if any uploads already exist (so users who hit the Upload page
    first don't get their data overwritten on the next rerun).
    """
    try:
        import streamlit as st
    except ImportError:
        return

    if st.session_state.get("session_db_seeded"):
        return

    # Lazy: data.ingest pulls get_db_path() from this module, so a top-level
    # import here would create a cycle.
    from data import ingest, query

    if len(ingest.list_uploads()) > 0:
        st.session_state["session_db_seeded"] = True
        return

    test_dir = os.path.join(config.PROJECT_ROOT, "test_data")
    if not os.path.isdir(test_dir):
        st.session_state["session_db_seeded"] = True
        return

    for filename, domain in _SEED_FILES:
        path = os.path.join(test_dir, filename)
        if not os.path.exists(path):
            continue
        with open(path, "rb") as f:
            file_bytes = f.read()
        df = ingest.parse(file_bytes, encoding="utf-8")
        mapping = {c: c for c in df.columns}  # canonical headers — identity
        ingest.load(df, domain, mapping, mode="replace",
                    filename=filename, file_bytes=file_bytes)

    st.session_state["session_db_seeded"] = True
    st.session_state["data_versions"] = {
        d: query.current_data_version(d)
        for d in ("demand", "supply", "financial", "sku_master")
    }


_cleanup_done = False


def cleanup_stale_sessions(max_age_hours: float = 6) -> None:
    """Delete session DB files older than ``max_age_hours``.

    Runs at most once per Python process — Streamlit reruns the script per
    interaction, but module state persists across reruns.
    """
    global _cleanup_done
    if _cleanup_done:
        return
    _cleanup_done = True

    cutoff = time.time() - max_age_hours * 3600
    if not os.path.isdir(SESSIONS_DIR):
        return
    for entry in os.scandir(SESSIONS_DIR):
        if not entry.is_file():
            continue
        try:
            if entry.stat().st_mtime < cutoff:
                os.remove(entry.path)
        except OSError:
            pass
