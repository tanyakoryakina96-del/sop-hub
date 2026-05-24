"""Upload page — CSV → detected schema → column mapping → validation → load.

The only page that writes to DuckDB; all writes go through `data.ingest`.
Session-state writes here populate `data_versions`, `last_upload_status`, and
`column_mapping_template` per CONTRACTS §4.
"""

import sys
from pathlib import Path

import pandas as pd
import streamlit as st

# Allow `import config` / `from data import ingest` when Streamlit launches
# this file directly (its parent is `pages/`, not the project root).
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import config  # noqa: F401, E402 — registers Plotly template
from data import ingest  # noqa: E402

st.set_page_config(page_title="Upload — S&OP Hub", page_icon="📥", layout="wide")
st.title("📥 Upload data")
st.caption(
    "CSV → detected schema → column mapping → validation → load. "
    "One active version per domain; prior uploads remain queryable for audit."
)

DOMAIN_LABELS = {
    "demand":         "Demand (forecast + actuals)",
    "supply":         "Supply (inventory + production)",
    "financial":      "Financial (revenue + margin)",
    "sku_master":     "SKU master",
    "channel_master": "Channel master",
    "region_master":  "Region master",
    "plant_master":   "Plant master",
}

# ---------------------------------------------------------------------------
# Step 1 — file + domain + mode
# ---------------------------------------------------------------------------

col1, col2, col3 = st.columns([3, 2, 2])
with col1:
    uploaded = st.file_uploader("CSV file", type=["csv"], accept_multiple_files=False)
with col2:
    domain = st.selectbox(
        "Domain", list(DOMAIN_LABELS.keys()),
        format_func=lambda d: DOMAIN_LABELS[d],
    )
with col3:
    mode = st.radio(
        "Mode", ["replace", "append"], horizontal=True,
        help=(
            "**Replace** marks the prior active upload as superseded; **Append** "
            "merges new rows into the prior active version, deduping on PK."
        ),
    )

if uploaded is None:
    st.info("Choose a CSV to begin.")
    st.stop()

# Cache bytes + schema in session_state so widget reruns don't re-detect.
file_bytes = uploaded.getvalue()
cache_key = (uploaded.name, len(file_bytes), domain)
if st.session_state.get("_upload_cache_key") != cache_key:
    with st.spinner("Detecting schema…"):
        st.session_state["_upload_schema"] = ingest.detect_schema(file_bytes)
        st.session_state["_upload_cache_key"] = cache_key

schema: ingest.SchemaReport = st.session_state["_upload_schema"]

st.success(
    f"Detected **{schema.row_count:,} rows** · **{len(schema.columns)} columns** · "
    f"encoding `{schema.encoding_used}`"
)

# ---------------------------------------------------------------------------
# Step 2 — column mapping
# ---------------------------------------------------------------------------

fields = ingest.canonical_fields_for(domain)
canonical_options = ["(skip)"] + fields["required"] + fields["optional"]

templates = st.session_state.setdefault("column_mapping_template", {})
saved_template = templates.get(domain, {})  # csv_col → canonical_field

st.subheader("Column mapping")
st.caption(
    f"Map each CSV column to a canonical field. Required for **{domain}**: "
    f"`{', '.join(fields['required'])}`. Optional: "
    f"`{', '.join(fields['optional']) or '—'}`."
)

mapping: dict[str, str] = {}
for col in schema.columns:
    cols = st.columns([3, 2, 3, 3])
    cols[0].markdown(f"**{col.name}**")
    cols[1].caption(f"type: `{col.detected_type}`")
    cols[2].caption("samples: " + ", ".join(s for s in col.sample_values[:3]) or "—")

    default = (
        saved_template.get(col.name)
        or col.suggested_field
        or "(skip)"
    )
    if default not in canonical_options:
        default = "(skip)"
    selected = cols[3].selectbox(
        f"map_{col.name}",
        canonical_options,
        index=canonical_options.index(default),
        label_visibility="collapsed",
        key=f"mapsel_{domain}_{col.name}",
    )
    if selected != "(skip)":
        mapping[col.name] = selected

# Detect double-mapping (two CSV cols → same canonical field)
dupes = {f for f in mapping.values() if list(mapping.values()).count(f) > 1}
if dupes:
    st.error(f"Multiple CSV columns mapped to the same field(s): {sorted(dupes)}")
    st.stop()

missing_required = set(fields["required"]) - set(mapping.values())
if missing_required:
    st.warning(f"Still missing required field(s): {sorted(missing_required)}")

# ---------------------------------------------------------------------------
# Step 3 — validate
# ---------------------------------------------------------------------------

st.subheader("Validation")

@st.cache_data(show_spinner=False)
def _parse_cached(file_bytes: bytes, encoding: str) -> pd.DataFrame:
    return ingest.parse(file_bytes, encoding)


df = _parse_cached(file_bytes, schema.encoding_used)

if missing_required:
    st.info("Map all required fields above to run validation.")
    st.stop()

with st.spinner("Validating…"):
    report = ingest.validate(df, domain, mapping)

m1, m2, m3 = st.columns(3)
m1.metric("Errors", len(report.errors))
m2.metric("Warnings", len(report.warnings))
m3.metric("Can load", "✅" if report.can_load else "❌")

if report.issues:
    issue_df = pd.DataFrame(
        [{"severity": i.severity, "row": i.row, "column": i.column, "message": i.message}
         for i in report.issues]
    )
    st.dataframe(issue_df, use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Step 4 — load
# ---------------------------------------------------------------------------

st.subheader("Load")
btn_disabled = not report.can_load
if btn_disabled:
    st.error("Fix errors above before loading.")

if st.button("Load into DuckDB", type="primary", disabled=btn_disabled):
    try:
        with st.spinner(f"Loading {len(df):,} rows…"):
            new_version = ingest.load(
                df=df,
                domain=domain,
                mapping=mapping,
                mode=mode,
                filename=uploaded.name,
                file_bytes=file_bytes,
            )

        # Update session state per CONTRACTS §4
        versions = st.session_state.setdefault("data_versions", {})
        versions[domain] = new_version
        templates[domain] = dict(mapping)
        st.session_state["last_upload_status"] = (
            f"Loaded {uploaded.name} → {domain} v{new_version} ({mode})"
        )

        # Invalidate cached query results (rule §3.3)
        st.cache_data.clear()

        st.success(st.session_state["last_upload_status"])
        st.balloons()
    except Exception as exc:  # noqa: BLE001 — surface to user, no PII risk locally
        st.session_state["last_upload_status"] = f"Load failed: {exc}"
        st.error(st.session_state["last_upload_status"])

# Persisted banner across reruns
if (status := st.session_state.get("last_upload_status")) and "Loaded" in status:
    st.info(status)

# ---------------------------------------------------------------------------
# Upload history
# ---------------------------------------------------------------------------

with st.expander("📜 Upload history", expanded=False):
    history = ingest.list_uploads()
    if history.empty:
        st.caption("No uploads yet.")
    else:
        st.dataframe(history, use_container_width=True, hide_index=True)
