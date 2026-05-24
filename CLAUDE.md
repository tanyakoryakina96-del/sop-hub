# S&OP Hub — Claude Orientation

A single-user S&OP reporting platform for a CPG company. Replaces a 2–3 day/month Excel→PowerPoint process with self-service dashboards and one-click exports.

## ⚠️ Critical context: this codebase has no human engineers

**This project is built and maintained by Claude only, via vibe coding.** The user is an S&OP Manager / Business Analyst — not a software engineer. They direct the work, review the output, and use the app. They do not write or debug code themselves.

That means **you are the sole maintainer**. Every line you write, you (or a future Claude session) will also be the one to read, debug, refactor, and extend. There is no teammate to ask, no Stack Overflow comment from a colleague, no "the original author would know." When you write code, write what a Claude session 6 months from now — with zero memory of this conversation — can pick up and continue from cold.

### Implications you must internalize

1. **Simplicity is the architecture, not a nice-to-have.** Every framework, abstraction, or service multiplies what a future Claude session must hold in context to make a safe change. The locked stack (below) is locked *because of this*, not because the user prefers it.
2. **Push back on over-engineering — even when the user asks for it.** If the user requests something that violates SPEC §0 maintainability rules (a microservice split, a new language, an ORM, a job queue, an auth system), say so before implementing. The constraint protects them from a codebase Claude can no longer evolve.
3. **Every change must be makeable by one Claude session.** If a feature would require more than ~500 LOC or ~5 files in one go, split it into phases. See CONTRACTS §6.
4. **Files are the only memory.** Conversation context vanishes between sessions. If a decision matters beyond this session, it goes in SPEC.md, CONTRACTS.md, code comments where genuinely non-obvious, or the user's auto-memory — never "I'll remember to do X later." There is no later-you that remembers.
5. **When in doubt, choose the option a cold Claude session can maintain.** Boring, well-known libraries (pandas, streamlit, duckdb) over clever ones. Inline code over premature abstraction. Explicit over implicit. Pinned versions over floating.

## First thing every session

**Read [CONTRACTS.md](CONTRACTS.md) before writing any code.** It's the pinned interface surface (~400 lines): module boundaries, function signatures, session-state keys, cross-cutting rules, phase scope. Every chunk implements *against* it. If something a chunk needs isn't in CONTRACTS, update CONTRACTS first, then code. **This is how the codebase stays coherent across sessions despite Claude having no memory of prior work.**

## When to load SPEC.md

[SPEC.md](SPEC.md) is 1,170 lines — load by section, not whole. Use it for:

| What you need | Section |
|---|---|
| Maintainability rules (push back on over-engineering) | §0 |
| Exact KPI formulas | §5.x and §8.2 |
| RAG thresholds and the `config.RAG` shape | §5.4 |
| Canonical CSV field names (what `mapping` targets) | §4.1 |
| DuckDB DDL | §8.1 |
| Validation rules | §3.1 |
| UI layout details (filter bar, chart sizing) | §10 |
| TKO brand integration / design tokens | §7.4, §10.3 |
| Domain glossary (S&OP, MAPE, DOS, LE, etc.) | §13 |

CONTRACTS already cross-references SPEC where relevant — follow those pointers rather than scanning the whole file.

## Hard architectural constraints

Stack is **locked** (Python 3.11 + Streamlit 1.36 + DuckDB 1.0 + pandas 2.2 + Plotly 5.22 + python-pptx 0.6.23 + openpyxl 3.1.4 + kaleido 0.2.1). Pinned in `requirements.txt`. **Do not** introduce:

- React, TypeScript, or any JS toolchain
- FastAPI or a separate backend service
- Celery, Redis, async/threading, background jobs
- Docker, Nginx, JWT, auth/RBAC
- ORMs, YAML/`.env` config, settings classes

These aren't preferences — they are load-bearing constraints. Each item on this list, if added, would push the codebase past what a single Claude session can safely modify. See SPEC §0 for all 10 maintainability rules and the *why* behind each.

**If a feature seems to require something on the "do not introduce" list, stop and discuss with the user first.** There is almost always a Streamlit/DuckDB/pandas way that's 90% as good and 10% of the maintenance surface.

## Working environment

- **OS:** Windows 11, default shell is PowerShell. Use PowerShell syntax (`$env:VAR`, `$null`, backtick for line continuation) when shelling out. Bash is also available via the Bash tool.
- **App URL:** `http://localhost:8501` (set in `.streamlit/config.toml`). The original SPEC pinned port 3000, but Streamlit 1.36's bundled frontend has port 8501 hardcoded — running on any other port leaves the WebSocket trying to reach 8501 and the app never finishes loading. 8501 is the canonical port for this stack.
- **Run:** `streamlit run app.py` from the project root
- **DuckDB file + raw CSVs:** under `storage/` (gitignored, created at first run)
- **TKO brand assets:** `TKO Design System.zip` at project root — unzip to `assets/tko/` (gitignored) as part of Phase 6

## Build approach

Seven phases, each scoped to one Claude session (~500 LOC max). See [CONTRACTS.md §6 Phase Contract Map](CONTRACTS.md) for what each phase introduces vs. consumes. At the end of every phase: app must still run end-to-end with sample CSVs, and CONTRACTS.md reflects any signature changes made.

Test data is generated by `generate_test_data.py` and lives in `test_data/`.

## Session hygiene

- One chunk = one session = ends with the app runnable
- Update CONTRACTS.md *first* when changing an interface
- End of session: commit, summarize what changed, `/clear`
- Start of session: read CONTRACTS.md, scan recent commits, then proceed
