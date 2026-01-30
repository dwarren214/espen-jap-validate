# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

FastAPI-based HTTP SQL API for ESPEN data. Allows chatbots (via Open Chat Studio custom actions) to query:
- **Remote MSSQL** (ESPEN analytical tables) via `/execute_sql_query`
- **Local PostgreSQL** (oncho projections) via `/oncho/execute_query`, `/oncho/top3`
- **SQLite** (metadata `espen.db`, daily campaign caches `campaign_dbs/`)
- **Campaign Hub API** (fetched daily, cached to SQLite)

## Commands

```bash
# Install dependencies
uv sync

# Run dev server
uv run uvicorn espen_sql_api.main:app --reload

# Lint
uv run ruff check .
uv run ruff check --fix .

# Test
uv run pytest tests/ -v

# Deploy (from deploy/ dir, requires Kamal + 1Password CLI + AWS CLI)
cd deploy && kamal deploy
```

Basic integration tests in `tests/` (SQLite endpoints tested locally, DB endpoints skipped without real connections).

## Architecture

**Src layout** (`src/espen_sql_api/`):
- `main.py` – FastAPI app, health endpoints, router includes
- `db.py` – database engines and session factories
- `auth.py` – API key authentication
- `config.py` – guardrail thresholds, app configuration
- `schemas.py` – Pydantic request models (SQLQuery, TableNames)
- `utils.py` – `quote_identifiers()`, `execute_query_with_retry()`, `validate_query_safety()`
- `response_guard.py` – `GuardrailThresholds`, `estimate_result_size()`, `build_guardrail_payload()`
- `routers/oncho.py` – PostgreSQL oncho projection endpoints
- `routers/espen.py` – MSSQL analytical table + SQLite metadata endpoints
- `routers/campaign.py` – Campaign Hub API + SQLite cache endpoints

**Three DB engines** (in `db.py`):
- `engine` (PostgreSQL) – `DATABASE_URL` – oncho projection data
- `remote_engine` (MSSQL via pyodbc) – `REMOTE_DATABASE_URL` – ESPEN analytical tables
- `meta_engine` (SQLite) – `espen.db` – table metadata lookup

**Query safety:** `validate_query_safety()` uses sqlglot AST to block unbounded `SELECT *` queries.

**Retry logic:** `execute_query_with_retry()` uses tenacity for transient MSSQL connection errors (08S01, etc).

**Campaign data:** Daily SQLite DB created from external API; old DBs auto-purged after 2 days.

**Metadata maintenance:** `espen.db` is rebuilt from `ESPEN_DB_Inventory_Final.xlsx` (obtained from ESPEN contact).

## Scripts

- `scripts/rebuild_espen_db.py` – Rebuild `espen.db` from Excel inventory
- `scripts/compare_db_schema.py --docker` – Compare local metadata with live MSSQL schema

## Environment Variables

See `.env.example`. Key vars:
- `DATABASE_URL`, `REMOTE_DATABASE_URL` – connection strings
- `API_KEY` – bearer auth for all endpoints
- `ESPEN_CAMPAIGN_HUB_KEY` – external API access
- `OPENAI_API_KEY` – fallback SQL transpilation
- `MAX_RETURNED_ROWS`, `MAX_RETURNED_BYTES`, `PREVIEW_ROW_COUNT` – guardrail limits
- `DB_POOL_*`, `DB_QUERY_TIMEOUT` – connection pool tuning

## API Endpoints

| Endpoint | Method | Database | Purpose |
|----------|--------|----------|---------|
| `/execute_sql_query` | POST | MSSQL | Run arbitrary SQL on ESPEN tables |
| `/fetch_column_names_and_types` | POST | SQLite | Get table schema from metadata |
| `/oncho/execute_query` | POST | PostgreSQL | Run SQL on oncho_projection |
| `/oncho/top3` | POST | PostgreSQL | Ranked scenario analysis |
| `/query_campaign_hub_data` | POST | SQLite | SQL on cached campaign data |
| `/fetch_campaign_hub_columns` | GET | SQLite | Campaign table columns |
| `/health` | GET | All | Pool status, connectivity |

All data endpoints require `Authorization: Bearer $API_KEY`.
