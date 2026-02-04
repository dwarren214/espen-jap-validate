# HTTP SQL API

A small API for querying the database.

This is used in conjunction with Open Chat Studio [custom actions](https://dimagi.github.io/open-chat-studio-docs/concepts/custom_actions/) to allow a bot to query the DB via an
HTTP API.

## Prerequisites

This project connects to multiple databases:
- **PostgreSQL** – oncho projection data
- **MSSQL** (Azure SQL) – ESPEN analytical tables
- **SQLite** – local metadata cache (`espen.db`)

## Setup

1. Install UV

   See https://docs.astral.sh/uv/getting-started/installation/

2. Install the dependencies

   ```shell
   uv sync   
   ```

3. Create a `.env` file by copying `.env.example` and updating the values

4. Run the app

    ```shell
    uv run uvicorn espen_sql_api.main:app --reload
    ```

5. Test the API

   ```
   export API_KEY="xxx"
   curl -X POST localhost:8000/fetch_column_names_and_types -H "Authorization: Bearer $API_KEY" -d'{"tables": ["Afro_Admin0"]}' -H "Content-type: application/json"
   ```

## Integration Testing

There are a basic set of tests which can be run against the production endpoint to verify it is functioning correctly:

```bash
uv run pytest -v
```

## Maintaining espen.db

The `espen.db` SQLite database contains metadata (table names and column descriptions) for the ESPEN analytical tables. This metadata is used by chatbots to understand the schema before constructing SQL queries.

### Updating the metadata

1. Obtain the latest `ESPEN_DB_Inventory_Final.xlsx` from your ESPEN contact
2. Place it in the project root
3. Run the rebuild script:
   ```bash
   uv run scripts/rebuild_espen_db.py
   ```
4. Commit the updated `espen.db` and `espen_tables.sql`

### Comparing with the live database

To verify the metadata matches the actual MSSQL schema, run inside the deployed container:

```bash
cd deploy && kamal app exec 'python scripts/compare_db_schema.py'
```

This will show:
- Tables in espen.db but not in MSSQL
- Tables in MSSQL but missing from espen.db
- Column differences for common tables
