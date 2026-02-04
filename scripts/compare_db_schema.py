#!/usr/bin/env python3
"""Compare espen.db metadata with actual MSSQL schema.

Usage:
    REMOTE_DATABASE_URL="mssql+pyodbc://..." ./scripts/compare_db_schema.py
"""

import json
import os
import sqlite3
import sys
from pathlib import Path


def compare():
    """Fetch from MSSQL and compare with local SQLite."""
    from sqlalchemy import create_engine, text

    remote_url = os.environ.get("REMOTE_DATABASE_URL")
    if not remote_url:
        print("Error: REMOTE_DATABASE_URL not set", file=sys.stderr)
        sys.exit(1)

    # Load local SQLite metadata
    db_path = Path(__file__).parent.parent / "espen.db"
    conn = sqlite3.connect(db_path)
    local_tables = {}
    for row in conn.execute("SELECT Name_Analytical_Table, Fields FROM espen_tables"):
        table_name, fields_json = row
        fields = json.loads(fields_json)
        local_tables[table_name] = {f["Name"] for f in fields}
    conn.close()

    # Fetch from MSSQL
    print("Fetching schema from MSSQL...")
    engine = create_engine(remote_url)

    remote_tables = {}
    with engine.connect() as conn:
        tables = conn.execute(text("""
            SELECT TABLE_NAME
            FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_TYPE = 'BASE TABLE'
              AND TABLE_SCHEMA = 'dbo'
            ORDER BY TABLE_NAME
        """)).fetchall()

        for (table_name,) in tables:
            columns = conn.execute(text("""
                SELECT COLUMN_NAME
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = :table_name
                  AND TABLE_SCHEMA = 'dbo'
                ORDER BY ORDINAL_POSITION
            """), {"table_name": table_name}).fetchall()
            remote_tables[table_name] = {col for (col,) in columns}

    # Compare
    local_names = set(local_tables.keys())
    remote_names = set(remote_tables.keys())

    only_local = local_names - remote_names
    only_remote = remote_names - local_names
    common = local_names & remote_names

    print(f"\nLocal (espen.db): {len(local_names)} tables")
    print(f"Remote (MSSQL):   {len(remote_names)} tables")
    print(f"Common:           {len(common)} tables")

    if only_local:
        print(f"\n--- Tables only in espen.db ({len(only_local)}) ---")
        for t in sorted(only_local):
            print(f"  {t}")

    if only_remote:
        print(f"\n--- Tables only in MSSQL ({len(only_remote)}) ---")
        for t in sorted(only_remote):
            print(f"  {t}")

    # Column differences for common tables
    tables_with_diff = []
    for table in sorted(common):
        local_cols = local_tables[table]
        remote_cols = remote_tables[table]
        only_local_cols = local_cols - remote_cols
        only_remote_cols = remote_cols - local_cols

        if only_local_cols or only_remote_cols:
            tables_with_diff.append((table, only_local_cols, only_remote_cols))

    if tables_with_diff:
        print(f"\n--- Column differences ({len(tables_with_diff)} tables) ---")
        for table, only_local_cols, only_remote_cols in tables_with_diff:
            print(f"\n  {table}:")
            if only_local_cols:
                print(f"    Only in espen.db: {', '.join(sorted(only_local_cols))}")
            if only_remote_cols:
                print(f"    Only in MSSQL:    {', '.join(sorted(only_remote_cols))}")
    else:
        print("\n--- No column differences in common tables ---")


if __name__ == "__main__":
    compare()
