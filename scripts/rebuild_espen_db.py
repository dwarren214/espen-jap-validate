#!/usr/bin/env python3
"""Regenerate espen_tables.sql and espen.db from Excel inventory.

Usage:
    uv run scripts/rebuild_espen_db.py
    uv run scripts/rebuild_espen_db.py --excel path/to/inventory.xlsx
"""

import argparse
import json
import sqlite3
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--excel",
        type=Path,
        default=Path(__file__).parent.parent / "ESPEN_DB_Inventory_Final.xlsx",
        help="Path to Excel inventory file (default: ESPEN_DB_Inventory_Final.xlsx)",
    )
    args = parser.parse_args()

    project_root = Path(__file__).parent.parent
    db_path = project_root / "espen.db"
    sql_path = project_root / "espen_tables.sql"

    if not args.excel.exists():
        print(f"Error: {args.excel} not found", file=sys.stderr)
        sys.exit(1)

    try:
        import pandas as pd
    except ImportError:
        print("Error: pandas not installed. Run: uv add --dev pandas openpyxl", file=sys.stderr)
        sys.exit(1)

    print(f"Reading {args.excel.name}...")
    xl = pd.ExcelFile(args.excel)

    # Read index sheet to get table names and their codes
    index_df = pd.read_excel(xl, sheet_name="Index")

    rows = []
    for _, row in index_df.iterrows():
        code = row["Code"]  # AT_1, AT_2, etc.
        table_name = row["Table"]

        if code not in xl.sheet_names:
            print(f"  Warning: sheet {code} not found for table {table_name}")
            continue

        # Read field definitions from corresponding sheet
        fields_df = pd.read_excel(xl, sheet_name=code)
        fields = [
            {"Name": str(r["FieldName"]), "Description": str(r["Description"])}
            for _, r in fields_df.iterrows()
            if pd.notna(r["FieldName"])
        ]

        rows.append((table_name, json.dumps(fields)))

    print(f"Found {len(rows)} tables")

    # Write SQL file
    with open(sql_path, "w") as f:
        for name, fields in rows:
            fields_escaped = fields.replace("'", "''")
            f.write(f"INSERT INTO espen_tables VALUES ('{name}', '{fields_escaped}');\n")
    print(f"Wrote {sql_path.name}")

    # Rebuild SQLite
    db_path.unlink(missing_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute("""
        CREATE TABLE espen_tables (
            Name_Analytical_Table TEXT PRIMARY KEY,
            Fields JSON
        )
    """)
    conn.executescript(sql_path.read_text())
    conn.commit()
    conn.close()

    print(f"Rebuilt {db_path.name} ({len(rows)} tables)")


if __name__ == "__main__":
    main()
