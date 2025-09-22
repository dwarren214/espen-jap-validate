import csv
import io
import logging
import os
import sqlite3
import sys
from pathlib import Path
from typing import List, Optional
import httpx
from datetime import datetime
from contextlib import contextmanager
import orjson
from dotenv import load_dotenv
from fastapi import Depends, FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from pydantic import BaseModel
from sqlalchemy import create_engine, text, Column, inspect, String, MetaData, Table, Date
from sqlalchemy.dialects.sqlite import Insert
from sqlalchemy.orm import sessionmaker

from utils import quote_identifiers  # Import the helper function

BASE_PATH = Path(__file__).resolve(strict=True).parent
load_dotenv()

logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

stream_handler = logging.StreamHandler(sys.stdout)
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)

app = FastAPI()

# Database configuration
DATABASE_URL = os.getenv("DATABASE_URL")

# Adjust for Heroku Postgres URL scheme
IS_POSTGRES = "postgres" in DATABASE_URL and "mssql" not in DATABASE_URL
if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

# Create SQLAlchemy engine and session
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)

meta_db_path = BASE_PATH / "espen.db"
meta_engine = create_engine("sqlite:///" + str(meta_db_path))
MetaSessionLocal = sessionmaker(bind=meta_engine)

# API Key configuration
API_KEY = os.getenv("API_KEY")  # Set this in your environment variables
if not API_KEY:
    raise ValueError("API_KEY environment variable is not set.")

ESPEN_CAMPAIGN_HUB_KEY = os.getenv("ESPEN_CAMPAIGN_HUB_KEY")
if not ESPEN_CAMPAIGN_HUB_KEY:
    logger.warning("ESPEN_CAMPAIGN_HUB_KEY environment variable is not set. Campaign data may not be available.")

ESPEN_CAMPAIGN_TABLE_NAME = "campaigns"

security = HTTPBearer()


# Dependency for API Key authentication
def api_key_auth(credentials: HTTPAuthorizationCredentials = Depends(security)):
    if credentials.scheme != "Bearer":
        raise HTTPException(status_code=403, detail="Invalid authentication scheme.")
    if credentials.credentials != API_KEY:
        raise HTTPException(status_code=403, detail="Invalid or missing API key.")


# Constants for response format thresholds
MAX_JSON_ROWS = 100  # Maximum number of rows for JSON response
MAX_JSON_CELLS = 1000  # Maximum total cells (rows * columns) for JSON response


# Pydantic models for request bodies
class SQLQuery(BaseModel):
    query: str
    force_json: Optional[bool] = False  # Optional parameter to force JSON response
    force_csv: Optional[bool] = False  # Optional parameter to force CSV response


class TableNames(BaseModel):
    tables: List[str]


@app.get("/up")
def status():
    return {"status": "up"}


# Endpoint to execute SQL queries
@app.post("/execute_sql_query", dependencies=[Depends(api_key_auth)])
def execute_sql_query(query_data: SQLQuery):
    session = SessionLocal()
    try:
        # Automatically quote identifiers in the query
        quoted_query = quote_identifiers(query_data.query, IS_POSTGRES)

        # Execute the quoted query with parameters
        result = session.execute(text(quoted_query))
        rows = result.fetchall()
        column_names = list(result.keys())

        # Calculate result size
        num_rows = len(rows)
        num_columns = len(column_names)
        total_cells = num_rows * num_columns

        # Determine if we should return CSV based on size thresholds
        should_return_csv = query_data.force_csv or (
            not query_data.force_json  # Not forcing JSON
            and (num_rows > MAX_JSON_ROWS or total_cells > MAX_JSON_CELLS)
        )

        if should_return_csv:
            # Create a string buffer for CSV data
            output = io.StringIO()
            writer = csv.writer(output)

            # Write header
            writer.writerow(column_names)

            # Write data rows
            for row in rows:
                writer.writerow(row)

            # Prepare the response
            output.seek(0)
            return StreamingResponse(
                iter([output.getvalue()]),
                media_type="text/csv",
                headers={
                    "Content-Disposition": "attachment; filename=query_results.csv"
                },
            )
        else:
            # Convert to list of dictionaries for JSON response
            result_list = [dict(zip(column_names, row)) for row in rows]
            return result_list

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        session.close()


# Endpoint to fetch column names and types
@app.post("/fetch_column_names_and_types", dependencies=[Depends(api_key_auth)])
def fetch_column_names_and_types(table_data: TableNames):
    meta_session = MetaSessionLocal()
    results = {}
    try:
        for table_name in table_data.tables:
            # Query to get fields and descriptions
            query = text("""
                SELECT "Fields"
                FROM espen_tables
                WHERE LOWER("Name_Analytical_Table") = :table_name
            """)
            res = meta_session.execute(
                query, {"table_name": table_name.lower()}
            ).fetchone()
            if res:
                fields = orjson.loads(res[0])
                results[table_name] = {
                    field["Name"]: field["Description"] for field in fields
                }
            else:
                results[table_name] = "Table not found"
        return results
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        meta_session.close()

@contextmanager
def campaign_hub_session():
    """Provides a transactional scope around a series of operations."""
    engine = create_engine('sqlite:///campaigns.db')
    session = sessionmaker(bind=engine)()

    try:
        yield session
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()

@app.post("/query_campaign_hub_data", dependencies=[Depends(api_key_auth)])
def query_campaign_hub_data(query_data: SQLQuery):
    """
    Fetches campaign data from the ESPEN Campaign Hub API and stores it in a local SQLite database.
    If the local database table does not exist or is outdated, it fetches fresh data from the API.
    Then, it executes the user's SQL query against the local database and returns the results.

    The API data is updated every night, so we check if the local data was updated today before deciding to fetch new
    data.
    """
    engine = create_engine('sqlite:///campaigns.db')
    inspector = inspect(engine)
    table_is_outdated = False
    if not inspector.has_table(ESPEN_CAMPAIGN_TABLE_NAME):
        logger.info(f"{ESPEN_CAMPAIGN_TABLE_NAME} table does not exist. Fetching data from API and creating table.")
        table_is_outdated = True
    else:
        with campaign_hub_session() as session:
            result = session.execute(text(f"SELECT updated_at FROM {ESPEN_CAMPAIGN_TABLE_NAME} LIMIT 1"))
            row = result.fetchone()

        if not row or row[0] != str(datetime.now().date()):
            logger.info(f"{ESPEN_CAMPAIGN_TABLE_NAME} table is empty or not updated today. Fetching data from API.")
            table_is_outdated = True

    if table_is_outdated:
        data = _fetch_campaign_data()
        sync_campaign_hub_db(data)
        
    with campaign_hub_session() as session:
        # Execute the user's query
        result = session.execute(text(query_data.query))
        rows = result.fetchall()

    # Convert results to a list of dictionaries for JSON response
    columns = result.keys()
    result_data = [dict(zip(columns, row)) for row in rows]
    
    return {"data": result_data, "row_count": len(result_data)}

@app.get("/fetch_campaign_hub_columns", dependencies=[Depends(api_key_auth)])
def fetch_column_names_and_types():
    results = {}
    with campaign_hub_session() as session:
        try:
            query = text(f"PRAGMA table_info({ESPEN_CAMPAIGN_TABLE_NAME});")
            results = session.execute(query).fetchall()
            columns_names = [res[1] for res in results]
            return columns_names
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))



def sync_campaign_hub_db(data):
    """
    Ensures the campaigns table exists and is upserted with the latest data.
    Detects schema changes and adds/removes columns as needed.
    
    Params:
    data - the latest set of records fetched from the API. This is the source of truth from which the table is built.
    """
    engine = create_engine('sqlite:///campaigns.db')
    inspector = inspect(engine)
    metadata = MetaData()

    # Get columns from incoming data
    single_record = data[0]
    
    new_column_names = [_key_to_column_name(key) for key in single_record.keys()]
    # Always include updated_at column
    new_column_names.append("updated_at")
    new_column_names_set = set(new_column_names)

    # Check if table exists and get current columns
    table_exists = inspector.has_table(ESPEN_CAMPAIGN_TABLE_NAME)
    
    if table_exists:
        # Get existing columns
        existing_columns = inspector.get_columns(ESPEN_CAMPAIGN_TABLE_NAME)
        existing_column_names = {col['name'] for col in existing_columns}
        
        # Detect column differences
        columns_to_add = new_column_names_set - existing_column_names
        columns_to_remove = existing_column_names - new_column_names_set

        # Add new columns
        if columns_to_add:
            logger.info(f"Adding new columns to {ESPEN_CAMPAIGN_TABLE_NAME}: {columns_to_add}")
            with engine.begin() as conn:
                for column_name in columns_to_add:
                    if column_name == "updated_at":
                        conn.execute(text(f"ALTER TABLE {ESPEN_CAMPAIGN_TABLE_NAME} ADD COLUMN {column_name} DATE"))
                    else:
                        conn.execute(text(f"ALTER TABLE {ESPEN_CAMPAIGN_TABLE_NAME} ADD COLUMN {column_name} TEXT"))
        
        # Remove obsolete columns
        if columns_to_remove:
            logger.info(f"Removing obsolete columns from {ESPEN_CAMPAIGN_TABLE_NAME}: {columns_to_remove}")
            _drop_columns_from_table(engine, columns_to_remove)
    else:
        # Create table for the first time
        logger.info(f"Creating new table {ESPEN_CAMPAIGN_TABLE_NAME}")
        columns = [Column(column_name, String, primary_key=True if column_name == "campaign_id" else False) 
                  for column_name in new_column_names if column_name != "updated_at"]
        # Add updated_at column
        columns.append(Column("updated_at", Date))
        table = Table(ESPEN_CAMPAIGN_TABLE_NAME, metadata, *columns)
        metadata.create_all(engine)

    # Recreate the table object with current schema for upsert operations
    metadata = MetaData()
    columns = [Column(column_name, String, primary_key=True if column_name == "campaign_id" else False) 
              for column_name in new_column_names if column_name != "updated_at"]
    columns.append(Column("updated_at", Date))
    table = Table(ESPEN_CAMPAIGN_TABLE_NAME, metadata, *columns, extend_existing=True)

    # Upsert the table with data
    campaigns_seen = []
        
    with campaign_hub_session() as session:
        with session.begin():
            for record in data:
                # Convert keys to match column names
                record = { _key_to_column_name(key): str(value) if value is not None else None for key, value in record.items()}
                campaigns_seen.append(record["campaign_id"])
                record["updated_at"] = datetime.now().date()
                insert_stmt = Insert(table).values(**record)
                upsert_stmt = insert_stmt.on_conflict_do_update(
                    index_elements=['campaign_id'],
                    set_=record
                )
                session.execute(upsert_stmt)

            # Delete records not in the latest fetch
            if campaigns_seen:
                delete_stmt = table.delete().where(~table.c.campaign_id.in_(campaigns_seen))
                session.execute(delete_stmt)
            
            session.commit()
        

def _key_to_column_name(key: str) -> str:
    return key.lower().replace(" ", "_").replace("-", "_").replace("(", "").replace(")", "")


def _drop_columns_from_table(engine, columns_to_remove):
    """
    Drops the specified columns from the table using ALTER TABLE DROP COLUMN.
    """
    with engine.begin() as conn:
        for column_name in columns_to_remove:
            drop_sql = f'ALTER TABLE {ESPEN_CAMPAIGN_TABLE_NAME} DROP COLUMN "{column_name}"'
            conn.execute(text(drop_sql))
            logger.info(f"Successfully dropped column '{column_name}' from {ESPEN_CAMPAIGN_TABLE_NAME}")

        

def _fetch_campaign_data():
    previous_year = datetime.now().year - 1
    headers = {"access_token": ESPEN_CAMPAIGN_HUB_KEY}
    response = httpx.get(url="https://lbdatabaseapi.azurewebsites.net/campaign_hub_download", headers=headers)

    if response.status_code == 200:
        records = response.json()
        final_response = []
        for record in records:
            # record.get("Campaign Start Year", 0) or 0) ensures we handle None values
            if record.get("WHO Region") == "AFRO" and (record.get("Campaign Start Year", 0) or 0) > previous_year:
                cleaned_data = record | {"Diseases Targeted": record.get("Diseases Targeted", "unspecified")}
                cleaned_data.pop("PCCS Coverage", None)
                cleaned_data.pop("Geographic Coverage", None)
                cleaned_data.pop("Therapeutic Coverage", None)
                cleaned_data.pop("Administrative Coverage", None)
                final_response.append(cleaned_data)

        return final_response
    else:
        raise HTTPException(status_code=response.status_code, detail=response.text)
