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
def campaign_hub_session(date=None):
    """Provides a transactional scope around a series of operations."""
    if date is None:
        date = datetime.now().date()
    db_path = get_campaign_db_path(date)
    engine = create_engine(f'sqlite:///{db_path}')
    session = sessionmaker(bind=engine)()

    try:
        yield session
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()

def get_campaign_db_path(date):
    if date is None:
        date = datetime.now().date()
    db_name = f'campaigns-{date}.db'
    return BASE_PATH / "campaign_dbs" / db_name

@app.post("/query_campaign_hub_data", dependencies=[Depends(api_key_auth)])
def query_campaign_hub_data(query_data: SQLQuery):
    """
    Fetches campaign data from the ESPEN Campaign Hub API and stores it in a local SQLite database.
    Uses date-specific database files (e.g., campaigns-2025-09-23.db) to ensure fresh data each day.
    If today's database doesn't exist, it fetches fresh data from the API and creates a new database.
    Then, it executes the user's SQL query against the database and returns the results.

    The API data is updated every night, so we create a new database file each day.
    """

    today = datetime.now().date()
    ensure_campaigns_db_exists()
    # Execute the user's query
    with campaign_hub_session(today) as session:
        result = session.execute(text(query_data.query))
        rows = result.fetchall()

    columns = result.keys()
    result_data = [dict(zip(columns, row)) for row in rows]
    
    return {"data": result_data, "row_count": len(result_data)}

@app.get("/fetch_campaign_hub_columns", dependencies=[Depends(api_key_auth)])
def fetch_campaign_hub_columns():
    ensure_campaigns_db_exists()

    today = datetime.now().date()
    results = {}
    with campaign_hub_session(today) as session:
        try:
            query = text(f"PRAGMA table_info({ESPEN_CAMPAIGN_TABLE_NAME});")
            results = session.execute(query).fetchall()
            columns_names = [res[1] for res in results]
            return columns_names
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))

def ensure_campaigns_db_exists():
    today = datetime.now().date()
    db_path = get_campaign_db_path(today)

    if not db_path.exists():
        # Create today's database if it doesn't exist
        logger.info(f"Database for {today} does not exist. Fetching data from API and creating new database.")
        data = fetch_campaign_data()
        
        # Check again in case another process created it in the meantime
        if not db_path.exists():
            create_db_from_data(db_path, data)
            # Piggy back on this call to do some cleanup
            remove_old_dbs()

def create_db_from_data(db_path, data):
    """
    Creates a new campaigns table with the latest data from the API.
    Uses a date-specific database file for the given date.
    
    Params:
    data - the latest set of records fetched from the API. This is the source of truth from which the table is built.
    date - the date for which to create the database (datetime.date object)
    """
    engine = create_engine(f'sqlite:///{db_path}')
    metadata = MetaData()

    # Look at the first record to see what columns we have
    single_record = data[0]
    column_names = [key_to_column_name(key) for key in single_record.keys()]

    # Create table
    logger.info(f"Creating new table {ESPEN_CAMPAIGN_TABLE_NAME} in {db_path}")
    columns = [Column(column_name, String, primary_key=True if column_name == "campaign_id" else False) 
              for column_name in column_names]

    table = Table(ESPEN_CAMPAIGN_TABLE_NAME, metadata, *columns)
    metadata.create_all(engine)

    # Insert the data
    today = datetime.now().date()
    with campaign_hub_session(today) as session:
        with session.begin():
            for record in data:
                record = { key_to_column_name(key): str(value) if value is not None else None for key, value in record.items()}
                insert_stmt = Insert(table).values(**record)
                session.execute(insert_stmt)
            
            session.commit()
        

def key_to_column_name(key: str) -> str:
    return key.lower().replace(" ", "_").replace("-", "_").replace("(", "").replace(")", "")


def fetch_campaign_data():
    previous_year = datetime.now().year - 1
    headers = {"access_token": ESPEN_CAMPAIGN_HUB_KEY}
    response = httpx.get(url="https://lbdatabaseapi.azurewebsites.net/campaign_hub_download", headers=headers)

    if response.status_code == 200:
        records = response.json()
        final_response = []
        for record in records:
            
            # Pre-filtering to only insert data we care about
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

def remove_old_dbs():
    """Removes old campaign database files, keeping only the last 2 days."""
    db_dir = BASE_PATH / "campaign_dbs"
    if not db_dir.exists():
        return

    today = datetime.now().date()
    for db_file in db_dir.glob("campaigns-*.db"):
        try:
            # campaigns-2025-09-23.db -> ['2025', '09', '23']
            date_str = db_file.stem.split("-")[1:]
            db_date = datetime.strptime("-".join(date_str), "%Y-%m-%d").date()
            if (today - db_date).days > 2:
                logger.info(f"Removing old database file: {db_file}")
                db_file.unlink()
        except Exception as e:
            logger.exception("Unable to remove db file %s: %s", db_file, e)


# This endpoint is a legacy one and will be removed once the OCS bot uses the endpoints above
@app.get("/fetch_campaign_hub_data", dependencies=[Depends(api_key_auth)])
def fetch_campaign_hub_data():
    """
    Fetches campaign data from the ESPEN Campaign Hub API. The data is filtered to include only records
    from the AFRO region with a campaign start year greater than last year.
    """
    previous_year = datetime.now().year - 1
    headers = {"access_token": os.getenv("ESPEN_CAMPAIGN_HUB_KEY")}
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