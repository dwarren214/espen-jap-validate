import csv
import io
import logging
import os
import sys
from pathlib import Path
from typing import List, Literal, Optional
from datetime import datetime
from contextlib import contextmanager
import orjson
import httpx
from dotenv import load_dotenv
from fastapi import Depends, FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from pydantic import BaseModel, field_validator, FieldValidationInfo
from sqlalchemy import (
    bindparam,
    create_engine,
    text,
    Column,
    inspect,
    String,
    Integer,
    Float,
    MetaData,
    Table,
)
from sqlalchemy.dialects.sqlite import Insert
from sqlalchemy.orm import Session, sessionmaker

from utils import quote_identifiers  # Import the helper function
from response_guard import (
    GuardrailThresholds,
    estimate_result_size,
    build_guardrail_payload,
)

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
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)

REMOTE_DATABASE_URL = os.getenv("REMOTE_DATABASE_URL")
remote_engine = create_engine(REMOTE_DATABASE_URL)
RemoteSessionLocal = sessionmaker(bind=remote_engine)

meta_db_path = BASE_PATH / "espen.db"
meta_engine = create_engine("sqlite:///" + str(meta_db_path))
MetaSessionLocal = sessionmaker(bind=meta_engine)

# Guardrail configuration
MAX_RETURNED_ROWS = int(os.getenv("MAX_RETURNED_ROWS", "5000"))
MAX_RETURNED_BYTES = int(os.getenv("MAX_RETURNED_BYTES", "400000"))
PREVIEW_ROW_COUNT = int(os.getenv("PREVIEW_ROW_COUNT", "50"))

guardrail_thresholds = GuardrailThresholds(
    max_rows=MAX_RETURNED_ROWS,
    max_bytes=MAX_RETURNED_BYTES,
    preview_rows=PREVIEW_ROW_COUNT,
)

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
    session = RemoteSessionLocal()
    try:
        # Automatically quote identifiers in the query
        quoted_query = quote_identifiers(query_data.query, is_postgres=False)

        # Execute the quoted query with parameters
        result = session.execute(text(quoted_query))
        rows = result.fetchall()
        column_names = list(result.keys())

        stats = estimate_result_size(rows, guardrail_thresholds)
        logger.info(
            "sql_query_stats|endpoint=execute_sql_query|rows=%s|cols=%s|bytes=%s|over_cap=%s",
            stats.row_count,
            stats.column_count,
            stats.bytes_estimate,
            stats.requires_guardrail,
        )

        if stats.requires_guardrail:
            return build_guardrail_payload(
                rows,
                column_names,
                stats,
                guardrail_thresholds,
            )

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
    try:
        # Execute the user's query
        with campaign_hub_session(today) as session:
            result = session.execute(text(query_data.query))
            rows = result.fetchall()

        columns = result.keys()
        result_data = [dict(zip(columns, row)) for row in rows]

        stats = estimate_result_size(rows, guardrail_thresholds)
        logger.info(
            "sql_query_stats|endpoint=query_campaign_hub_data|rows=%s|cols=%s|bytes=%s|over_cap=%s",
            stats.row_count,
            len(columns),
            stats.bytes_estimate,
            stats.requires_guardrail,
        )

        if stats.requires_guardrail:
            return build_guardrail_payload(
                rows,
                columns,
                stats,
                guardrail_thresholds,
            )

        return {"data": result_data, "row_count": len(result_data)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

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
    column_info = [get_column_name_and_type(key, value) for key, value in single_record.items()]

    # Create table
    logger.info(f"Creating new table {ESPEN_CAMPAIGN_TABLE_NAME} in {db_path}")
    columns = [Column(column_name, column_type, primary_key=True if column_name == "campaign_id" else False) 
              for column_name, column_type in column_info]

    table = Table(ESPEN_CAMPAIGN_TABLE_NAME, metadata, *columns)
    metadata.create_all(engine)

    # Insert the data
    today = datetime.now().date()
    with campaign_hub_session(today) as session:
        with session.begin():
            for record in data:
                record = { key_to_column_name(key): value if value is not None else None for key, value in record.items()}
                insert_stmt = Insert(table).values(**record)
                session.execute(insert_stmt)
            
            session.commit()
        

def key_to_column_name(key: str) -> str:
    return key.lower().replace(" ", "_").replace("-", "_").replace("(", "").replace(")", "")

def get_column_name_and_type(key: str, value: any) -> tuple[str, String | Integer | Float]:
    column_type = String
    if isinstance(value, int):
        column_type = Integer
    elif isinstance(value, float):
        column_type = Float
    return key_to_column_name(key), column_type


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


class Top3Request(BaseModel):
    location_type: Literal["country", "iu"]
    location_values: List[str]
    start_year: int
    end_year: int
    threshold: float = 0.01

    @field_validator("location_values")
    @classmethod
    def validate_location_values(cls, value: List[str]) -> List[str]:
        if not value:
            raise ValueError("location_values must contain at least one value.")
        return value

    @field_validator("end_year")
    @classmethod
    def validate_year_range(cls, end_year: int, info: FieldValidationInfo) -> int:
        start_year = info.data.get("start_year") if info.data else None
        if start_year is not None and end_year < start_year:
            raise ValueError("end_year must be greater than or equal to start_year.")
        return end_year

    @field_validator("threshold")
    @classmethod
    def validate_threshold(cls, threshold: float) -> float:
        if threshold <= 0:
            raise ValueError("threshold must be a positive number.")
        return threshold


class ScenarioRank(BaseModel):
    scenario_label: str
    success_rate_pct: Optional[float]
    median_years: Optional[float]
    total_cost: float
    min_remaining_p50: Optional[float]
    max_remaining_p50: Optional[float]


LOCATION_COLUMN_MAP = {
    "country": "country_name",
    "iu": "iu_name",
}


def build_top3_query(location_column: str):
    query = text(
        f"""
        WITH filtered AS (
            SELECT
                scenario_label,
                country_name,
                iu_id,
                iu_name,
                year,
                percentile_50 AS p50,
                fitz_cost AS cost
            FROM oncho_projection
            WHERE year BETWEEN :start_year AND :end_year
              AND {location_column} IN :location_values
        ),
        start_rows AS (
            SELECT *
            FROM filtered
            WHERE year = :start_year
        ),
        start_status AS (
            SELECT
                scenario_label,
                iu_id,
                iu_name,
                p50 AS start_p50
            FROM start_rows
        ),
        total_counts AS (
            SELECT scenario_label, COUNT(*) AS total_iu_count
            FROM start_status
            GROUP BY scenario_label
        ),
        start_success_counts AS (
            SELECT scenario_label, COUNT(*) AS initial_success_count
            FROM start_status
            WHERE start_p50 <= :threshold
            GROUP BY scenario_label
        ),
        eligible_units AS (
            SELECT scenario_label, iu_id
            FROM start_status
            WHERE start_p50 > :threshold
        ),
        eligible AS (
            SELECT f.*
            FROM filtered f
            JOIN eligible_units eu
              ON f.scenario_label = eu.scenario_label
             AND f.iu_id = eu.iu_id
        ),
        scenarios AS (
            SELECT scenario_label
            FROM total_counts
        ),
        success_status AS (
            SELECT
                e.scenario_label,
                e.iu_id,
                MIN(CASE WHEN e.p50 <= :threshold THEN e.year END) AS success_year
            FROM eligible e
            GROUP BY e.scenario_label, e.iu_id
        ),
        success_years AS (
            SELECT scenario_label, iu_id, success_year
            FROM success_status
            WHERE success_year IS NOT NULL
        ),
        success_counts AS (
            SELECT scenario_label, COUNT(*) AS successful_iu_count
            FROM success_years
            GROUP BY scenario_label
        ),
        success_durations AS (
            SELECT
                sy.scenario_label,
                sy.iu_id,
                (sy.success_year - :start_year + 1) AS duration
            FROM success_years sy
        ),
        duration_ranks AS (
            SELECT
                scenario_label,
                duration,
                ROW_NUMBER() OVER (PARTITION BY scenario_label ORDER BY duration) AS rn,
                COUNT(*) OVER (PARTITION BY scenario_label) AS cnt
            FROM success_durations
        ),
        median_durations AS (
            SELECT
                scenario_label,
                AVG(duration) AS median_years
            FROM duration_ranks
            WHERE rn IN (
                CAST(((cnt + 1) / 2.0) AS INTEGER),
                CAST(((cnt + 2) / 2.0) AS INTEGER)
            )
            GROUP BY scenario_label
        ),
        cost_rows AS (
            SELECT
                e.scenario_label,
                e.iu_id,
                e.cost,
                ss.success_year
            FROM eligible e
            LEFT JOIN success_status ss
              ON e.scenario_label = ss.scenario_label
             AND e.iu_id = ss.iu_id
            WHERE e.year BETWEEN :start_year AND COALESCE(ss.success_year, :end_year)
        ),
        iu_costs AS (
            SELECT
                scenario_label,
                iu_id,
                SUM(cost) AS iu_cost
            FROM cost_rows
            GROUP BY scenario_label, iu_id
        ),
        scenario_costs AS (
            SELECT
                scenario_label,
                SUM(iu_cost) AS total_cost
            FROM iu_costs
            GROUP BY scenario_label
        ),
        end_year_vals AS (
            SELECT
                e.scenario_label,
                e.iu_id,
                e.p50
            FROM eligible e
            WHERE e.year = :end_year
        ),
        unsuccessful_end_vals AS (
            SELECT
                ev.scenario_label,
                ev.p50
            FROM end_year_vals ev
            JOIN success_status ss
              ON ev.scenario_label = ss.scenario_label
             AND ev.iu_id = ss.iu_id
            WHERE ss.success_year IS NULL
              AND ev.p50 > :threshold
        ),
        minmax_remaining AS (
            SELECT
                scenario_label,
                MIN(p50) AS min_remaining_p50,
                MAX(p50) AS max_remaining_p50
            FROM unsuccessful_end_vals
            GROUP BY scenario_label
        ),
        final_metrics AS (
            SELECT
                sc.scenario_label,
                CASE
                    WHEN tc.total_iu_count > 0 THEN ROUND(
                        100.0 * (
                            COALESCE(ssc.initial_success_count, 0) +
                            COALESCE(succ.successful_iu_count, 0)
                        ) / tc.total_iu_count,
                        2
                    )
                    ELSE NULL
                END AS success_rate_pct,
                md.median_years,
                COALESCE(costs.total_cost, 0.0) AS total_cost,
                mm.min_remaining_p50,
                mm.max_remaining_p50
            FROM scenarios sc
            LEFT JOIN total_counts tc ON sc.scenario_label = tc.scenario_label
            LEFT JOIN start_success_counts ssc ON sc.scenario_label = ssc.scenario_label
            LEFT JOIN success_counts succ ON sc.scenario_label = succ.scenario_label
            LEFT JOIN median_durations md ON sc.scenario_label = md.scenario_label
            LEFT JOIN scenario_costs costs ON sc.scenario_label = costs.scenario_label
            LEFT JOIN minmax_remaining mm ON sc.scenario_label = mm.scenario_label
        )
        SELECT
            scenario_label,
            success_rate_pct,
            median_years,
            total_cost,
            min_remaining_p50,
            max_remaining_p50
        FROM final_metrics
        ORDER BY
            (success_rate_pct IS NOT NULL) DESC,
            success_rate_pct DESC,
            (median_years IS NULL),
            median_years,
            total_cost
        LIMIT 3
        """
    )
    return query.bindparams(bindparam("location_values", expanding=True))


def run_top3_query(session: Session, payload: Top3Request) -> List[ScenarioRank]:
    location_column = LOCATION_COLUMN_MAP[payload.location_type]
    statement = build_top3_query(location_column)
    params = {
        "start_year": payload.start_year,
        "end_year": payload.end_year,
        "threshold": payload.threshold,
        "location_values": payload.location_values,
    }
    result = session.execute(statement, params)
    rows = result.mappings().all()
    scenarios: List[ScenarioRank] = []
    for row in rows:
        scenarios.append(
            ScenarioRank(
                scenario_label=row["scenario_label"],
                success_rate_pct=float(row["success_rate_pct"]) if row["success_rate_pct"] is not None else None,
                median_years=float(row["median_years"]) if row["median_years"] is not None else None,
                total_cost=float(row["total_cost"]) if row["total_cost"] is not None else 0.0,
                min_remaining_p50=float(row["min_remaining_p50"]) if row["min_remaining_p50"] is not None else None,
                max_remaining_p50=float(row["max_remaining_p50"]) if row["max_remaining_p50"] is not None else None,
            )
        )
    return scenarios


@app.post("/oncho/top3", dependencies=[Depends(api_key_auth)])
def oncho_top3(request: Top3Request):
    session = SessionLocal()
    try:
        results = run_top3_query(session, request)
        return results
    except HTTPException:
        session.rollback()
        raise
    except Exception as exc:
        session.rollback()
        raise HTTPException(status_code=400, detail=str(exc))
    finally:
        session.close()


@app.post("/oncho/execute_query", dependencies=[Depends(api_key_auth)])
def oncho_execute_query(query_data: SQLQuery):
    session = SessionLocal()
    try:
        # Automatically quote identifiers in the query for PostgreSQL
        quoted_query = quote_identifiers(query_data.query, is_postgres=True)

        result = session.execute(text(quoted_query))
        rows = result.fetchall()
        column_names = list(result.keys())

        stats = estimate_result_size(rows, guardrail_thresholds)
        logger.info(
            "sql_query_stats|endpoint=oncho_execute_query|rows=%s|cols=%s|bytes=%s|over_cap=%s",
            stats.row_count,
            stats.column_count,
            stats.bytes_estimate,
            stats.requires_guardrail,
        )

        if stats.requires_guardrail:
            return build_guardrail_payload(
                rows,
                column_names,
                stats,
                guardrail_thresholds,
            )

        result_list = [dict(zip(column_names, row)) for row in rows]
        return result_list

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        session.close()
