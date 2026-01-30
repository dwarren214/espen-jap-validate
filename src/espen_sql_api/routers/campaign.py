"""Campaign Hub endpoints - external API + SQLite cache."""

import logging
import os
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import httpx
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy import (
    create_engine,
    text,
    Column,
    String,
    Integer,
    Float,
    MetaData,
    Table,
)
from sqlalchemy.dialects.sqlite import Insert
from sqlalchemy.orm import sessionmaker

from ..auth import api_key_auth
from ..config import BASE_PATH, ESPEN_CAMPAIGN_HUB_KEY, guardrail_thresholds
from ..response_guard import estimate_result_size, build_guardrail_payload
from ..schemas import SQLQuery
from ..utils import validate_query_safety

logger = logging.getLogger(__name__)

router = APIRouter(tags=["campaign"])

ESPEN_CAMPAIGN_TABLE_NAME = "campaigns"


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


def ensure_campaigns_db_exists():
    today = datetime.now().date()
    db_path = get_campaign_db_path(today)

    if not db_path.exists():
        logger.info(f"Database for {today} does not exist. Fetching data from API and creating new database.")
        data = fetch_campaign_data()

        if not db_path.exists():
            create_db_from_data(db_path, data)
            remove_old_dbs()


def create_db_from_data(db_path: Path, data):
    """Creates a new campaigns table with the latest data from the API."""
    engine = create_engine(f'sqlite:///{db_path}')
    metadata = MetaData()

    single_record = data[0]
    column_info = [get_column_name_and_type(key, value) for key, value in single_record.items()]

    logger.info(f"Creating new table {ESPEN_CAMPAIGN_TABLE_NAME} in {db_path}")
    columns = [
        Column(column_name, column_type, primary_key=(column_name == "campaign_id"))
        for column_name, column_type in column_info
    ]

    table = Table(ESPEN_CAMPAIGN_TABLE_NAME, metadata, *columns)
    metadata.create_all(engine)

    today = datetime.now().date()
    with campaign_hub_session(today) as session:
        with session.begin():
            for record in data:
                record = {key_to_column_name(key): value for key, value in record.items()}
                insert_stmt = Insert(table).values(**record)
                session.execute(insert_stmt)
            session.commit()


def key_to_column_name(key: str) -> str:
    return key.lower().replace(" ", "_").replace("-", "_").replace("(", "").replace(")", "")


def get_column_name_and_type(key: str, value) -> tuple[str, type]:
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
            date_str = db_file.stem.split("-")[1:]
            db_date = datetime.strptime("-".join(date_str), "%Y-%m-%d").date()
            if (today - db_date).days > 2:
                logger.info(f"Removing old database file: {db_file}")
                db_file.unlink()
        except Exception as e:
            logger.exception("Unable to remove db file %s: %s", db_file, e)


@router.post("/query_campaign_hub_data", dependencies=[Depends(api_key_auth)])
def query_campaign_hub_data(query_data: SQLQuery):
    """Execute SQL query against cached campaign data."""
    is_safe, error_msg = validate_query_safety(query_data.query)
    if not is_safe:
        logger.warning("query_validation_failed|endpoint=query_campaign_hub_data|query=%s|reason=%s", query_data.query[:100], error_msg)
        raise HTTPException(status_code=400, detail=error_msg)

    today = datetime.now().date()
    ensure_campaigns_db_exists()
    try:
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
            return build_guardrail_payload(rows, columns, stats, guardrail_thresholds)

        return {"data": result_data, "row_count": len(result_data)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.get("/fetch_campaign_hub_columns", dependencies=[Depends(api_key_auth)])
def fetch_campaign_hub_columns():
    ensure_campaigns_db_exists()

    today = datetime.now().date()
    with campaign_hub_session(today) as session:
        try:
            query = text(f"PRAGMA table_info({ESPEN_CAMPAIGN_TABLE_NAME});")
            results = session.execute(query).fetchall()
            return [res[1] for res in results]
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))


@router.get("/fetch_campaign_hub_data", dependencies=[Depends(api_key_auth)])
def fetch_campaign_hub_data():
    """Legacy endpoint - fetches campaign data directly from API."""
    previous_year = datetime.now().year - 1
    headers = {"access_token": os.getenv("ESPEN_CAMPAIGN_HUB_KEY")}
    response = httpx.get(url="https://lbdatabaseapi.azurewebsites.net/campaign_hub_download", headers=headers)

    if response.status_code == 200:
        records = response.json()
        final_response = []
        for record in records:
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
