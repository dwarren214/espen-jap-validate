"""MSSQL endpoints for ESPEN analytical tables + SQLite metadata."""

import csv
import io
import logging

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import StreamingResponse
from sqlalchemy import text
import orjson

from ..auth import api_key_auth
from ..config import guardrail_thresholds
from ..db import RemoteSessionLocal, MetaSessionLocal
from ..response_guard import estimate_result_size, build_guardrail_payload
from ..schemas import SQLQuery, TableNames
from ..utils import quote_identifiers, execute_query_with_retry, validate_query_safety

logger = logging.getLogger(__name__)

router = APIRouter(tags=["espen"])

# Constants for response format thresholds
MAX_JSON_ROWS = 100
MAX_JSON_CELLS = 1000


@router.post("/execute_sql_query", dependencies=[Depends(api_key_auth)])
def execute_sql_query(query_data: SQLQuery):
    is_safe, error_msg = validate_query_safety(query_data.query)
    if not is_safe:
        logger.warning("query_validation_failed|query=%s|reason=%s", query_data.query[:100], error_msg)
        raise HTTPException(status_code=400, detail=error_msg)

    session = RemoteSessionLocal()
    try:
        quoted_query = quote_identifiers(query_data.query, is_postgres=False)
        result = execute_query_with_retry(session, quoted_query)
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
            return build_guardrail_payload(rows, column_names, stats, guardrail_thresholds)

        num_rows = len(rows)
        num_columns = len(column_names)
        total_cells = num_rows * num_columns

        should_return_csv = query_data.force_csv or (
            not query_data.force_json
            and (num_rows > MAX_JSON_ROWS or total_cells > MAX_JSON_CELLS)
        )

        if should_return_csv:
            output = io.StringIO()
            writer = csv.writer(output)
            writer.writerow(column_names)
            for row in rows:
                writer.writerow(row)
            output.seek(0)
            return StreamingResponse(
                iter([output.getvalue()]),
                media_type="text/csv",
                headers={"Content-Disposition": "attachment; filename=query_results.csv"},
            )
        else:
            return [dict(zip(column_names, row)) for row in rows]

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        session.close()


@router.post("/fetch_column_names_and_types", dependencies=[Depends(api_key_auth)])
def fetch_column_names_and_types(table_data: TableNames):
    meta_session = MetaSessionLocal()
    results = {}
    try:
        for table_name in table_data.tables:
            query = text("""
                SELECT "Fields"
                FROM espen_tables
                WHERE LOWER("Name_Analytical_Table") = :table_name
            """)
            res = meta_session.execute(query, {"table_name": table_name.lower()}).fetchone()
            if res:
                fields = orjson.loads(res[0])
                results[table_name] = {field["Name"]: field["Description"] for field in fields}
            else:
                results[table_name] = "Table not found"
        return results
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        meta_session.close()
