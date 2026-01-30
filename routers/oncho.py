"""PostgreSQL endpoints for oncho projection data."""

import logging
from typing import List, Literal, Optional

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, field_validator, FieldValidationInfo
from sqlalchemy import bindparam, text
from sqlalchemy.orm import Session

from auth import api_key_auth
from config import guardrail_thresholds
from db import SessionLocal
from response_guard import estimate_result_size, build_guardrail_payload
from schemas import SQLQuery
from utils import quote_identifiers, execute_query_with_retry, validate_query_safety

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/oncho", tags=["oncho"])


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


@router.post("/top3", dependencies=[Depends(api_key_auth)])
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


@router.post("/execute_query", dependencies=[Depends(api_key_auth)])
def oncho_execute_query(query_data: SQLQuery):
    is_safe, error_msg = validate_query_safety(query_data.query)
    if not is_safe:
        logger.warning("query_validation_failed|endpoint=oncho_execute_query|query=%s|reason=%s", query_data.query[:100], error_msg)
        raise HTTPException(status_code=400, detail=error_msg)

    session = SessionLocal()
    try:
        quoted_query = quote_identifiers(query_data.query, is_postgres=True)
        result = execute_query_with_retry(session, quoted_query)
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
            return build_guardrail_payload(rows, column_names, stats, guardrail_thresholds)

        return [dict(zip(column_names, row)) for row in rows]

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        session.close()
