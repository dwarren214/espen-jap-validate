"""Utilities for estimating query result sizes before serialization."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Sequence, List, Dict, Optional


@dataclass(frozen=True)
class GuardrailThresholds:
    """Configuration for safe result sizes."""

    max_rows: int
    max_bytes: int
    preview_rows: int


@dataclass(frozen=True)
class QueryResultStats:
    """Measured or estimated characteristics of a SQL result set."""

    row_count: int
    column_count: int
    bytes_estimate: int

    @property
    def exceeds_row_cap(self) -> bool:
        return self.row_count > self.max_rows_bound

    @property
    def exceeds_byte_cap(self) -> bool:
        return self.bytes_estimate > self.max_bytes_bound

    @property
    def requires_guardrail(self) -> bool:
        return self.exceeds_row_cap or self.exceeds_byte_cap

    # internal bounds (injected at construction time)
    max_rows_bound: int = 0
    max_bytes_bound: int = 0


def estimate_result_size(
    rows: Sequence[Sequence[object]],
    thresholds: GuardrailThresholds,
) -> QueryResultStats:
    """Compute coarse metrics for a result set.

    The byte estimate intentionally keeps things simple (string lengths) to avoid
    large allocations while still providing a safety signal for oversized
    responses.
    """

    row_count = len(rows)
    column_count = len(rows[0]) if rows else 0

    bytes_estimate = _estimate_bytes(rows)

    return QueryResultStats(
        row_count=row_count,
        column_count=column_count,
        bytes_estimate=bytes_estimate,
        max_rows_bound=thresholds.max_rows,
        max_bytes_bound=thresholds.max_bytes,
    )


def _estimate_bytes(rows: Sequence[Sequence[object]]) -> int:
    total = 0
    for row in rows:
        for value in row:
            if value is None:
                total += 4  # allow room for "null"
            elif isinstance(value, (int, float)):
                total += len(str(value))
            else:
                text = str(value)
                total += len(text.encode("utf-8", "ignore"))
        # comma + newline separators
        total += 2
    return total


def slice_preview(
    rows: Sequence[Sequence[object]],
    preview_rows: int,
) -> Iterable[Sequence[object]]:
    """Return the leading rows for preview purposes."""

    if preview_rows <= 0:
        return []
    return list(rows[:preview_rows])


def build_guardrail_payload(
    rows: Sequence[Sequence[object]],
    column_names: Sequence[str],
    stats: QueryResultStats,
    thresholds: GuardrailThresholds,
    preview_rows_override: Optional[int] = None,
    requested_max_rows: Optional[int] = None,
) -> Dict[str, object]:
    preview_rows = preview_rows_override or thresholds.preview_rows
    preview = slice_preview(rows, preview_rows)
    preview_dicts = [dict(zip(column_names, row)) for row in preview]

    payload: Dict[str, object] = {
        "status": "too_large",
        "row_count": stats.row_count,
        "column_count": stats.column_count,
        "bytes_estimate": stats.bytes_estimate,
        "preview_row_count": len(preview_dicts),
        "preview": preview_dicts,
        "suggested_filters": ["country", "disease", "year"],
    }

    if requested_max_rows is not None:
        payload["max_rows_requested"] = requested_max_rows

    return payload
