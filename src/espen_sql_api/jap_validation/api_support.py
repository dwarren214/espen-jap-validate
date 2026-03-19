"""JAP API helpers for structured errors, tracing, and lightweight logging."""

from __future__ import annotations

import json
import logging
from time import perf_counter
from typing import Any
from uuid import uuid4

from fastapi import HTTPException, Request
from fastapi.exceptions import RequestValidationError
from fastapi.routing import APIRoute
from fastapi.responses import JSONResponse

REQUEST_ID_HEADER = "X-Request-ID"
route_logger = logging.getLogger("espen_sql_api.routers.jap_validation")

ERROR_CODE_MESSAGES = {
    "ATTACHMENT_MISSING": "A workbook attachment is required.",
    "AUTHENTICATION_FAILED": "Authentication failed for this request.",
    "UNSUPPORTED_FILE_TYPE": "Only .xlsx and .xlsm workbooks are supported.",
    "FILE_TOO_LARGE": "The uploaded workbook exceeds the maximum allowed size.",
    "FILE_REFERENCE_NOT_FOUND": "The uploaded file reference could not be found.",
    "FILE_REFERENCE_EXPIRED": "The uploaded file reference has expired. Please re-upload the workbook.",
    "UNSUPPORTED_FORM_TYPE": "This endpoint does not support the requested form type.",
    "REQUEST_VALIDATION_ERROR": "The request body did not match the expected schema.",
    "WORKBOOK_PARSE_FAILED": "The workbook could not be parsed.",
    "REQUIRED_SHEET_MISSING": "The workbook is missing a required sheet.",
    "REQUIRED_CELL_MISSING": "The workbook is missing a required cell.",
    "INTERNAL_STORAGE_ERROR": "The validation service encountered an internal processing error.",
}


def get_or_create_correlation_id(request: Request) -> str:
    """Resolve an inbound request ID or generate one for JAP endpoints."""
    existing = getattr(request.state, "correlation_id", None)
    if existing:
        return existing

    inbound = request.headers.get(REQUEST_ID_HEADER, "").strip()
    correlation_id = inbound or f"req_{uuid4().hex}"
    request.state.correlation_id = correlation_id
    return correlation_id


def build_error_payload(
    *,
    error_code: str,
    correlation_id: str,
    message: str | None = None,
    details: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build the canonical JAP error payload."""
    return {
        "error_code": error_code,
        "message": message or ERROR_CODE_MESSAGES[error_code],
        "details": details or {},
        "correlation_id": correlation_id,
    }


def error_response(
    *,
    status_code: int,
    error_code: str,
    correlation_id: str,
    message: str | None = None,
    details: dict[str, Any] | None = None,
) -> JSONResponse:
    """Return a structured JAP error response with request ID header."""
    response = JSONResponse(
        status_code=status_code,
        content=build_error_payload(
            error_code=error_code,
            message=message,
            details=details,
            correlation_id=correlation_id,
        ),
    )
    response.headers[REQUEST_ID_HEADER] = correlation_id
    return response


def duration_ms(start_time: float) -> int:
    """Return elapsed time in milliseconds from a perf_counter start."""
    return int((perf_counter() - start_time) * 1000)


def log_jap_event(
    logger: logging.Logger,
    *,
    event: str,
    correlation_id: str,
    endpoint: str,
    level: int = logging.INFO,
    **fields: Any,
) -> None:
    """Emit a compact key-value log line for JAP endpoint activity."""
    parts = [
        f"event={event}",
        f"correlation_id={correlation_id}",
        f"endpoint={endpoint}",
    ]
    for key, value in fields.items():
        if value is None:
            continue
        parts.append(f"{key}={_format_log_value(value)}")
    logger.log(level, "|".join(parts))


def _format_log_value(value: Any) -> str:
    if isinstance(value, (dict, list, tuple)):
        return json.dumps(value, sort_keys=True, separators=(",", ":"))
    text = str(value).strip()
    text = " ".join(text.split())
    return text.replace("|", "/")


class JAPValidationRoute(APIRoute):
    """JAP-only route wrapper to normalize tracing and framework error responses."""

    def get_route_handler(self):
        original_handler = super().get_route_handler()

        async def custom_handler(request: Request):
            correlation_id = get_or_create_correlation_id(request)
            endpoint = request.url.path

            try:
                response = await original_handler(request)
            except RequestValidationError as exc:
                log_jap_event(
                    route_logger,
                    event="jap_request_failed",
                    correlation_id=correlation_id,
                    endpoint=endpoint,
                    level=logging.WARNING,
                    error_code="REQUEST_VALIDATION_ERROR",
                )
                return error_response(
                    status_code=422,
                    error_code="REQUEST_VALIDATION_ERROR",
                    correlation_id=correlation_id,
                    details={"errors": exc.errors()},
                )
            except HTTPException as exc:
                error_code = "AUTHENTICATION_FAILED" if exc.status_code == 403 else "REQUEST_VALIDATION_ERROR"
                log_jap_event(
                    route_logger,
                    event="jap_request_failed",
                    correlation_id=correlation_id,
                    endpoint=endpoint,
                    level=logging.WARNING,
                    error_code=error_code,
                    status_code=exc.status_code,
                )
                details = {"detail": exc.detail}
                if exc.headers:
                    details["headers"] = exc.headers
                return error_response(
                    status_code=exc.status_code,
                    error_code=error_code,
                    correlation_id=correlation_id,
                    message=str(exc.detail),
                    details=details,
                )

            response.headers[REQUEST_ID_HEADER] = correlation_id
            return response

        return custom_handler
