"""JAP validation endpoints."""

import logging
from pathlib import Path
from time import perf_counter

from fastapi import APIRouter, Depends, File, Form, Request, Response, UploadFile

from ..auth import api_key_auth
from ..config import JAP_MAX_UPLOAD_BYTES
from ..jap_validation.api_support import (
    JAPValidationRoute,
    REQUEST_ID_HEADER,
    duration_ms,
    error_response,
    get_or_create_correlation_id,
    log_jap_event,
)
from ..jap_validation.models import UploadResponse, ValidateJRSMRequest, ValidateJRSMResponse
from ..jap_validation.response_builder import build_validation_response
from ..jap_validation.storage import (
    cleanup_expired_uploads,
    get_upload_metadata,
    is_expired,
    resolve_upload_path,
    save_upload,
)
from ..jap_validation.validators.jrsm import validate_jrsm_workbook

logger = logging.getLogger(__name__)

router = APIRouter(tags=["jap_validation"], route_class=JAPValidationRoute)

ALLOWED_UPLOAD_EXTENSIONS = {".xlsx", ".xlsm"}


@router.post("/upload", dependencies=[Depends(api_key_auth)], response_model=UploadResponse)
async def upload_workbook(
    request: Request,
    response: Response,
    file: UploadFile = File(...),
    source: str | None = Form(default=None),
    original_filename: str | None = Form(default=None),
):
    """Upload JRSM workbook and return a time-bound file reference."""
    started_at = perf_counter()
    correlation_id = get_or_create_correlation_id(request)
    endpoint = "/upload"

    try:
        cleaned_count = cleanup_expired_uploads()
        if cleaned_count:
            log_jap_event(
                logger,
                event="jap_cleanup_completed",
                correlation_id=correlation_id,
                endpoint=endpoint,
                cleaned_references=cleaned_count,
            )
    except Exception:
        logger.exception(
            "event=jap_cleanup_failed|correlation_id=%s|endpoint=%s",
            correlation_id,
            endpoint,
        )

    effective_name = Path(original_filename or file.filename or "").name
    extension = Path(effective_name).suffix.lower()
    if extension not in ALLOWED_UPLOAD_EXTENSIONS:
        log_jap_event(
            logger,
            event="jap_upload_failed",
            correlation_id=correlation_id,
            endpoint=endpoint,
            level=logging.WARNING,
            error_code="UNSUPPORTED_FILE_TYPE",
            duration_ms=duration_ms(started_at),
            file_name=effective_name or None,
            received_extension=extension or "MISSING",
        )
        await file.close()
        return error_response(
            status_code=400,
            error_code="UNSUPPORTED_FILE_TYPE",
            correlation_id=correlation_id,
            details={
                "allowed_extensions": sorted(ALLOWED_UPLOAD_EXTENSIONS),
                "received_extension": extension or None,
            },
        )

    file_bytes = await file.read()
    if not file_bytes:
        log_jap_event(
            logger,
            event="jap_upload_failed",
            correlation_id=correlation_id,
            endpoint=endpoint,
            level=logging.WARNING,
            error_code="ATTACHMENT_MISSING",
            duration_ms=duration_ms(started_at),
            file_name=effective_name or None,
        )
        await file.close()
        return error_response(
            status_code=400,
            error_code="ATTACHMENT_MISSING",
            correlation_id=correlation_id,
            message="The uploaded workbook is empty.",
            details={"file_name": effective_name or None},
        )
    if len(file_bytes) > JAP_MAX_UPLOAD_BYTES:
        log_jap_event(
            logger,
            event="jap_upload_failed",
            correlation_id=correlation_id,
            endpoint=endpoint,
            level=logging.WARNING,
            error_code="FILE_TOO_LARGE",
            duration_ms=duration_ms(started_at),
            file_name=effective_name or None,
            size_bytes=len(file_bytes),
            max_upload_bytes=JAP_MAX_UPLOAD_BYTES,
        )
        await file.close()
        return error_response(
            status_code=413,
            error_code="FILE_TOO_LARGE",
            correlation_id=correlation_id,
            details={
                "file_name": effective_name or None,
                "size_bytes": len(file_bytes),
                "max_upload_bytes": JAP_MAX_UPLOAD_BYTES,
            },
        )

    content_type = file.content_type or "application/octet-stream"
    try:
        metadata = save_upload(
            file_name=effective_name,
            content_type=content_type,
            data=file_bytes,
            source=source,
        )
    except Exception:
        logger.exception(
            "event=jap_upload_failed|correlation_id=%s|endpoint=%s|error_code=INTERNAL_STORAGE_ERROR",
            correlation_id,
            endpoint,
        )
        return error_response(
            status_code=500,
            error_code="INTERNAL_STORAGE_ERROR",
            correlation_id=correlation_id,
            details={"file_name": effective_name or None},
        )
    finally:
        await file.close()

    response.headers[REQUEST_ID_HEADER] = correlation_id
    log_jap_event(
        logger,
        event="jap_upload_completed",
        correlation_id=correlation_id,
        endpoint=endpoint,
        duration_ms=duration_ms(started_at),
        file_reference=metadata.file_reference,
        file_name=metadata.file_name,
        size_bytes=metadata.size_bytes,
        ttl_seconds=metadata.ttl_seconds,
        storage_backend=metadata.storage_backend,
    )
    return UploadResponse.model_validate(metadata.model_dump())


@router.post(
    "/validate/jrsm",
    dependencies=[Depends(api_key_auth)],
    response_model=ValidateJRSMResponse,
    response_model_exclude_none=True,
)
def validate_jrsm(
    request: Request,
    http_response: Response,
    payload: ValidateJRSMRequest,
):
    """Orchestrate JRSM validation from stored file reference."""
    started_at = perf_counter()
    correlation_id = get_or_create_correlation_id(request)
    endpoint = "/validate/jrsm"
    metadata = get_upload_metadata(payload.file_reference)

    try:
        cleaned_count = cleanup_expired_uploads()
        if cleaned_count:
            log_jap_event(
                logger,
                event="jap_cleanup_completed",
                correlation_id=correlation_id,
                endpoint=endpoint,
                cleaned_references=cleaned_count,
            )
    except Exception:
        logger.exception(
            "event=jap_cleanup_failed|correlation_id=%s|endpoint=%s",
            correlation_id,
            endpoint,
        )
    if metadata is None:
        log_jap_event(
            logger,
            event="jap_validate_failed",
            correlation_id=correlation_id,
            endpoint=endpoint,
            level=logging.WARNING,
            error_code="FILE_REFERENCE_NOT_FOUND",
            duration_ms=duration_ms(started_at),
            file_reference=payload.file_reference,
        )
        return error_response(
            status_code=404,
            error_code="FILE_REFERENCE_NOT_FOUND",
            correlation_id=correlation_id,
            details={"file_reference": payload.file_reference},
        )
    if is_expired(metadata):
        log_jap_event(
            logger,
            event="jap_validate_failed",
            correlation_id=correlation_id,
            endpoint=endpoint,
            level=logging.WARNING,
            error_code="FILE_REFERENCE_EXPIRED",
            duration_ms=duration_ms(started_at),
            file_reference=payload.file_reference,
            expires_at=metadata.expires_at,
        )
        return error_response(
            status_code=410,
            error_code="FILE_REFERENCE_EXPIRED",
            correlation_id=correlation_id,
            details={
                "file_reference": payload.file_reference,
                "expires_at": metadata.expires_at,
            },
        )

    workbook_path = resolve_upload_path(payload.file_reference)
    if workbook_path is None or not workbook_path.exists():
        log_jap_event(
            logger,
            event="jap_validate_failed",
            correlation_id=correlation_id,
            endpoint=endpoint,
            level=logging.WARNING,
            error_code="FILE_REFERENCE_NOT_FOUND",
            duration_ms=duration_ms(started_at),
            file_reference=payload.file_reference,
        )
        return error_response(
            status_code=404,
            error_code="FILE_REFERENCE_NOT_FOUND",
            correlation_id=correlation_id,
            details={"file_reference": payload.file_reference},
        )

    try:
        validation_result = validate_jrsm_workbook(
            workbook_path=workbook_path,
            country=payload.country,
            year_for_request_of_medicine=payload.year_for_request_of_medicine,
            metadata=payload.metadata,
        )
        validation_response = build_validation_response(
            file_reference=payload.file_reference,
            validation_result=validation_result,
        )
    except Exception:
        logger.exception(
            "event=jap_validate_failed|correlation_id=%s|endpoint=%s|error_code=INTERNAL_STORAGE_ERROR|file_reference=%s",
            correlation_id,
            endpoint,
            payload.file_reference,
        )
        return error_response(
            status_code=500,
            error_code="INTERNAL_STORAGE_ERROR",
            correlation_id=correlation_id,
            details={"file_reference": payload.file_reference},
        )

    http_response.headers[REQUEST_ID_HEADER] = correlation_id
    summary = validation_response.executive_summary
    log_jap_event(
        logger,
        event="jap_validate_completed",
        correlation_id=correlation_id,
        endpoint=endpoint,
        duration_ms=duration_ms(started_at),
        file_reference=payload.file_reference,
        run_id=validation_response.run_id,
        validation_outcome=validation_response.validation_outcome,
        error_count=summary.error_count,
        warn_count=summary.warn_count,
        info_count=summary.info_count,
        total_findings=summary.total_findings,
    )

    return validation_response
