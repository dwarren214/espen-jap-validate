"""JAP validation endpoints."""

import logging
from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile

from ..auth import api_key_auth
from ..config import JAP_MAX_UPLOAD_BYTES
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

router = APIRouter(tags=["jap_validation"])

ALLOWED_UPLOAD_EXTENSIONS = {".xlsx", ".xlsm"}


@router.post("/upload", dependencies=[Depends(api_key_auth)], response_model=UploadResponse)
async def upload_workbook(
    file: UploadFile = File(...),
    source: str | None = Form(default=None),
    original_filename: str | None = Form(default=None),
):
    """Upload JRSM workbook and return a time-bound file reference."""
    try:
        cleanup_expired_uploads()
    except Exception:
        logger.exception("cleanup_expired_uploads_failed")

    effective_name = Path(original_filename or file.filename or "").name
    extension = Path(effective_name).suffix.lower()
    if extension not in ALLOWED_UPLOAD_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Invalid file extension. Allowed: .xlsx, .xlsm")

    file_bytes = await file.read()
    if not file_bytes:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    if len(file_bytes) > JAP_MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="Uploaded file exceeds JAP_MAX_UPLOAD_BYTES.")

    content_type = file.content_type or "application/octet-stream"
    try:
        metadata = save_upload(
            file_name=effective_name,
            content_type=content_type,
            data=file_bytes,
            source=source,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        await file.close()

    return UploadResponse.model_validate(metadata.model_dump())


@router.post(
    "/validate/jrsm",
    dependencies=[Depends(api_key_auth)],
    response_model=ValidateJRSMResponse,
    response_model_exclude_none=True,
)
def validate_jrsm(request: ValidateJRSMRequest):
    """Orchestrate JRSM validation from stored file reference."""
    metadata = get_upload_metadata(request.file_reference)
    if metadata is None:
        raise HTTPException(status_code=404, detail="Unknown file_reference.")
    if is_expired(metadata):
        raise HTTPException(status_code=410, detail="Expired file_reference.")

    workbook_path = resolve_upload_path(request.file_reference)
    if workbook_path is None or not workbook_path.exists():
        raise HTTPException(status_code=404, detail="Unknown file_reference.")

    try:
        validation_result = validate_jrsm_workbook(
            workbook_path=workbook_path,
            country=request.country,
            year_for_request_of_medicine=request.year_for_request_of_medicine,
            metadata=request.metadata,
        )
        response = build_validation_response(
            file_reference=request.file_reference,
            validation_result=validation_result,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    return response
