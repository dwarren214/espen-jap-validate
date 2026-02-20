"""Storage primitives for JAP workbook upload lifecycle."""

from __future__ import annotations

import json
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

from ..config import (
    JAP_FILE_TTL_SECONDS,
    JAP_METADATA_DIR,
    JAP_RUNS_DIR,
    JAP_UPLOAD_BACKEND,
    JAP_UPLOAD_DIR,
)
from .models import UploadMetadata

logger = logging.getLogger(__name__)

LOCAL_FS_BACKEND = "local_fs"


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _format_utc(value: datetime) -> str:
    return value.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _parse_utc(value: str) -> datetime:
    normalized = value[:-1] + "+00:00" if value.endswith("Z") else value
    return datetime.fromisoformat(normalized).astimezone(timezone.utc)


def _metadata_path(file_reference: str) -> Path:
    return JAP_METADATA_DIR / f"{file_reference}.json"


def _ensure_storage_dirs() -> None:
    JAP_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    JAP_METADATA_DIR.mkdir(parents=True, exist_ok=True)
    JAP_RUNS_DIR.mkdir(parents=True, exist_ok=True)


def _generate_file_reference() -> str:
    return f"upl_{uuid4().hex}"


def save_upload(
    *,
    file_name: str,
    content_type: str,
    data: bytes,
    source: str | None = None,
) -> UploadMetadata:
    """Persist upload bytes + metadata and return lifecycle metadata."""
    if JAP_UPLOAD_BACKEND != LOCAL_FS_BACKEND:
        raise RuntimeError(f"Unsupported JAP_UPLOAD_BACKEND '{JAP_UPLOAD_BACKEND}'.")

    _ensure_storage_dirs()

    safe_name = Path(file_name).name or "upload.xlsx"
    extension = Path(safe_name).suffix.lower()
    file_reference = _generate_file_reference()
    stored_file_name = f"{file_reference}{extension}"
    upload_path = JAP_UPLOAD_DIR / stored_file_name
    metadata_path = _metadata_path(file_reference)

    uploaded_at = _utc_now()
    expires_at = uploaded_at + timedelta(seconds=JAP_FILE_TTL_SECONDS)

    metadata = UploadMetadata(
        file_reference=file_reference,
        file_name=safe_name,
        content_type=content_type,
        size_bytes=len(data),
        storage_backend=JAP_UPLOAD_BACKEND,
        uploaded_at=_format_utc(uploaded_at),
        expires_at=_format_utc(expires_at),
        ttl_seconds=JAP_FILE_TTL_SECONDS,
        stored_file_name=stored_file_name,
        source=source,
    )

    try:
        upload_path.write_bytes(data)
        with metadata_path.open("w", encoding="utf-8") as handle:
            json.dump(metadata.model_dump(), handle, ensure_ascii=True, indent=2)
    except Exception:
        if upload_path.exists():
            upload_path.unlink()
        if metadata_path.exists():
            metadata_path.unlink()
        raise

    return metadata


def get_upload_metadata(file_reference: str) -> UploadMetadata | None:
    """Fetch persisted metadata sidecar by file reference."""
    metadata_path = _metadata_path(file_reference)
    if not metadata_path.exists():
        return None

    with metadata_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    return UploadMetadata.model_validate(payload)


def resolve_upload_path(file_reference: str) -> Path | None:
    """Resolve local upload path from file reference metadata."""
    metadata = get_upload_metadata(file_reference)
    if metadata is None:
        return None
    return JAP_UPLOAD_DIR / metadata.stored_file_name


def is_expired(metadata: UploadMetadata | dict[str, Any]) -> bool:
    """Return True when metadata expiry timestamp has passed."""
    expires_at = metadata.expires_at if isinstance(metadata, UploadMetadata) else metadata["expires_at"]
    return _utc_now() >= _parse_utc(expires_at)


def cleanup_expired_uploads() -> int:
    """Delete expired upload + metadata files; returns number of cleaned references."""
    if not JAP_METADATA_DIR.exists():
        return 0

    removed_count = 0
    for metadata_file in JAP_METADATA_DIR.glob("*.json"):
        try:
            with metadata_file.open("r", encoding="utf-8") as handle:
                payload = json.load(handle)
            metadata = UploadMetadata.model_validate(payload)
            if not is_expired(metadata):
                continue

            upload_path = JAP_UPLOAD_DIR / metadata.stored_file_name
            if upload_path.exists():
                upload_path.unlink()
            metadata_file.unlink()
            removed_count += 1
        except Exception:
            logger.exception("Failed cleanup for metadata file '%s'.", metadata_file)

    return removed_count
