"""Story 2 tests for JAP upload endpoint and file-reference lifecycle."""

import json
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

os.environ.setdefault("API_KEY", "test-key")

from espen_sql_api.jap_validation import storage  # noqa: E402
from espen_sql_api.routers import jap_validation  # noqa: E402


def _configure_storage(monkeypatch: pytest.MonkeyPatch, base_dir: Path, ttl_seconds: int = 60) -> None:
    upload_dir = base_dir / "uploads"
    metadata_dir = base_dir / "metadata"
    runs_dir = base_dir / "runs"

    monkeypatch.setattr(storage, "JAP_UPLOAD_BACKEND", "local_fs")
    monkeypatch.setattr(storage, "JAP_UPLOAD_DIR", upload_dir)
    monkeypatch.setattr(storage, "JAP_METADATA_DIR", metadata_dir)
    monkeypatch.setattr(storage, "JAP_RUNS_DIR", runs_dir)
    monkeypatch.setattr(storage, "JAP_FILE_TTL_SECONDS", ttl_seconds)


@pytest.fixture
def auth_headers():
    return {"Authorization": "Bearer test-key"}


@pytest.fixture
def client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    _configure_storage(monkeypatch, tmp_path / "jap_validation", ttl_seconds=60)
    monkeypatch.setattr(jap_validation, "JAP_MAX_UPLOAD_BYTES", 1024 * 1024)

    app = FastAPI()
    app.include_router(jap_validation.router)
    return TestClient(app)


def test_upload_success(client: TestClient, auth_headers):
    payload_bytes = b"dummy workbook content"
    response = client.post(
        "/upload",
        headers=auth_headers,
        files={"file": ("sample.xlsm", payload_bytes, "application/vnd.ms-excel.sheet.macroEnabled.12")},
        data={"source": "ocs_python_node"},
    )

    assert response.status_code == 200
    body = response.json()
    assert body["file_reference"].startswith("upl_")
    assert body["file_name"] == "sample.xlsm"
    assert body["content_type"] == "application/vnd.ms-excel.sheet.macroEnabled.12"
    assert body["size_bytes"] == len(payload_bytes)
    assert body["storage_backend"] == "local_fs"
    assert body["ttl_seconds"] == 60
    assert set(body.keys()) == {
        "file_reference",
        "file_name",
        "content_type",
        "size_bytes",
        "storage_backend",
        "uploaded_at",
        "expires_at",
        "ttl_seconds",
    }

    metadata = storage.get_upload_metadata(body["file_reference"])
    assert metadata is not None
    upload_path = storage.resolve_upload_path(body["file_reference"])
    assert upload_path is not None
    assert upload_path.exists()


def test_upload_invalid_extension_returns_400(client: TestClient, auth_headers):
    response = client.post(
        "/upload",
        headers=auth_headers,
        files={"file": ("not_allowed.txt", b"hello", "text/plain")},
    )
    assert response.status_code == 400


def test_upload_empty_file_returns_400(client: TestClient, auth_headers):
    response = client.post(
        "/upload",
        headers=auth_headers,
        files={"file": ("empty.xlsx", b"", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 400


def test_upload_oversize_returns_413(client: TestClient, auth_headers, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setattr(jap_validation, "JAP_MAX_UPLOAD_BYTES", 4)
    response = client.post(
        "/upload",
        headers=auth_headers,
        files={"file": ("big.xlsx", b"12345", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 413


def test_metadata_lookup_and_expiry_behavior(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    _configure_storage(monkeypatch, tmp_path / "jap_validation", ttl_seconds=120)

    saved = storage.save_upload(
        file_name="metadata_test.xlsx",
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        data=b"abc",
        source="unit_test",
    )
    assert saved.file_reference.startswith("upl_")

    loaded = storage.get_upload_metadata(saved.file_reference)
    assert loaded is not None
    assert loaded.file_reference == saved.file_reference
    assert not storage.is_expired(loaded)

    metadata_path = storage.JAP_METADATA_DIR / f"{saved.file_reference}.json"
    with metadata_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    payload["expires_at"] = (
        datetime.now(timezone.utc) - timedelta(seconds=5)
    ).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    with metadata_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=True, indent=2)

    removed = storage.cleanup_expired_uploads()
    assert removed == 1
    assert storage.get_upload_metadata(saved.file_reference) is None
    assert storage.resolve_upload_path(saved.file_reference) is None
