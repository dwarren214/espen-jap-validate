"""Story 11 tests for JAP structured errors, tracing, and retention hardening."""

import io
import json
import logging
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from openpyxl import Workbook

os.environ.setdefault("API_KEY", "test-key")

from espen_sql_api.jap_validation import storage  # noqa: E402
from espen_sql_api.routers import jap_validation  # noqa: E402

REQUIRED_SHEETS = [
    "INTRO",
    "COUNTRY_INFO",
    "DEC",
    "ALB_MBD",
    "PZQ",
    "IVM",
    "IVM+",
    "SUMMARY",
    "SHIPMENT",
    "DATA_POLICY",
]
INTRO_CELL_VALUES = {
    "B2": "Joint request for selected PC medicines v.4.4",
    "E33": "Rwanda",
    "E35": 2026,
    "E37": "Endemic",
    "E39": "Non-endemic",
    "E41": "Non-endemic",
    "E43": "Non-endemic",
    "E45": 1,
    "E48": 0.1,
    "E49": 0.2,
    "E50": 0.3,
    "E51": 0.4,
}


def _configure_storage(monkeypatch: pytest.MonkeyPatch, base_dir: Path, ttl_seconds: int = 300) -> None:
    upload_dir = base_dir / "uploads"
    metadata_dir = base_dir / "metadata"
    runs_dir = base_dir / "runs"

    monkeypatch.setattr(storage, "JAP_UPLOAD_BACKEND", "local_fs")
    monkeypatch.setattr(storage, "JAP_UPLOAD_DIR", upload_dir)
    monkeypatch.setattr(storage, "JAP_METADATA_DIR", metadata_dir)
    monkeypatch.setattr(storage, "JAP_RUNS_DIR", runs_dir)
    monkeypatch.setattr(storage, "JAP_FILE_TTL_SECONDS", ttl_seconds)


def _build_workbook_bytes() -> bytes:
    workbook = Workbook()
    intro = workbook.active
    intro.title = "INTRO"

    for sheet_name in REQUIRED_SHEETS:
        if sheet_name == "INTRO":
            continue
        workbook.create_sheet(title=sheet_name)

    for cell_address, value in INTRO_CELL_VALUES.items():
        intro[cell_address] = value

    country_info = workbook["COUNTRY_INFO"]
    country_info["B10"] = "rw"
    country_info["C10"] = "district"
    country_info["D10"] = 100
    country_info["H10"] = 1

    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


@pytest.fixture
def auth_headers():
    return {"Authorization": "Bearer test-key"}


@pytest.fixture
def client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    _configure_storage(monkeypatch, tmp_path / "jap_validation")
    monkeypatch.setattr(jap_validation, "JAP_MAX_UPLOAD_BYTES", 1024 * 1024)

    app = FastAPI()
    app.include_router(jap_validation.router)
    return TestClient(app)


def _upload_valid_workbook(client: TestClient, auth_headers: dict[str, str]) -> str:
    upload_response = client.post(
        "/upload",
        headers=auth_headers,
        files={
            "file": (
                "story11.xlsx",
                _build_workbook_bytes(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    assert upload_response.status_code == 200
    return upload_response.json()["file_reference"]


def test_upload_invalid_extension_returns_structured_error_and_echoes_request_id(client: TestClient, auth_headers):
    request_id = "req-story11-upload"
    response = client.post(
        "/upload",
        headers={**auth_headers, "X-Request-ID": request_id},
        files={"file": ("invalid.txt", b"hello", "text/plain")},
    )

    assert response.status_code == 400
    assert response.headers["X-Request-ID"] == request_id
    assert response.json() == {
        "error_code": "UNSUPPORTED_FILE_TYPE",
        "message": "Only .xlsx and .xlsm workbooks are supported.",
        "details": {
            "allowed_extensions": [".xlsm", ".xlsx"],
            "received_extension": ".txt",
        },
        "correlation_id": request_id,
    }


def test_validate_unknown_reference_returns_structured_error_and_generated_request_id(client: TestClient, auth_headers):
    response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json={
            "file_reference": "upl_missing_reference",
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
        },
    )

    body = response.json()
    assert response.status_code == 404
    assert body["error_code"] == "FILE_REFERENCE_NOT_FOUND"
    assert body["message"] == "The uploaded file reference could not be found."
    assert body["details"] == {"file_reference": "upl_missing_reference"}
    assert body["correlation_id"].startswith("req_")
    assert response.headers["X-Request-ID"] == body["correlation_id"]


def test_auth_and_schema_errors_use_structured_jap_payloads(client: TestClient):
    auth_response = client.post("/upload")
    auth_body = auth_response.json()

    assert auth_response.status_code == 403
    assert auth_body["error_code"] == "AUTHENTICATION_FAILED"
    assert auth_body["message"] == "Not authenticated"
    assert auth_body["correlation_id"].startswith("req_")
    assert auth_response.headers["X-Request-ID"] == auth_body["correlation_id"]

    schema_response = client.post(
        "/validate/jrsm",
        headers={"Authorization": "Bearer test-key"},
        json={},
    )
    schema_body = schema_response.json()

    assert schema_response.status_code == 422
    assert schema_body["error_code"] == "REQUEST_VALIDATION_ERROR"
    assert schema_body["message"] == "The request body did not match the expected schema."
    assert "errors" in schema_body["details"]
    assert schema_body["correlation_id"].startswith("req_")
    assert schema_response.headers["X-Request-ID"] == schema_body["correlation_id"]


def test_validate_expired_reference_returns_structured_error_and_logs_failure(
    client: TestClient,
    auth_headers,
    caplog: pytest.LogCaptureFixture,
):
    file_reference = _upload_valid_workbook(client, auth_headers)
    metadata_path = storage.JAP_METADATA_DIR / f"{file_reference}.json"
    with metadata_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    expires_at = (
        datetime.now(timezone.utc) - timedelta(seconds=5)
    ).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    payload["expires_at"] = expires_at
    with metadata_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=True, indent=2)

    caplog.set_level(logging.INFO, logger="espen_sql_api.routers.jap_validation")
    request_id = "req-story11-expired"
    response = client.post(
        "/validate/jrsm",
        headers={**auth_headers, "X-Request-ID": request_id},
        json={
            "file_reference": file_reference,
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
        },
    )

    assert response.status_code == 410
    assert response.headers["X-Request-ID"] == request_id
    assert response.json() == {
        "error_code": "FILE_REFERENCE_EXPIRED",
        "message": "The uploaded file reference has expired. Please re-upload the workbook.",
        "details": {
            "file_reference": file_reference,
            "expires_at": expires_at,
        },
        "correlation_id": request_id,
    }
    assert any(
        "event=jap_validate_failed" in record.message
        and "correlation_id=req-story11-expired" in record.message
        and "error_code=FILE_REFERENCE_EXPIRED" in record.message
        for record in caplog.records
    )


def test_success_responses_include_request_id_and_validation_logging(
    client: TestClient,
    auth_headers,
    caplog: pytest.LogCaptureFixture,
):
    caplog.set_level(logging.INFO, logger="espen_sql_api.routers.jap_validation")
    request_id = "req-story11-success"

    upload_response = client.post(
        "/upload",
        headers={**auth_headers, "X-Request-ID": request_id},
        files={
            "file": (
                "success.xlsx",
                _build_workbook_bytes(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    assert upload_response.status_code == 200
    assert upload_response.headers["X-Request-ID"] == request_id
    file_reference = upload_response.json()["file_reference"]

    validate_response = client.post(
        "/validate/jrsm",
        headers={**auth_headers, "X-Request-ID": request_id},
        json={
            "file_reference": file_reference,
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
        },
    )

    assert validate_response.status_code == 200
    assert validate_response.headers["X-Request-ID"] == request_id
    assert any(
        "event=jap_upload_completed" in record.message
        and "correlation_id=req-story11-success" in record.message
        for record in caplog.records
    )
    assert any(
        "event=jap_validate_completed" in record.message
        and "correlation_id=req-story11-success" in record.message
        and f"file_reference={file_reference}" in record.message
        for record in caplog.records
    )
