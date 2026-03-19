"""Story 3 tests for validate endpoint orchestration and response contract."""

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


def _configure_storage(monkeypatch: pytest.MonkeyPatch, base_dir: Path, ttl_seconds: int = 300) -> None:
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
    _configure_storage(monkeypatch, tmp_path / "jap_validation")
    monkeypatch.setattr(jap_validation, "JAP_MAX_UPLOAD_BYTES", 1024 * 1024)

    app = FastAPI()
    app.include_router(jap_validation.router)
    return TestClient(app)


def _upload_valid_workbook(client: TestClient, auth_headers: dict[str, str]) -> str:
    upload_response = client.post(
        "/upload",
        headers=auth_headers,
        files={"file": ("contract.xlsx", b"workbook-content", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert upload_response.status_code == 200
    return upload_response.json()["file_reference"]


def test_validate_success_returns_required_contract(client: TestClient, auth_headers):
    file_reference = _upload_valid_workbook(client, auth_headers)
    response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json={
            "file_reference": file_reference,
            "form_type": "jrsm",
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
        },
    )
    assert response.status_code == 200
    body = response.json()

    assert set(body.keys()) == {
        "run_id",
        "file_reference",
        "status",
        "validation_outcome",
        "summary_text",
        "executive_summary",
        "sheet_summaries",
        "findings",
        "response_text_blocks",
    }
    assert body["file_reference"] == file_reference
    assert body["status"] == "completed"
    assert isinstance(body["summary_text"], str)
    assert isinstance(body["executive_summary"], dict)
    assert isinstance(body["sheet_summaries"], list)
    assert isinstance(body["findings"], list)
    assert isinstance(body["response_text_blocks"], list)


def test_validate_unknown_file_reference_returns_404(client: TestClient, auth_headers):
    response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json={
            "file_reference": "upl_unknown_reference",
            "form_type": "jrsm",
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
        },
    )
    assert response.status_code == 404


def test_validate_expired_file_reference_returns_410(client: TestClient, auth_headers):
    file_reference = _upload_valid_workbook(client, auth_headers)
    metadata_path = storage.JAP_METADATA_DIR / f"{file_reference}.json"
    with metadata_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    payload["expires_at"] = (
        datetime.now(timezone.utc) - timedelta(seconds=5)
    ).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    with metadata_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=True, indent=2)

    response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json={
            "file_reference": file_reference,
            "form_type": "jrsm",
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
        },
    )
    assert response.status_code == 410


def test_validate_response_is_text_only_contract(client: TestClient, auth_headers):
    file_reference = _upload_valid_workbook(client, auth_headers)
    response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json={
            "file_reference": file_reference,
            "form_type": "jrsm",
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
        },
    )
    assert response.status_code == 200
    body = response.json()

    assert "response_text_blocks" in body
    assert all(set(block.keys()) == {"title", "body"} for block in body["response_text_blocks"])
    assert "content_type" not in body
    assert "download_url" not in body
    assert "csv" not in body
    assert "file_bytes" not in body


def test_validate_repeat_calls_are_deterministic_except_run_id(client: TestClient, auth_headers):
    file_reference = _upload_valid_workbook(client, auth_headers)
    payload = {
        "file_reference": file_reference,
        "form_type": "jrsm",
        "country": "Rwanda",
        "year_for_request_of_medicine": 2026,
    }

    first = client.post("/validate/jrsm", headers=auth_headers, json=payload)
    second = client.post("/validate/jrsm", headers=auth_headers, json=payload)

    assert first.status_code == 200
    assert second.status_code == 200

    first_body = first.json()
    second_body = second.json()
    assert first_body["run_id"] != ""
    assert second_body["run_id"] != ""

    first_body_wo_run = {key: value for key, value in first_body.items() if key != "run_id"}
    second_body_wo_run = {key: value for key, value in second_body.items() if key != "run_id"}
    assert first_body_wo_run == second_body_wo_run


def test_validate_defaults_form_variation_to_generic(client: TestClient, auth_headers):
    file_reference = _upload_valid_workbook(client, auth_headers)
    response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json={
            "file_reference": file_reference,
            "form_type": "jrsm",
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
        },
    )

    assert response.status_code == 200


def test_validate_rejects_unsupported_form_type_for_jrsm_endpoint(client: TestClient, auth_headers):
    file_reference = _upload_valid_workbook(client, auth_headers)
    response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json={
            "file_reference": file_reference,
            "form_type": "jrf",
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
        },
    )

    assert response.status_code == 400
    assert response.json() == {
        "error_code": "UNSUPPORTED_FORM_TYPE",
        "message": "This endpoint does not support the requested form type.",
        "details": {
            "form_type": "jrf",
            "supported_form_types": ["jrsm"],
        },
        "correlation_id": response.headers["X-Request-ID"],
    }


def test_validate_rejects_removed_metadata_field(client: TestClient, auth_headers):
    file_reference = _upload_valid_workbook(client, auth_headers)
    response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json={
            "file_reference": file_reference,
            "form_type": "jrsm",
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
            "metadata": {"submitted_by": "tests"},
        },
    )

    body = response.json()
    assert response.status_code == 422
    assert body["error_code"] == "REQUEST_VALIDATION_ERROR"
    assert body["message"] == "The request body did not match the expected schema."
    assert any(error["loc"][-1] == "metadata" for error in body["details"]["errors"])


def test_validate_rejects_invalid_form_variation(client: TestClient, auth_headers):
    file_reference = _upload_valid_workbook(client, auth_headers)
    response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json={
            "file_reference": file_reference,
            "form_type": "jrsm",
            "form_variation": "custom",
            "country": "Rwanda",
            "year_for_request_of_medicine": 2026,
        },
    )

    body = response.json()
    assert response.status_code == 422
    assert body["error_code"] == "REQUEST_VALIDATION_ERROR"
    assert body["message"] == "The request body did not match the expected schema."
    assert any(error["loc"][-1] == "form_variation" for error in body["details"]["errors"])
