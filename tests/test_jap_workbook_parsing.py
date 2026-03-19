"""Story 4 tests for JRSM workbook parsing and template conformity checks."""

import io
import os
from pathlib import Path

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from openpyxl import Workbook

os.environ.setdefault("API_KEY", "test-key")

from espen_sql_api.jap_validation import storage  # noqa: E402
from espen_sql_api.routers import jap_validation  # noqa: E402

EXPECTED_VERSION_MARKER = "Joint request for selected PC medicines v.4.4"
CORRUPT_WORKBOOK_BYTES_PATH = Path("tests/fixtures/jrsm/corrupt_workbook.bin")
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
REQUIRED_INTRO_CELLS = [
    "B2",
    "E33",
    "E35",
    "E37",
    "E38",
    "E39",
    "E40",
    "E41",
    "E42",
    "E43",
    "E44",
    "E45",
    "E48",
    "E49",
    "E50",
    "E51",
]


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
    monkeypatch.setattr(jap_validation, "JAP_MAX_UPLOAD_BYTES", 1024 * 1024 * 10)

    app = FastAPI()
    app.include_router(jap_validation.router)
    return TestClient(app)


def _build_workbook_bytes(
    *,
    missing_sheets: set[str] | None = None,
    missing_intro_cells: set[str] | None = None,
    version_marker: str = EXPECTED_VERSION_MARKER,
) -> bytes:
    workbook = Workbook()
    intro_sheet = workbook.active
    intro_sheet.title = "INTRO"

    missing_sheets = missing_sheets or set()
    missing_intro_cells = missing_intro_cells or set()

    for sheet_name in REQUIRED_SHEETS:
        if sheet_name == "INTRO" or sheet_name in missing_sheets:
            continue
        workbook.create_sheet(title=sheet_name)

    for cell_address in REQUIRED_INTRO_CELLS:
        if cell_address in missing_intro_cells:
            continue
        if cell_address == "B2":
            intro_sheet[cell_address] = version_marker
        else:
            intro_sheet[cell_address] = f"value_{cell_address}"

    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _upload_and_validate(
    client: TestClient,
    auth_headers: dict[str, str],
    workbook_bytes: bytes,
    *,
    form_variation: str | None = None,
    findings_mode: str = "exceptions_only",
) -> dict:
    upload_response = client.post(
        "/upload",
        headers=auth_headers,
        files={"file": ("story4.xlsx", workbook_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert upload_response.status_code == 200
    file_reference = upload_response.json()["file_reference"]

    payload = {
        "file_reference": file_reference,
        "form_type": "jrsm",
        "country": "Rwanda",
        "year_for_request_of_medicine": 2026,
        "findings_mode": findings_mode,
    }
    if form_variation is not None:
        payload["form_variation"] = form_variation

    validate_response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json=payload,
    )
    assert validate_response.status_code == 200
    return validate_response.json()


def test_valid_workbook_parse_path_has_no_parse_failure(client: TestClient, auth_headers):
    workbook_bytes = _build_workbook_bytes()
    body = _upload_and_validate(client, auth_headers, workbook_bytes)
    assert not any(
        finding["rule_id"] == "JRSM.WORKBOOK.FILE_PARSE" and "WORKBOOK_PARSE_FAILED" in (finding.get("actual") or "")
        for finding in body["findings"]
    )


def test_corrupt_workbook_bytes_produce_parse_failure_finding(client: TestClient, auth_headers):
    body = _upload_and_validate(client, auth_headers, CORRUPT_WORKBOOK_BYTES_PATH.read_bytes())
    parse_findings = [finding for finding in body["findings"] if finding["rule_id"] == "JRSM.WORKBOOK.FILE_PARSE"]
    assert parse_findings
    assert any("WORKBOOK_PARSE_FAILED" in (finding.get("actual") or "") for finding in parse_findings)


def test_missing_required_sheet_produces_required_sheet_finding(client: TestClient, auth_headers):
    workbook_bytes = _build_workbook_bytes(missing_sheets={"IVM+"})
    body = _upload_and_validate(client, auth_headers, workbook_bytes)
    sheet_findings = [finding for finding in body["findings"] if finding["rule_id"] == "JRSM.WORKBOOK.REQUIRED_SHEETS"]
    assert sheet_findings
    assert any("IVM+" in finding["message"] for finding in sheet_findings)


def test_missing_required_cell_produces_required_cell_finding(client: TestClient, auth_headers):
    workbook_bytes = _build_workbook_bytes(missing_intro_cells={"E35"})
    body = _upload_and_validate(client, auth_headers, workbook_bytes)
    required_cell_findings = [finding for finding in body["findings"] if finding["rule_id"] == "JRSM.TEMPLATE.REQUIRED_CELLS"]
    assert required_cell_findings
    assert any("INTRO!E35" in (finding.get("actual") or "") for finding in required_cell_findings)


def test_version_marker_emits_info_or_warn_signal(client: TestClient, auth_headers):
    valid_body = _upload_and_validate(
        client,
        auth_headers,
        _build_workbook_bytes(version_marker=EXPECTED_VERSION_MARKER),
        findings_mode="full",
    )
    valid_marker_findings = [finding for finding in valid_body["findings"] if finding["rule_id"] == "JRSM.TEMPLATE.VERSION_MARKER"]
    assert valid_marker_findings
    assert any(finding["severity"] == "info" for finding in valid_marker_findings)

    invalid_body = _upload_and_validate(
        client,
        auth_headers,
        _build_workbook_bytes(version_marker="Unexpected marker value"),
        findings_mode="full",
    )
    invalid_marker_findings = [finding for finding in invalid_body["findings"] if finding["rule_id"] == "JRSM.TEMPLATE.VERSION_MARKER"]
    assert invalid_marker_findings
    assert any(finding["severity"] == "warn" for finding in invalid_marker_findings)
