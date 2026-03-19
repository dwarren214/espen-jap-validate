"""Story 5 tests for Intro request consistency and required Intro fields."""

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
    intro_country: str = "Rwanda",
    intro_year: object = 2026,
    blank_intro_cells: set[str] | None = None,
    intro_year_number_format: str | None = None,
) -> bytes:
    workbook = Workbook()
    intro_sheet = workbook.active
    intro_sheet.title = "INTRO"

    for sheet_name in REQUIRED_SHEETS:
        if sheet_name == "INTRO":
            continue
        workbook.create_sheet(title=sheet_name)

    for cell_address in REQUIRED_INTRO_CELLS:
        intro_sheet[cell_address] = f"value_{cell_address}"
    intro_sheet["B2"] = EXPECTED_VERSION_MARKER
    intro_sheet["E33"] = intro_country
    intro_sheet["E35"] = intro_year
    if intro_year_number_format:
        intro_sheet["E35"].number_format = intro_year_number_format

    for cell_address in (blank_intro_cells or set()):
        intro_sheet[cell_address] = "   "

    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _upload_and_validate(
    client: TestClient,
    auth_headers: dict[str, str],
    workbook_bytes: bytes,
    *,
    request_country: str,
    request_year: int,
) -> dict:
    upload_response = client.post(
        "/upload",
        headers=auth_headers,
        files={"file": ("story5.xlsx", workbook_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert upload_response.status_code == 200
    file_reference = upload_response.json()["file_reference"]

    validate_response = client.post(
        "/validate/jrsm",
        headers=auth_headers,
        json={
            "file_reference": file_reference,
            "form_type": "jrsm",
            "country": request_country,
            "year_for_request_of_medicine": request_year,
        },
    )
    assert validate_response.status_code == 200
    return validate_response.json()


def test_country_mismatch_emits_error_at_intro_e33(client: TestClient, auth_headers):
    body = _upload_and_validate(
        client,
        auth_headers,
        _build_workbook_bytes(intro_country="Kenya", intro_year=2026),
        request_country="Rwanda",
        request_year=2026,
    )
    country_findings = [finding for finding in body["findings"] if finding["rule_id"] == "JRSM.INTRO.COUNTRY_MATCH"]
    assert country_findings
    assert any(finding["severity"] == "error" for finding in country_findings)
    assert any(finding.get("sheet") == "INTRO" and finding.get("cell") == "E33" for finding in country_findings)


def test_year_mismatch_emits_error_at_intro_e35(client: TestClient, auth_headers):
    body = _upload_and_validate(
        client,
        auth_headers,
        _build_workbook_bytes(intro_country="Rwanda", intro_year=2025),
        request_country="Rwanda",
        request_year=2026,
    )
    year_findings = [finding for finding in body["findings"] if finding["rule_id"] == "JRSM.INTRO.YEAR_MATCH"]
    assert year_findings
    assert any(finding["severity"] == "error" for finding in year_findings)
    assert any(finding.get("sheet") == "INTRO" and finding.get("cell") == "E35" for finding in year_findings)


def test_missing_required_intro_fields_emit_per_cell_findings(client: TestClient, auth_headers):
    body = _upload_and_validate(
        client,
        auth_headers,
        _build_workbook_bytes(
            intro_country="Rwanda",
            intro_year=2026,
            blank_intro_cells={"E37", "E49"},
        ),
        request_country="Rwanda",
        request_year=2026,
    )
    missing_field_findings = [
        finding for finding in body["findings"] if finding["rule_id"] == "JRSM.INTRO.REQUIRED_FIELDS_PRESENT"
    ]
    missing_cells = {finding.get("cell") for finding in missing_field_findings}
    assert {"E37", "E49"} <= missing_cells
    assert all(finding["severity"] == "error" for finding in missing_field_findings)


def test_intro_spacer_cells_are_not_required(client: TestClient, auth_headers):
    body = _upload_and_validate(
        client,
        auth_headers,
        _build_workbook_bytes(
            intro_country="Rwanda",
            intro_year=2026,
            blank_intro_cells={"E38", "E40", "E42", "E44"},
        ),
        request_country="Rwanda",
        request_year=2026,
    )
    missing_field_findings = [
        finding for finding in body["findings"] if finding["rule_id"] == "JRSM.INTRO.REQUIRED_FIELDS_PRESENT"
    ]
    assert not any(
        finding.get("cell") in {"E38", "E40", "E42", "E44"} for finding in missing_field_findings
    )


def test_country_normalization_handles_case_and_whitespace(client: TestClient, auth_headers):
    body = _upload_and_validate(
        client,
        auth_headers,
        _build_workbook_bytes(intro_country="  rWaNdA    republic  ", intro_year=2026),
        request_country="rwanda republic",
        request_year=2026,
    )
    country_findings = [finding for finding in body["findings"] if finding["rule_id"] == "JRSM.INTRO.COUNTRY_MATCH"]
    assert not country_findings


@pytest.mark.parametrize(
    ("intro_year", "year_format"),
    [
        (" 2026.0 ", None),
        (2026.0, "0"),
    ],
)
def test_year_normalization_handles_numeric_like_values(
    client: TestClient,
    auth_headers,
    intro_year: object,
    year_format: str | None,
):
    body = _upload_and_validate(
        client,
        auth_headers,
        _build_workbook_bytes(
            intro_country="Rwanda",
            intro_year=intro_year,
            intro_year_number_format=year_format,
        ),
        request_country="Rwanda",
        request_year=2026,
    )
    year_findings = [finding for finding in body["findings"] if finding["rule_id"] == "JRSM.INTRO.YEAR_MATCH"]
    assert not year_findings
