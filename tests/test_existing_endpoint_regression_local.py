"""Local non-regression tests for pre-existing ESPEN, Oncho, and Campaign endpoints."""

from __future__ import annotations

import json
import os
from contextlib import contextmanager

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

os.environ.setdefault("API_KEY", "test-key")
os.environ.setdefault("DATABASE_URL", "postgresql+psycopg2://user:pass@localhost:5432/testdb")
os.environ.setdefault(
    "REMOTE_DATABASE_URL",
    "postgresql+psycopg2://user:pass@localhost:5432/testdb",
)

from espen_sql_api.routers import campaign, espen, oncho  # noqa: E402


class _FakeQueryResult:
    def __init__(self, rows, columns):
        self._rows = rows
        self._columns = columns

    def fetchall(self):
        return self._rows

    def fetchone(self):
        return self._rows[0] if self._rows else None

    def keys(self):
        return self._columns

    def mappings(self):
        return self

    def all(self):
        return [dict(zip(self._columns, row)) for row in self._rows]


class _FakeSession:
    def __init__(self, execute_results=None):
        self.execute_results = list(execute_results or [])
        self.executed = []
        self.closed = False
        self.rolled_back = False
        self.committed = False

    def execute(self, statement, params=None):
        self.executed.append((statement, params))
        if not self.execute_results:
            raise AssertionError("No fake execute result configured.")
        result = self.execute_results.pop(0)
        if isinstance(result, Exception):
            raise result
        return result

    def rollback(self):
        self.rolled_back = True

    def close(self):
        self.closed = True

    def commit(self):
        self.committed = True


@pytest.fixture
def auth_headers():
    return {"Authorization": "Bearer test-key"}


@pytest.fixture
def client():
    app = FastAPI()
    app.include_router(espen.router)
    app.include_router(oncho.router)
    app.include_router(campaign.router)
    return TestClient(app)


def test_existing_endpoints_require_auth(client: TestClient):
    assert client.post("/execute_sql_query", json={"query": "SELECT 1"}).status_code == 403
    assert client.post("/fetch_column_names_and_types", json={"tables": ["Afro_Admin0"]}).status_code == 403
    assert client.post("/oncho/execute_query", json={"query": "SELECT 1"}).status_code == 403
    assert client.post("/oncho/top3", json={}).status_code == 403
    assert client.post("/query_campaign_hub_data", json={"query": "SELECT 1"}).status_code == 403
    assert client.get("/fetch_campaign_hub_columns").status_code == 403
    assert client.get("/fetch_campaign_hub_data").status_code == 403


def test_fetch_column_names_and_types_returns_known_and_unknown_tables(
    client: TestClient,
    auth_headers,
    monkeypatch: pytest.MonkeyPatch,
):
    fields_payload = json.dumps(
        [
            {"Name": "ADM0_NAME", "Description": "Country name"},
            {"Name": "ISO3", "Description": "Country ISO3"},
        ]
    )
    fake_meta_session = _FakeSession(
        execute_results=[
            _FakeQueryResult([(fields_payload,)], ["Fields"]),
            _FakeQueryResult([], ["Fields"]),
        ]
    )
    monkeypatch.setattr(espen, "MetaSessionLocal", lambda: fake_meta_session)

    response = client.post(
        "/fetch_column_names_and_types",
        headers=auth_headers,
        json={"tables": ["Afro_Admin0", "MissingTable"]},
    )

    assert response.status_code == 200
    assert response.json() == {
        "Afro_Admin0": {"ADM0_NAME": "Country name", "ISO3": "Country ISO3"},
        "MissingTable": "Table not found",
    }
    assert fake_meta_session.closed is True


def test_execute_sql_query_returns_json_rows_without_changing_contract(
    client: TestClient,
    auth_headers,
    monkeypatch: pytest.MonkeyPatch,
):
    fake_session = _FakeSession()
    fake_result = _FakeQueryResult([(1, "Rwanda")], ["id", "name"])
    monkeypatch.setattr(espen, "RemoteSessionLocal", lambda: fake_session)
    monkeypatch.setattr(espen, "validate_query_safety", lambda query: (True, ""))
    monkeypatch.setattr(espen, "quote_identifiers", lambda query, is_postgres=False: query)
    monkeypatch.setattr(espen, "execute_query_with_retry", lambda session, query: fake_result)

    response = client.post(
        "/execute_sql_query",
        headers=auth_headers,
        json={"query": "SELECT id, name FROM sample_table"},
    )

    assert response.status_code == 200
    assert response.json() == [{"id": 1, "name": "Rwanda"}]
    assert fake_session.closed is True


def test_oncho_endpoints_preserve_execute_and_top3_contracts(
    client: TestClient,
    auth_headers,
    monkeypatch: pytest.MonkeyPatch,
):
    fake_query_session = _FakeSession()
    fake_result = _FakeQueryResult([(1,)], ["value"])
    monkeypatch.setattr(oncho, "SessionLocal", lambda: fake_query_session)
    monkeypatch.setattr(oncho, "validate_query_safety", lambda query: (True, ""))
    monkeypatch.setattr(oncho, "quote_identifiers", lambda query, is_postgres=True: query)
    monkeypatch.setattr(oncho, "execute_query_with_retry", lambda session, query: fake_result)

    execute_response = client.post(
        "/oncho/execute_query",
        headers=auth_headers,
        json={"query": "SELECT 1 AS value"},
    )

    assert execute_response.status_code == 200
    assert execute_response.json() == [{"value": 1}]
    assert fake_query_session.closed is True

    fake_top3_session = _FakeSession()
    monkeypatch.setattr(oncho, "SessionLocal", lambda: fake_top3_session)
    monkeypatch.setattr(
        oncho,
        "run_top3_query",
        lambda session, payload: [
            oncho.ScenarioRank(
                scenario_label="Scenario A",
                success_rate_pct=91.5,
                median_years=3.0,
                total_cost=1200.0,
                min_remaining_p50=0.0,
                max_remaining_p50=0.3,
            )
        ],
    )

    top3_response = client.post(
        "/oncho/top3",
        headers=auth_headers,
        json={
            "location_type": "country",
            "location_values": ["Rwanda"],
            "start_year": 2025,
            "end_year": 2028,
            "threshold": 0.01,
        },
    )

    assert top3_response.status_code == 200
    assert top3_response.json() == [
        {
            "scenario_label": "Scenario A",
            "success_rate_pct": 91.5,
            "median_years": 3.0,
            "total_cost": 1200.0,
            "min_remaining_p50": 0.0,
            "max_remaining_p50": 0.3,
        }
    ]
    assert fake_top3_session.closed is True


def test_campaign_endpoints_preserve_cached_query_and_fetch_contracts(
    client: TestClient,
    auth_headers,
    monkeypatch: pytest.MonkeyPatch,
):
    monkeypatch.setattr(campaign, "validate_query_safety", lambda query: (True, ""))
    monkeypatch.setattr(campaign, "ensure_campaigns_db_exists", lambda: None)

    @contextmanager
    def fake_campaign_session(date=None):
        yield _FakeSession(
            execute_results=[
                _FakeQueryResult([(101, "AFRO", 2026)], ["campaign_id", "who_region", "campaign_start_year"]),
            ]
        )

    monkeypatch.setattr(campaign, "campaign_hub_session", fake_campaign_session)

    query_response = client.post(
        "/query_campaign_hub_data",
        headers=auth_headers,
        json={"query": "SELECT campaign_id, who_region, campaign_start_year FROM campaigns"},
    )

    assert query_response.status_code == 200
    assert query_response.json() == {
        "data": [{"campaign_id": 101, "who_region": "AFRO", "campaign_start_year": 2026}],
        "row_count": 1,
    }

    @contextmanager
    def fake_column_session(date=None):
        yield _FakeSession(
            execute_results=[
                _FakeQueryResult(
                    [(0, "campaign_id", "INTEGER"), (1, "who_region", "TEXT")],
                    ["cid", "name", "type"],
                )
            ]
        )

    monkeypatch.setattr(campaign, "campaign_hub_session", fake_column_session)
    columns_response = client.get("/fetch_campaign_hub_columns", headers=auth_headers)
    assert columns_response.status_code == 200
    assert columns_response.json() == ["campaign_id", "who_region"]

    class _FakeHTTPResponse:
        status_code = 200

        @staticmethod
        def json():
            return [
                {"WHO Region": "AFRO", "Campaign Start Year": 2026, "campaign_id": 1},
                {"WHO Region": "EMRO", "Campaign Start Year": 2026, "campaign_id": 2},
                {"WHO Region": "AFRO", "Campaign Start Year": 2023, "campaign_id": 3},
            ]

    monkeypatch.setattr(campaign.httpx, "get", lambda url, headers: _FakeHTTPResponse())
    fetch_response = client.get("/fetch_campaign_hub_data", headers=auth_headers)

    assert fetch_response.status_code == 200
    assert fetch_response.json() == [
        {
            "WHO Region": "AFRO",
            "Campaign Start Year": 2026,
            "campaign_id": 1,
            "Diseases Targeted": "unspecified",
        }
    ]
