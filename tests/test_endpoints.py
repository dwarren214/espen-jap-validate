"""Integration tests for API endpoints against deployed API."""

import os

import httpx
import pytest
from dotenv import load_dotenv

load_dotenv()

RUN_DEPLOYED_ENDPOINT_TESTS = os.environ.get("RUN_DEPLOYED_ENDPOINT_TESTS") == "1"
BASE_URL = os.environ.get("TEST_API_URL", "https://espen-sql-api.openchatstudio.com")
api_key = os.environ.get("API_KEY")
AUTH_HEADER = {"Authorization": f"Bearer {api_key}"}

pytestmark = pytest.mark.skipif(
    not RUN_DEPLOYED_ENDPOINT_TESTS,
    reason="Deployed endpoint smoke tests are opt-in. Set RUN_DEPLOYED_ENDPOINT_TESTS=1 to run them.",
)


@pytest.fixture(scope="module")
def client():
    """HTTP client for API requests."""
    with httpx.Client(base_url=BASE_URL, timeout=30.0) as c:
        yield c


class TestHealthEndpoints:
    """Test health and status endpoints."""

    def test_up(self, client):
        response = client.get("/up")
        assert response.status_code == 200
        assert response.json() == {"status": "up"}

    def test_health_returns_response(self, client):
        """Health endpoint returns a response."""
        response = client.get("/health")
        assert response.status_code == 200
        data = response.json()
        assert "status" in data
        assert "databases" in data


class TestAuthRequired:
    """Test that endpoints require authentication."""

    def test_execute_sql_query_requires_auth(self, client):
        response = client.post("/execute_sql_query", json={"query": "SELECT 1"})
        assert response.status_code == 403

    def test_fetch_column_names_requires_auth(self, client):
        response = client.post("/fetch_column_names_and_types", json={"tables": ["test"]})
        assert response.status_code == 403

    def test_oncho_execute_query_requires_auth(self, client):
        response = client.post("/oncho/execute_query", json={"query": "SELECT 1"})
        assert response.status_code == 403

    def test_campaign_hub_requires_auth(self, client):
        response = client.post("/query_campaign_hub_data", json={"query": "SELECT 1"})
        assert response.status_code == 403


class TestSQLiteEndpoints:
    """Test endpoints that use local SQLite (espen.db)."""

    def test_fetch_column_names_valid_table(self, client):
        """Fetch columns for a known table in espen.db."""
        response = client.post(
            "/fetch_column_names_and_types",
            json={"tables": ["Afro_Admin0"]},
            headers=AUTH_HEADER,
        )
        assert response.status_code == 200
        data = response.json()
        assert "Afro_Admin0" in data
        assert isinstance(data["Afro_Admin0"], dict)

    def test_fetch_column_names_unknown_table(self, client):
        """Unknown table returns 'Table not found'."""
        response = client.post(
            "/fetch_column_names_and_types",
            json={"tables": ["NonExistentTable"]},
            headers=AUTH_HEADER,
        )
        assert response.status_code == 200
        data = response.json()
        assert data["NonExistentTable"] == "Table not found"

    def test_fetch_column_names_multiple_tables(self, client):
        """Can fetch multiple tables at once."""
        response = client.post(
            "/fetch_column_names_and_types",
            json={"tables": ["Afro_Admin0", "Afro_Country_LF_data"]},
            headers=AUTH_HEADER,
        )
        assert response.status_code == 200
        data = response.json()
        assert "Afro_Admin0" in data
        assert "Afro_Country_LF_data" in data


class TestMSSQLEndpoints:
    """Test MSSQL endpoints."""

    def test_execute_sql_query(self, client):
        response = client.post(
            "/execute_sql_query",
            json={"query": "SELECT TOP 1 * FROM Afro_Admin0"},
            headers=AUTH_HEADER,
        )
        assert response.status_code == 200
        data = response.json()
        assert len(data) == 1


class TestPostgreSQLEndpoints:
    """Test PostgreSQL endpoints."""

    def test_oncho_execute_query(self, client):
        response = client.post(
            "/oncho/execute_query",
            json={"query": "SELECT 1 AS value"},
            headers=AUTH_HEADER,
        )
        assert response.status_code == 200
        data = response.json()
        assert len(data) == 1


class TestCampaignEndpoints:
    """Test Campaign Hub endpoints."""

    def test_fetch_campaign_hub_columns(self, client):
        """Fetch campaign hub columns."""
        response = client.get(
            "/fetch_campaign_hub_columns",
            headers=AUTH_HEADER,
        )
        assert response.status_code == 200
        data = response.json()
        assert isinstance(data, list)
        assert "campaign_id" in data

    def test_query_campaign_hub_data_invalid_query(self, client):
        """Invalid SQL returns error."""
        response = client.post(
            "/query_campaign_hub_data",
            json={"query": "INVALID SQL"},
            headers=AUTH_HEADER,
        )
        assert response.status_code == 400
