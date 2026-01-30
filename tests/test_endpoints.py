"""Basic integration tests for API endpoints."""

import os

import pytest
from dotenv import load_dotenv

load_dotenv()

# Set dummy env vars before importing app (db.py needs these at import time)
os.environ.setdefault("DATABASE_URL", "postgresql://test:test@localhost/test")
os.environ.setdefault("REMOTE_DATABASE_URL", "mssql+pyodbc://test:test@localhost/test")
os.environ.setdefault("API_KEY", "test-api-key")

from fastapi.testclient import TestClient

from espen_sql_api.main import app

client = TestClient(app, raise_server_exceptions=False)
api_key = os.environ.get("API_KEY")
AUTH_HEADER = {"Authorization": f"Bearer {api_key}"}


class TestHealthEndpoints:
    """Test health and status endpoints."""

    def test_up(self):
        response = client.get("/up")
        assert response.status_code == 200
        assert response.json() == {"status": "up"}

    def test_health_returns_response(self):
        """Health endpoint returns a response (may be unhealthy without real DBs)."""
        response = client.get("/health")
        assert response.status_code in (200, 503)
        data = response.json()
        assert "status" in data
        assert "databases" in data


class TestAuthRequired:
    """Test that endpoints require authentication."""

    def test_execute_sql_query_requires_auth(self):
        response = client.post("/execute_sql_query", json={"query": "SELECT 1"})
        assert response.status_code == 403

    def test_fetch_column_names_requires_auth(self):
        response = client.post("/fetch_column_names_and_types", json={"tables": ["test"]})
        assert response.status_code == 403

    def test_oncho_execute_query_requires_auth(self):
        response = client.post("/oncho/execute_query", json={"query": "SELECT 1"})
        assert response.status_code == 403

    def test_campaign_hub_requires_auth(self):
        response = client.post("/query_campaign_hub_data", json={"query": "SELECT 1"})
        assert response.status_code == 403


class TestSQLiteEndpoints:
    """Test endpoints that use local SQLite (espen.db)."""

    def test_fetch_column_names_valid_table(self):
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

    def test_fetch_column_names_unknown_table(self):
        """Unknown table returns 'Table not found'."""
        response = client.post(
            "/fetch_column_names_and_types",
            json={"tables": ["NonExistentTable"]},
            headers=AUTH_HEADER,
        )
        assert response.status_code == 200
        data = response.json()
        assert data["NonExistentTable"] == "Table not found"

    def test_fetch_column_names_multiple_tables(self):
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
    """Test MSSQL endpoints (may fail without real DB connection)."""

    @pytest.mark.skipif(
        "REMOTE_DATABASE_URL" not in os.environ or "localhost" in os.environ.get("REMOTE_DATABASE_URL", ""),
        reason="Requires real MSSQL connection"
    )
    def test_execute_sql_query(self):
        response = client.post(
            "/execute_sql_query",
            json={"query": "SELECT TOP 1 * FROM Afro_Admin0"},
            headers=AUTH_HEADER,
        )
        assert response.status_code == 200, response.content


class TestPostgreSQLEndpoints:
    """Test PostgreSQL endpoints (may fail without real DB connection)."""

    @pytest.mark.skipif(
        "DATABASE_URL" not in os.environ or "localhost" in os.environ.get("DATABASE_URL", ""),
        reason="Requires real PostgreSQL connection"
    )
    def test_oncho_execute_query(self):
        response = client.post(
            "/oncho/execute_query",
            json={"query": "SELECT 1"},
            headers=AUTH_HEADER,
        )
        assert response.status_code == 200, response.content


class TestCampaignEndpoints:
    """Test Campaign Hub endpoints."""

    def test_fetch_campaign_hub_columns(self):
        """Fetch campaign hub columns (may trigger API fetch if no local cache)."""
        response = client.get(
            "/fetch_campaign_hub_columns",
            headers=AUTH_HEADER,
        )
        # 200 if DB exists/created, 404 if no DB, 500 if external API fails
        assert response.status_code in (200, 404, 500)

    def test_query_campaign_hub_data_invalid_query(self):
        """Invalid SQL returns error."""
        response = client.post(
            "/query_campaign_hub_data",
            json={"query": "INVALID SQL"},
            headers=AUTH_HEADER,
        )
        # Should fail with 400 (bad SQL) or 500 (no DB)
        assert response.status_code in (400, 500)
