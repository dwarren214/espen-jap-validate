#!/usr/bin/env python3
"""Run test suite inside Docker container (no local MSSQL driver needed).

Usage:
    ./scripts/run_tests_docker.py
    ./scripts/run_tests_docker.py -v
    ./scripts/run_tests_docker.py -k test_health
"""

import subprocess
import sys
from pathlib import Path


def main():
    project_root = Path(__file__).parent.parent
    pytest_args = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else "-v"

    # Build image with dev dependencies
    print("Building Docker image with test dependencies...")
    subprocess.run(
        ["docker", "build", "-t", "espen-sql-api:test", "."],
        cwd=project_root,
        check=True,
    )

    # Run pytest in container
    # Use venv pip/python, run as root to install test deps
    print(f"\nRunning: pytest tests/ {pytest_args}\n")
    pytest_args_list = sys.argv[1:] if len(sys.argv) > 1 else ["-v"]

    # Load env vars from .env file if it exists
    env_file = project_root / ".env"
    env_args = []
    if env_file.exists():
        env_args = ["--env-file", str(env_file)]
    else:
        # Fallback to dummy values for basic tests
        env_args = [
            "-e", "DATABASE_URL=postgresql://test:test@localhost/test",
            "-e", "REMOTE_DATABASE_URL=mssql+pyodbc://test:test@localhost/test",
            "-e", "API_KEY=test-api-key",
        ]

    result = subprocess.run(
        [
            "docker", "run", "--rm",
            "--user", "root",
            *env_args,
            "-v", f"{project_root}/tests:/code/tests:ro",
            "-v", f"{project_root}/espen.db:/code/espen.db:ro",
            "--entrypoint", "sh",
            "espen-sql-api:test",
            "-c", f"uv pip install -q pytest pytest-asyncio && /code/.venv/bin/pytest tests/ {' '.join(pytest_args_list)}",
        ],
    )
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
