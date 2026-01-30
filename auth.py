"""API authentication."""

import os

from dotenv import load_dotenv

load_dotenv()

from fastapi import Depends, HTTPException  # noqa: E402
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer  # noqa: E402

API_KEY = os.getenv("API_KEY")
if not API_KEY:
    raise ValueError("API_KEY environment variable is not set.")

security = HTTPBearer()


def api_key_auth(credentials: HTTPAuthorizationCredentials = Depends(security)):
    """Dependency for API Key authentication."""
    if credentials.scheme != "Bearer":
        raise HTTPException(status_code=403, detail="Invalid authentication scheme.")
    if credentials.credentials != API_KEY:
        raise HTTPException(status_code=403, detail="Invalid or missing API key.")
