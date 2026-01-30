"""Pydantic models for request/response bodies."""

from typing import List, Optional

from pydantic import BaseModel


class SQLQuery(BaseModel):
    query: str
    force_json: Optional[bool] = False
    force_csv: Optional[bool] = False


class TableNames(BaseModel):
    tables: List[str]
