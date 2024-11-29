from fastapi import FastAPI, HTTPException, Depends
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Dict, Optional
import os
import json
import io
import csv
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker
from utils import quote_identifiers  # Import the helper function

app = FastAPI()

# Database configuration
DATABASE_URL = os.getenv('DATABASE_URL')

# Adjust for Heroku Postgres URL scheme
if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

# Create SQLAlchemy engine and session
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)

# API Key configuration
API_KEY = os.getenv('API_KEY')  # Set this in your environment variables
if not API_KEY:
    raise ValueError("API_KEY environment variable is not set.")

security = HTTPBearer()

# Dependency for API Key authentication
def api_key_auth(credentials: HTTPAuthorizationCredentials = Depends(security)):
    if credentials.scheme != "Bearer":
        raise HTTPException(status_code=403, detail="Invalid authentication scheme.")
    if credentials.credentials != API_KEY:
        raise HTTPException(status_code=403, detail="Invalid or missing API key.")

# Constants for response format thresholds
MAX_JSON_ROWS = 100  # Maximum number of rows for JSON response
MAX_JSON_CELLS = 1000  # Maximum total cells (rows * columns) for JSON response

# Pydantic models for request bodies
class SQLQuery(BaseModel):
    query: str
    force_json: Optional[bool] = False  # Optional parameter to force JSON response
    force_csv: Optional[bool] = False  # Optional parameter to force CSV response

class TableNames(BaseModel):
    tables: List[str]

# Endpoint to execute SQL queries
@app.post("/execute_sql_query", dependencies=[Depends(api_key_auth)])
def execute_sql_query(query_data: SQLQuery):
    session = SessionLocal()
    try:
        # Automatically quote identifiers in the query
        quoted_query = quote_identifiers(query_data.query)
        
        # Execute the quoted query with parameters
        result = session.execute(text(quoted_query))
        rows = result.fetchall()
        column_names = list(result.keys())
        
        # Calculate result size
        num_rows = len(rows)
        num_columns = len(column_names)
        total_cells = num_rows * num_columns

        # Determine if we should return CSV based on size thresholds
        should_return_csv = query_data.force_csv or (
            not query_data.force_json and  # Not forcing JSON
            (num_rows > MAX_JSON_ROWS or total_cells > MAX_JSON_CELLS)
        )

        if should_return_csv:
            # Create a string buffer for CSV data
            output = io.StringIO()
            writer = csv.writer(output)
            
            # Write header
            writer.writerow(column_names)
            
            # Write data rows
            for row in rows:
                writer.writerow(row)
            
            # Prepare the response
            output.seek(0)
            return StreamingResponse(
                iter([output.getvalue()]),
                media_type="text/csv",
                headers={
                    "Content-Disposition": "attachment; filename=query_results.csv"
                }
            )
        else:
            # Convert to list of dictionaries for JSON response
            result_list = [dict(zip(column_names, row)) for row in rows]
            return result_list
            
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        session.close()

# Endpoint to fetch column names and types
@app.post("/fetch_column_names_and_types", dependencies=[Depends(api_key_auth)])
def fetch_column_names_and_types(table_data: TableNames):
    session = SessionLocal()
    results = {}
    try:
        for table_name in table_data.tables:
            # Quote the table name
            quoted_table = f'"{table_name}"'
            
            # Query to get fields and descriptions
            query = text(f"""
                SELECT "Fields"
                FROM espen_tables
                WHERE LOWER("Name_Analytical_Table") = :table_name
            """)
            res = session.execute(query, {'table_name': table_name.lower()}).fetchone()
            if res:
                fields = res[0]
                results[table_name] = {field['Name']: field['Description'] for field in fields}
            else:
                results[table_name] = "Table not found"
        return results
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    finally:
        session.close()
