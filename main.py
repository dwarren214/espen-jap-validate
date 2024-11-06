from fastapi import FastAPI, HTTPException, Depends
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel
from typing import List, Dict
import os
import json
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

# Pydantic models for request bodies
class SQLQuery(BaseModel):
    query: str

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
        column_names = result.keys()
        
        # Convert to list of dictionaries
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
