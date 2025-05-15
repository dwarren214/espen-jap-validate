import sqlglot
from datetime import datetime


def quote_identifiers(query: str, is_postgres: bool = True) -> str:
    output_lang = "postgres" if is_postgres else "tsql"
    result = sqlglot.transpile(query, identify=True, write=output_lang)[0]
    
    # Log the query and result
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("query_logs.log", "a") as log_file:
        log_file.write(f"[{timestamp}] {query}\n")
        log_file.write(f"Postgres: [{is_postgres}]\n")
        log_file.write(f"[{timestamp}] {result}\n\n")
    
    return result