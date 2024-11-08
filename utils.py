import sqlglot

def quote_identifiers(query: str) -> str:
    return sqlglot.transpile(query, identify=True, write='postgres')[0]
