import sqlglot


def quote_identifiers(query: str, is_postgres: bool = True) -> str:
    output_lang = "postgres" if is_postgres else "tsql"
    return sqlglot.transpile(query, identify=True, write=output_lang)[0]