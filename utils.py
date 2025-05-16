import logging

import sqlglot
from datetime import datetime


def quote_identifiers(query: str, is_postgres: bool = True) -> str:
    output_lang = "postgres" if is_postgres else "tsql"
    result = sqlglot.transpile(query, identify=True, write=output_lang)[0]
    logging.info("query_transpiled|%s|%s", query, result)
    return result
