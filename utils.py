# utils.py
import sqlparse
from sqlparse.sql import IdentifierList, Identifier
from sqlparse.tokens import Keyword, DML

def quote_identifiers(query: str) -> str:
    """
    Parses the SQL query and automatically quotes table and column names.
    
    Note: This function assumes that identifiers (table and column names) do not include
    any SQL reserved keywords and are separated by spaces or commas.
    """
    parsed = sqlparse.parse(query)
    if not parsed:
        return query  # Return original if parsing fails

    stmt = parsed[0]
    quoted_query = ""

    # Flags to determine when to quote identifiers
    quote_next_identifier = False
    in_where = False

    for token in stmt.tokens:
        if token.ttype is DML and token.value.upper() == 'SELECT':
            quoted_query += token.value + " "
            quote_next_identifier = True
            continue

        if token.ttype is Keyword and token.value.upper() == 'FROM':
            quoted_query += token.value + " "
            quote_next_identifier = True
            continue

        if token.ttype is Keyword and token.value.upper() in ['WHERE', 'LIMIT', 'ORDER BY', 'GROUP BY']:
            quoted_query += token.value + " "
            quote_next_identifier = False
            continue

        if isinstance(token, IdentifierList):
            identifiers = []
            for identifier in token.get_identifiers():
                identifiers.append(f'"{identifier.get_real_name()}"')
            quoted_query += ", ".join(identifiers) + " "
            continue

        if isinstance(token, Identifier):
            quoted_query += f'"{token.get_real_name()}" '
            continue

        if token.ttype is Keyword:
            quoted_query += token.value + " "
            continue

        # Handle other tokens (e.g., literals, operators)
        quoted_query += token.value + " "

    return quoted_query.strip()