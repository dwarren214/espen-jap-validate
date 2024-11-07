import sqlparse
from sqlparse.sql import IdentifierList, Identifier, Function, TokenList
from sqlparse.tokens import Keyword, DML, Whitespace, Punctuation

def is_subselect(parsed):
    if not parsed.is_group:
        return False
    for item in parsed.tokens:
        if item.ttype is DML and item.value.upper() == 'SELECT':
            return True
    return False

def quote_identifiers(query: str) -> str:
    """
    Parses the SQL query and automatically quotes table and column names,
    excluding SQL keywords and functions.
    """
    parsed = sqlparse.parse(query)
    if not parsed:
        return query  # Return original if parsing fails

    stmt = parsed[0]
    quoted_query = ""

    # Iterate through the tokens recursively
    def process_tokens(tokens: TokenList):
        nonlocal quoted_query
        for token in tokens:
            if token.is_group:
                # Recursively process subgroups
                process_tokens(token)
            elif isinstance(token, Function):
                # Do not quote function names
                quoted_query += token.value
            elif isinstance(token, IdentifierList):
                identifiers = []
                for identifier in token.get_identifiers():
                    identifiers.append(quote_identifier(identifier))
                quoted_query += ", ".join(identifiers) + " "
            elif isinstance(token, Identifier):
                quoted_query += quote_identifier(token) + " "
            elif token.ttype is Keyword or token.ttype is DML:
                quoted_query += token.value.upper() + " "
            elif token.ttype is Whitespace:
                quoted_query += " "
            elif token.ttype is Punctuation:
                quoted_query += token.value
            else:
                quoted_query += token.value + " "

    def quote_identifier(identifier: Identifier) -> str:
        # If the identifier is a function, do not quote
        if isinstance(identifier, Function):
            return identifier.value
        # Otherwise, quote the real name
        return f'"{identifier.get_real_name()}"'

    process_tokens(stmt.tokens)
    return ' '.join(quoted_query.split())

