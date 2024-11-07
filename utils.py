import re

def quote_identifiers(query: str) -> str:
    """
    Quotes identifiers (table names and column names) in an SQL query string.
    Assumes identifiers do not contain whitespace or special characters,
    except for underscores. Ignores strings inside single quotes.
    """
    # Regular expression to match identifiers outside of single quotes
    identifier_pattern = re.compile(r'(?<!\')\b([A-Za-z_][A-Za-z0-9_]*)\b(?!\')')

    # List of SQL keywords/functions to ignore when quoting
    keywords = {
        "SELECT", "FROM", "WHERE", "AND", "OR", "AS", "COUNT", "DISTINCT", "JOIN",
        "ON", "IN", "GROUP", "BY", "ORDER", "LIMIT", "OFFSET", "IS", "NULL", "NOT",
        "BETWEEN", "LIKE", "HAVING", "CASE", "WHEN", "THEN", "ELSE", "END"
    }

    def replacer(match):
        word = match.group(0)
        # Quote the identifier if it is not a keyword
        if word.upper() not in keywords:
            return f'"{word}"'
        return word

    # Replace identifiers in the query string
    quoted_query = identifier_pattern.sub(replacer, query)
    return quoted_query