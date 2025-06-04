import logging
import openai
import sqlglot

def quote_identifiers(query: str, is_postgres: bool = True) -> str:
    """
    Transpiles an SQL query to either PostgreSQL or T-SQL, quoting identifiers.
    Falls back to OpenAI API for tanspiling if sqlglot fails.

    Args:
        query (str): The SQL query to transpile.
        is_postgres (bool): True if the target dialect is PostgreSQL, False for T-SQL (MSSQL).

    Returns:
        str: The transpiled SQL query.
    """
    output_lang = "postgres"
    system_prompt_target = "PostgreSQL"

    if not is_postgres:
        output_lang = "tsql"
        system_prompt_target = "T-SQL for MSSQL Server"
        query = query.replace("CURDATE()", "GETDATE()") # Common function replacement for MSSQL

    try:
        # Attempt to transpile using sqlglot
        # The 'identify=True' argument in transpile helps in quoting identifiers
        # based on the target dialect.
        # sqlglot.transpile returns a list of strings, so we take the first element.
        transpiled_queries = sqlglot.transpile(query, read=None, write=output_lang, identify=True, pretty=False)
        if not transpiled_queries:
            raise ValueError("sqlglot.transpile returned an empty list.")
        result = transpiled_queries[0]
        logging.info("query_transpiled_sqlglot|input_query:%s|output_query:%s", query, result)
        return result
    except Exception as e:
        logging.warning("sqlglot_transpilation_failed|input_query:%s|error:%s|falling_back_to_openai", query, e)

        # Fallback to OpenAI API
        try:
            # Ensure the OpenAI API key is set in your environment variables (OPENAI_API_KEY)
            # If you are using a version of openai library >= 1.0.0
            client = openai.OpenAI() # Initializes the client using OPENAI_API_KEY from env

            system_message = f"Convert the supplied SQL query to {system_prompt_target}. Output nothing but the final valid SQL query. Do not use any formatting, markdown, code block, explanations or other extraneous outputs."

            response = client.chat.completions.create(
                model="gpt-4.1-nano",
                messages=[
                    {"role": "system", "content": system_message},
                    {"role": "user", "content": query}
                ],
                temperature=0, # For deterministic output
                max_tokens=2048 # Adjust as needed
            )

            if response.choices and response.choices[0].message and response.choices[0].message.content:
                openai_result = response.choices[0].message.content.strip()
                logging.info("query_transpiled_openai|input_query:%s|output_query:%s", query, openai_result)
                return openai_result
            else:
                logging.error("openai_call_failed|input_query:%s|reason:No content in response", query)
                # Return original query or raise an error if OpenAI response is not as expected
                return query # Or raise an exception

        except Exception as openai_error:
            logging.error("openai_api_call_failed|input_query:%s|error:%s", query, openai_error)
            # If OpenAI also fails, return the original query (or handle as appropriate)
            return query # Or raise openai_error