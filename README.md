# HTTP SQL API

A small API for querying the database.

This is used in conjunction with Open Chat Studio [custom actions](https://dimagi.github.io/open-chat-studio-docs/concepts/custom_actions/) to allow a bot to query the DB via an
HTTP API.

## Prerequisites

This project requires access to a PostgreSQL database.

## Setup

1. Install UV

   See https://docs.astral.sh/uv/getting-started/installation/

2. Install the dependencies

   ```shell
   uv sync   
   ```

3. Create a `.env` file with the following content:

   ```shell
    DATABASE_URL=postgresql://user:password@host:port/dbname
    API_KEY=random-key
    ```

4. Run the app

    ```shell
    uv run uvicorn main:app --reload
    ```
