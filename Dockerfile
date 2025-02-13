FROM python:3.12-slim-bookworm

RUN --mount=target=/var/lib/apt/lists,type=cache,sharing=locked \
    --mount=target=/var/cache/apt,type=cache,sharing=locked \
    rm -f /etc/apt/apt.conf.d/docker-clean && \
    apt-get update \
    && apt-get install -y \
    # psycopg2 dependencies
    libpq-dev \
    # psql client for dbshell
    postgresql-client \
    # cleaning up unused files
    && apt-get purge -y --auto-remove -o APT::AutoRemove::RecommendsImportant=false

# This approximately follows this guide: https://hynek.me/articles/docker-uv/
# Which creates a standalone environment with the dependencies.
# - Silence uv complaining about not being able to use hard links,
# - tell uv to byte-compile packages for faster application startups,
# - prevent uv from accidentally downloading isolated Python builds,
# - pick a Python (use `/usr/bin/python3.12` on uv 0.5.0 and later),
# - and finally declare `/app` as the target for `uv sync`.
ENV UV_LINK_MODE=copy \
    UV_COMPILE_BYTECODE=1 \
    UV_PYTHON_DOWNLOADS=never \
    UV_PROJECT_ENVIRONMENT=/code/.venv

COPY --from=ghcr.io/astral-sh/uv:0.5.31 /uv /uvx /bin/

# Synchronize dependencies.
# This layer is cached until uv.lock or pyproject.toml change.
RUN --mount=type=cache,target=/root/.cache \
    --mount=type=bind,source=uv.lock,target=uv.lock \
    --mount=type=bind,source=pyproject.toml,target=pyproject.toml \
    uv sync -v \
        --locked \
        --no-dev


ENV PYTHONUNBUFFERED=1
ENV DEBUG=0

RUN addgroup --system fastapi \
    && adduser --system --ingroup fastapi fastapi

WORKDIR /code
COPY --chown=fastapi:fastapi main.py utils.py /code

USER fastapi

# make sure we use the virtualenv python/uvuicorn by default
ENV PATH="/code/.venv/bin:$PATH"

CMD ["uvicorn", "main:app", "--host=0.0.0.0", "--port=8000"]
