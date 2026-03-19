FROM python:3.12-slim-bookworm

ENV ACCEPT_EULA=Y
RUN apt-get update && apt-get install -y --no-install-recommends curl gnupg


RUN --mount=target=/var/lib/apt/lists,type=cache,sharing=locked \
    --mount=target=/var/cache/apt,type=cache,sharing=locked \
    rm -f /etc/apt/apt.conf.d/docker-clean \
    && curl -fsSL https://packages.microsoft.com/keys/microsoft.asc | gpg --dearmor -o /usr/share/keyrings/microsoft-prod.gpg \
    && curl https://packages.microsoft.com/config/debian/12/prod.list > /etc/apt/sources.list.d/microsoft-prod.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends \
        unixodbc-dev \
        msodbcsql18 \
    && apt-get purge -y --auto-remove -o APT::AutoRemove::RecommendsImportant=false \
    && rm -rf /var/lib/apt/lists/*

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
        --no-dev \
        --no-install-project


ENV PYTHONUNBUFFERED=1
ENV DEBUG=0

RUN addgroup --system --gid 1001 fastapi \
    && adduser --system --uid 1001 --ingroup fastapi fastapi

WORKDIR /code
COPY --chown=fastapi:fastapi src/espen_sql_api /code/espen_sql_api
COPY --chown=fastapi:fastapi docs/jrsm_summary_shipment_validation_reference.md /code/docs/
COPY --chown=fastapi:fastapi espen.db /code
COPY --chown=fastapi:fastapi scripts /code/scripts

USER fastapi

# make sure we use the virtualenv python/uvuicorn by default
ENV PATH="/code/.venv/bin:$PATH"

CMD ["uvicorn", "espen_sql_api.main:app", "--host=0.0.0.0", "--port=8000"]
