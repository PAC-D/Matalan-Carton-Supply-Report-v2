#!/bin/sh
set -eu

cd "$(dirname "$0")"

# Local development uses PostgreSQL in Docker:
#   docker run -d --name carton-pg -e POSTGRES_PASSWORD=postgres -p 5432:5432 postgres:18
DATABASE_URL=${DATABASE_URL:-postgresql://postgres:postgres@localhost:5432/postgres}
export DATABASE_URL

if command -v python3 >/dev/null 2>&1; then
    PYEXE=$(command -v python3)
elif command -v python >/dev/null 2>&1; then
    PYEXE=$(command -v python)
else
    printf '%s\n' '[ERROR] Python 3.14 or newer was not found on PATH.' >&2
    exit 1
fi

if [ ! -x .venv/bin/python ]; then
    printf '%s\n' '[setup] Creating virtual environment...'
    "$PYEXE" -m venv .venv || {
        printf '%s\n' '[ERROR] Could not create the virtual environment.' >&2
        exit 1
    }
fi

VENV_PY=.venv/bin/python

printf '%s\n' '[setup] Installing project requirements...'
"$VENV_PY" -m pip install --quiet --upgrade pip || {
    printf '%s\n' '[ERROR] Could not upgrade pip.' >&2
    exit 1
}
"$VENV_PY" -m pip install --quiet -r requirements.txt || {
    printf '%s\n' '[ERROR] Setup failed. Check Python and your internet connection.' >&2
    exit 1
}

printf '\nMatalan Carton Supply Report\nhttp://localhost:8501\nDatabase: local PostgreSQL (start it with: docker start carton-pg)\nPress Ctrl+C to stop.\n\n'
exec "$VENV_PY" -m streamlit run app.py --server.address=localhost --server.port=8501