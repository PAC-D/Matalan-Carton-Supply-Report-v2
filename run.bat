@echo off
setlocal
cd /d "%~dp0"

rem Local development uses PostgreSQL in Docker:
rem   docker run -d --name carton-pg -e POSTGRES_PASSWORD=postgres -p 5432:5432 postgres:18
if not defined DATABASE_URL set "DATABASE_URL=postgresql://postgres:postgres@localhost:5432/postgres"

set "PYEXE="
for /f "delims=" %%i in ('where python 2^>nul') do if not defined PYEXE set "PYEXE=%%i"
if not defined PYEXE (
    echo [ERROR] Python 3.14 or newer was not found on PATH.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo [setup] Creating virtual environment...
    "%PYEXE%" -m venv .venv
    if errorlevel 1 goto :failed
)
set "VENV_PY=.venv\Scripts\python.exe"

echo [setup] Installing project requirements...
"%VENV_PY%" -m pip install --quiet --upgrade pip
"%VENV_PY%" -m pip install --quiet -r requirements.txt
if errorlevel 1 goto :failed

echo.
echo Matalan Carton Supply Report
echo http://localhost:8501
echo Database: local PostgreSQL (start it with: docker start carton-pg)
echo Press Ctrl+C to stop.
echo.
"%VENV_PY%" -m streamlit run app.py --server.address=localhost --server.port=8501
goto :end

:failed
echo [ERROR] Setup failed. Check Python and your internet connection.
pause
exit /b 1

:end
endlocal
