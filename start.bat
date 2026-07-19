@echo off
REM ════════════════════════════════════════════════════════════════════
REM  GP Track Record Analyzer — double-click launcher (Windows)
REM  First run: sets itself up (needs internet, a few minutes).
REM  Every run after that: starts instantly, no internet needed.
REM ════════════════════════════════════════════════════════════════════
cd /d "%~dp0"

REM ── Find Python ─────────────────────────────────────────────────────
set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if not defined PY (
    where python >nul 2>nul
    if not errorlevel 1 set "PY=python"
)
if not defined PY (
    echo.
    echo  Python was not found on this computer.
    echo.
    echo  1. Go to  https://www.python.org/downloads/  and install Python 3.10+
    echo  2. IMPORTANT: tick "Add python.exe to PATH" on the first install screen
    echo  3. Double-click this file again
    echo.
    pause
    exit /b 1
)

REM ── First-time setup ────────────────────────────────────────────────
if not exist "venv" (
    echo First-time setup: creating the app's private Python environment...
    %PY% -m venv venv
    if errorlevel 1 goto :fail
    echo Installing components — this needs internet and takes a few minutes...
    venv\Scripts\python -m pip install --upgrade pip >nul 2>nul
    venv\Scripts\pip install -r requirements.txt
    if errorlevel 1 goto :fail
)

REM ── Launch ──────────────────────────────────────────────────────────
echo.
echo Starting the GP Track Record Analyzer — your browser will open shortly.
echo Keep this black window open while you work; close it to stop the app.
echo.
venv\Scripts\streamlit run app\app.py
pause
exit /b 0

:fail
echo.
echo  Setup did not finish. Please take a screenshot of this window and
echo  send it to the tool maintainer. (Common cause: the office network
echo  blocks Python package downloads — IT can allow pypi.org.)
echo.
pause
exit /b 1
