@echo off
REM ===================================================================
REM  Full pipeline, one command:  raw GP file -> finished analysis.
REM
REM     pipeline.bat "<raw GP track record>.xlsx"
REM
REM  (or drag the raw file onto this script)
REM
REM  Step 1  python app\headless.py   parse + auto-map -> inputs sheet
REM  Step 2  cscript build.vbs        inject inputs into TR-Analyzer.xlsm,
REM                                   run the VBA build, save "... - Analysis.xlsx"
REM
REM  Prerequisites (one-time, per machine):
REM    - run start.bat once so the Python venv exists
REM    - assembled TR-Analyzer.xlsm in this vba\ folder (see README.md)
REM    - this folder added to Excel's Trusted Locations
REM
REM  The parse step prints the auto-mapping summary. If it flags
REM  NEEDS-REVIEW / UNMAPPED fields, prefer the Streamlit app for that
REM  GP - the pipeline still builds, but check those columns.
REM ===================================================================
setlocal
pushd "%~dp0.."

if "%~1"=="" (
    echo Usage: pipeline.bat ^<raw GP track record .xlsx^>
    pause & exit /b 2
)
if not exist "venv\Scripts\python.exe" (
    echo Python environment not found - run start.bat once first.
    pause & exit /b 1
)
if not exist "vba\TR-Analyzer.xlsm" (
    echo vba\TR-Analyzer.xlsm not found - assemble it once per vba\README.md.
    pause & exit /b 1
)

set "INPUTS=%TEMP%\tr_inputs.xlsx"
set "OUT=%~dpn1 - Analysis.xlsx"

echo [1/2] Parsing "%~nx1" ...
venv\Scripts\python app\headless.py "%~f1" -o "%INPUTS%"
if errorlevel 1 (
    echo Parse failed - see the message above.
    pause & exit /b 1
)

echo [2/2] Building the analysis workbook in Excel ...
cscript //nologo "vba\build.vbs" "vba\TR-Analyzer.xlsm" "%INPUTS%" "%OUT%"
if errorlevel 1 (
    echo Build failed - see the message above.
    pause & exit /b 1
)
del "%INPUTS%" >nul 2>nul
echo.
echo Done: %OUT%
pause
