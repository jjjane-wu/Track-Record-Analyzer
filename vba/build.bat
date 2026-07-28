@echo off
REM ===================================================================
REM  One-command build: raw inputs workbook -> finished analysis .xlsx
REM
REM  Usage (or drag an inputs workbook onto this file):
REM    build.bat <inputs.xlsx> [output.xlsx]
REM
REM  Expects TR-Analyzer.xlsm next to this script (see vba/README.md
REM  for the one-time assembly of that template).
REM ===================================================================
setlocal
pushd "%~dp0"

if "%~1"=="" (
    echo Usage: build.bat ^<workbook-with-Deal-Level-Inputs.xlsx^> [output.xlsx]
    pause & exit /b 2
)
set "TPL=%~dp0TR-Analyzer.xlsm"
set "OUT=%~2"
if "%OUT%"=="" set "OUT=%~dpn1 - Analysis.xlsx"

cscript //nologo "%~dp0build.vbs" "%TPL%" "%~f1" "%OUT%"
if errorlevel 1 (
    echo.
    echo Build failed - see the message above.
    pause & exit /b 1
)
echo Done: %OUT%
pause
