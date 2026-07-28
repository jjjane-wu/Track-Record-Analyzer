#!/usr/bin/env bash
# Headless-ish build on macOS (for development testing on this laptop).
# Usage: ./mac_build.sh /path/to/TR-Analyzer.xlsm
# The workbook must already contain a "Deal Level Inputs" sheet.
set -e
F="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
[ -f "$F" ] || { echo "not found: $F"; exit 1; }
osascript <<EOF
tell application "Microsoft Excel"
    activate
    open (POSIX file "$F")
    run VB macro "BuildHeadless"
    save active workbook
end tell
EOF
echo "OK: $F"
