#!/usr/bin/env bash
# ════════════════════════════════════════════════════════════════════
#  GP Track Record Analyzer — double-click launcher (Mac)
#  First run: sets itself up (needs internet, a few minutes).
#  Every run after that: starts instantly, no internet needed.
# ════════════════════════════════════════════════════════════════════
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
    echo
    echo "  Python was not found on this computer."
    echo
    echo "  1. Go to https://www.python.org/downloads/ and install Python 3.10+"
    echo "  2. Double-click this file again"
    echo
    read -r -p "Press Enter to close..."
    exit 1
fi

if [ ! -d venv ]; then
    echo "First-time setup: creating the app's private Python environment..."
    python3 -m venv venv || { echo "Setup failed at environment creation."; read -r -p "Press Enter to close..."; exit 1; }
    echo "Installing components — this needs internet and takes a few minutes..."
    venv/bin/python -m pip install --upgrade pip >/dev/null 2>&1
    venv/bin/pip install -r requirements.txt || {
        echo
        echo "  Setup did not finish. Please screenshot this window and send it"
        echo "  to the tool maintainer. (Common cause: the office network blocks"
        echo "  Python package downloads — IT can allow pypi.org.)"
        echo
        read -r -p "Press Enter to close..."
        exit 1
    }
fi

echo
echo "Starting the GP Track Record Analyzer — your browser will open shortly."
echo "Keep this window open while you work; close it to stop the app."
echo
venv/bin/streamlit run app/app.py
