#!/bin/bash
# JANDS Price Comparator — Mac launcher
# Double-click this file to start the app.

cd "$(dirname "$0")"

# Check Python 3 is available
if ! command -v python3 &>/dev/null; then
    osascript -e 'display alert "Python 3 not found" message "Please install Python 3 from https://www.python.org/downloads/ then try again."'
    exit 1
fi

# Create virtual environment if it doesn't exist
if [ ! -d ".venv" ]; then
    echo "Setting up environment for the first time (this takes about a minute)..."
    python3 -m venv .venv
    .venv/bin/pip install --quiet --upgrade pip
    .venv/bin/pip install --quiet -r requirements.txt
    echo "Setup complete."
fi

echo "Starting JANDS Price Comparator..."
echo "Open your browser to: http://localhost:8501"
echo "(Press Ctrl+C in this window to stop the app)"
.venv/bin/python -m streamlit run app.py --server.headless true
