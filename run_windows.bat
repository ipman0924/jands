@echo off
:: JANDS Price Comparator — Windows launcher
:: Double-click this file to start the app.

cd /d "%~dp0"

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found. Please install Python 3 from https://www.python.org/downloads/
    echo Make sure to tick "Add Python to PATH" during installation.
    pause
    exit /b 1
)

:: Create virtual environment if it doesn't exist
if not exist ".venv\" (
    echo Setting up environment for the first time - this takes about a minute...
    python -m venv .venv
    .venv\Scripts\pip install --quiet --upgrade pip
    .venv\Scripts\pip install --quiet -r requirements.txt
    echo Setup complete.
)

echo Starting JANDS Price Comparator...
echo Open your browser to: http://localhost:8501
echo Press Ctrl+C in this window to stop the app.
.venv\Scripts\python -m streamlit run app.py --server.headless true
pause
