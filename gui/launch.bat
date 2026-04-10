@echo off
:: Excel Engine — Windows Launcher
:: Double-click this file to start the Excel Engine GUI

title Excel Engine GUI
echo ============================================
echo        Excel Engine GUI
echo    Starting up... please wait
echo ============================================
echo.

:: Navigate to excel-engine directory (parent of gui\)
cd /d "%~dp0\.."

:: Check for Python
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo.
    echo    Download Python from: https://www.python.org/downloads/
    echo    IMPORTANT: Check "Add Python to PATH" during installation!
    echo.
    pause
    exit /b 1
)

:: Check Python version
python -c "import sys; exit(0 if sys.version_info >= (3, 10) else 1)" 2>nul
if errorlevel 1 (
    echo ERROR: Python 3.10 or later is required.
    echo    Your version:
    python --version
    echo.
    echo    Download Python 3.10+ from: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo    Found Python:
python --version
echo.

:: Create virtual environment if needed
if not exist ".venv" (
    echo    Setting up virtual environment (first time only^)...
    python -m venv .venv
)

:: Activate venv
call .venv\Scripts\activate.bat

:: Install/update dependencies
echo    Checking dependencies...
pip install -q -r gui\requirements.txt 2>nul
pip install -q -e . 2>nul

echo.
echo    Starting Excel Engine GUI...
echo    Your browser will open automatically.
echo    To stop: close this window or press Ctrl+C.
echo.

:: Launch Streamlit
streamlit run gui\app.py --server.headless false --browser.gatherUsageStats false
