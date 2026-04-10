#!/bin/bash
# Excel Engine — macOS Launcher
# Double-click this file to start the Excel Engine GUI

set -e

# Navigate to the excel-engine directory (parent of gui/)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR/.."

echo "╔══════════════════════════════════════════╗"
echo "║        📊 Excel Engine GUI               ║"
echo "║   Starting up... please wait             ║"
echo "╚══════════════════════════════════════════╝"
echo ""

# Check for Python 3.10+
PYTHON=""
for cmd in python3 python; do
    if command -v "$cmd" &>/dev/null; then
        version=$("$cmd" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>/dev/null)
        major=$(echo "$version" | cut -d. -f1)
        minor=$(echo "$version" | cut -d. -f2)
        if [ "$major" -ge 3 ] && [ "$minor" -ge 10 ]; then
            PYTHON="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo "❌ Python 3.10 or later is required but not found."
    echo ""
    echo "   Install Python from: https://www.python.org/downloads/"
    echo ""
    echo "   Press any key to exit..."
    read -n 1
    exit 1
fi

echo "✓ Found Python: $($PYTHON --version)"

# Create virtual environment if needed
if [ ! -d ".venv" ]; then
    echo "📦 Setting up virtual environment (first time only)..."
    $PYTHON -m venv .venv
fi

# Activate venv
source .venv/bin/activate

# Install/update dependencies
echo "📦 Checking dependencies..."
pip install -q -r gui/requirements.txt 2>/dev/null
pip install -q -e . 2>/dev/null

echo ""
echo "🚀 Starting Excel Engine GUI..."
echo "   Your browser will open automatically."
echo "   To stop: close this terminal window or press Ctrl+C."
echo ""

# Launch Streamlit (--server.headless false opens browser)
streamlit run gui/app.py --server.headless false --browser.gatherUsageStats false
