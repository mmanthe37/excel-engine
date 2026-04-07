#!/usr/bin/env bash
set -euo pipefail

# ── Excel Engine Copilot Plugin Installer ──
# Installs the plugin to ~/.copilot/installed-plugins/excel-engine/

PLUGIN_NAME="excel-engine"
PLUGIN_SRC="$(cd "$(dirname "$0")" && pwd)"
PLUGIN_DEST="$HOME/.copilot/installed-plugins/$PLUGIN_NAME"
ENGINE_ROOT="$(dirname "$PLUGIN_SRC")"

echo "╔══════════════════════════════════════════╗"
echo "║  Excel Engine — Copilot Plugin Installer ║"
echo "╚══════════════════════════════════════════╝"
echo ""

# ── Step 1: Check Python ──
echo "▸ Checking Python..."
if command -v python3 &>/dev/null; then
    PY_VER=$(python3 --version 2>&1)
    echo "  ✓ $PY_VER"
    if echo "$PY_VER" | grep -q "3.14"; then
        echo "  ⚠ Python 3.14 detected — some tools (uvx/typer) may break."
        echo "    Recommend Python 3.12 or 3.13."
    fi
else
    echo "  ✗ Python 3 not found. Please install Python 3.10+."
    exit 1
fi

# ── Step 2: Install Python dependencies ──
echo ""
echo "▸ Installing Python dependencies..."
pip install --quiet openpyxl xlwings pyautogui python-docx Pillow pdfplumber 2>/dev/null || \
pip3 install --quiet openpyxl xlwings pyautogui python-docx Pillow pdfplumber 2>/dev/null || \
echo "  ⚠ pip install failed — install manually: pip install openpyxl xlwings pyautogui python-docx Pillow pdfplumber"

# Verify core packages
echo "  Checking packages..."
python3 -c "import openpyxl; print('  ✓ openpyxl', openpyxl.__version__)" 2>/dev/null || echo "  ✗ openpyxl missing"
python3 -c "import xlwings; print('  ✓ xlwings', xlwings.__version__)" 2>/dev/null || echo "  ✗ xlwings missing"
python3 -c "import pyautogui; print('  ✓ pyautogui')" 2>/dev/null || echo "  ✗ pyautogui missing"
python3 -c "import docx; print('  ✓ python-docx')" 2>/dev/null || echo "  ✗ python-docx missing"

# ── Step 3: Install Excel Engine package (if available) ──
echo ""
echo "▸ Checking Excel Engine package..."
if [ -f "$ENGINE_ROOT/pyproject.toml" ]; then
    echo "  Found engine at $ENGINE_ROOT"
    pip install --quiet -e "$ENGINE_ROOT" 2>/dev/null || \
    pip3 install --quiet -e "$ENGINE_ROOT" 2>/dev/null || \
    echo "  ⚠ Could not install excel-engine package"
    python3 -c "from excel_engine import ExcelEngine; print('  ✓ ExcelEngine importable')" 2>/dev/null || \
    echo "  ⚠ ExcelEngine not importable — agent will use manual layer commands"
else
    echo "  ⚠ Excel Engine source not found at $ENGINE_ROOT"
fi

# ── Step 4: Copy plugin files ──
echo ""
echo "▸ Installing plugin to $PLUGIN_DEST..."
mkdir -p "$PLUGIN_DEST"

# Remove old installation if exists
if [ -d "$PLUGIN_DEST/agents" ] || [ -d "$PLUGIN_DEST/skills" ]; then
    echo "  Removing previous installation..."
    rm -rf "$PLUGIN_DEST/agents" "$PLUGIN_DEST/skills" "$PLUGIN_DEST/README.md" "$PLUGIN_DEST/install.sh"
fi

# Copy plugin structure
cp -R "$PLUGIN_SRC/agents" "$PLUGIN_DEST/"
cp -R "$PLUGIN_SRC/skills" "$PLUGIN_DEST/"
cp "$PLUGIN_SRC/README.md" "$PLUGIN_DEST/"
cp "$PLUGIN_SRC/install.sh" "$PLUGIN_DEST/"

echo "  ✓ Plugin files installed"

# ── Step 5: Verify environment ──
echo ""
echo "▸ Verifying environment..."

# Check screencapture
if screencapture -x ~/Desktop/_plugin_test.png 2>/dev/null; then
    rm -f ~/Desktop/_plugin_test.png
    echo "  ✓ Screen capture works"
else
    echo "  ⚠ Screen capture failed — grant Terminal screen recording permission"
fi

# Check System Events accessibility
if osascript -e 'tell application "System Events" to get name of first process' &>/dev/null; then
    echo "  ✓ Accessibility permissions OK"
else
    echo "  ⚠ Accessibility permissions needed — grant in System Preferences > Privacy > Accessibility"
fi

# Disable Dictation auto-enable
defaults write com.apple.HIToolbox AppleDictationAutoEnable -int 0 2>/dev/null && \
echo "  ✓ Dictation auto-enable disabled" || true

# ── Done ──
echo ""
echo "╔══════════════════════════════════════════╗"
echo "║  ✅ Installation complete!                ║"
echo "╚══════════════════════════════════════════╝"
echo ""
echo "Plugin installed to: $PLUGIN_DEST"
echo ""
echo "To use, tell Copilot:"
echo '  "Complete this Excel assignment: workbook.xlsx using instructions.docx"'
echo '  "Do my SAM Module 3 for me"'
echo ""
