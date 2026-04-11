#!/bin/bash
# Excel Engine — Build, Sign, and optionally Notarize the macOS .app
# Double-click this file from Finder to run it.
#
# Prerequisites:
#   - "Developer ID Application" certificate in Keychain
#   - Python venv at .venv/ with dependencies installed
#
# For notarization, set these before running (or export in your shell profile):
#   export APPLE_ID="michaelamanthe2@gmail.com"
#   export APPLE_TEAM_ID="25QSUNYFC9"
#   export NOTARY_PASSWORD="xxxx-xxxx-xxxx-xxxx"  # app-specific password

set -e
cd "$(dirname "$0")"

echo "========================================="
echo "  Excel Engine — Build & Sign"
echo "========================================="
echo ""

# Activate virtual environment
if [ -d ".venv" ]; then
    source .venv/bin/activate
else
    echo "ERROR: No .venv found. Run: python -m venv .venv && pip install -e '.[all]' -r gui/requirements.txt pyinstaller"
    echo "Press Enter to close..."
    read
    exit 1
fi

# Build + Sign
echo "Building and signing (you may see a Keychain dialog — click 'Always Allow')..."
echo ""
python gui/build_app.py --clean --sign

# Notarize if credentials are available
if [ -n "$APPLE_ID" ] && [ -n "$NOTARY_PASSWORD" ]; then
    echo ""
    echo "Notarizing with Apple..."
    python gui/build_app.py --sign-only --notarize
else
    echo ""
    echo "Skipping notarization (set APPLE_ID + NOTARY_PASSWORD to enable)."
    echo "Without notarization, users must right-click > Open the first time."
fi

echo ""
echo "========================================="
echo "  Done! App is at: dist/Excel Engine.app"
echo "========================================="
echo ""
echo "Press Enter to close..."
read
