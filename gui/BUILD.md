# Building Standalone Executables

Build Excel Engine GUI as a standalone application that requires no Python installation.

## Prerequisites

```bash
# From the excel-engine project root:
source .venv/bin/activate  # or your preferred venv
pip install pyinstaller streamlit openpyxl
pip install -e .
```

## Build

```bash
# Build for your current platform
python gui/build_app.py

# Clean build (removes previous artifacts)
python gui/build_app.py --clean
```

## Output

| Platform | Output Location | How to Run |
|----------|----------------|------------|
| macOS | `dist/Excel Engine.app` | Double-click the app |
| Windows | `dist/ExcelEngine/ExcelEngine.exe` | Double-click the exe |
| Linux | `dist/excel-engine-gui/excel-engine-gui` | `./excel-engine-gui` |

## Distribution

### macOS
```bash
# Create a DMG for easy distribution
hdiutil create -volname "Excel Engine" -srcfolder "dist/Excel Engine.app" -ov -format UDZO "dist/ExcelEngine.dmg"
```

### Windows
Zip the `dist/ExcelEngine/` folder and share.

### Linux
Tar the `dist/excel-engine-gui/` folder:
```bash
tar czf excel-engine-gui-linux.tar.gz -C dist excel-engine-gui
```

## Notes

- The build must be run on the target platform (build on macOS for .app, Windows for .exe, etc.)
- First launch may take a few seconds as Streamlit extracts its assets
- The executable includes Python, Streamlit, openpyxl, and the Excel Engine — no external dependencies needed
- File size will be ~100-200MB due to bundled Python + Streamlit
