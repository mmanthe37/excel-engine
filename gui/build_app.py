"""
Excel Engine GUI — Build standalone executables.

Usage:
    python gui/build_app.py          # Build for current platform
    python gui/build_app.py --clean  # Clean build artifacts first

Requirements:
    pip install pyinstaller streamlit openpyxl
"""
import os
import subprocess
import sys
import platform
from pathlib import Path

def build():
    gui_dir = Path(__file__).parent
    project_dir = gui_dir.parent
    
    clean = "--clean" in sys.argv
    
    # Platform-specific settings
    system = platform.system()
    if system == "Darwin":
        name = "Excel Engine"
        icon_flag = []  # Add --icon=gui/icon.icns if icon exists
        extra = ["--windowed"]  # .app bundle on macOS
    elif system == "Windows":
        name = "ExcelEngine"
        icon_flag = []  # Add --icon=gui/icon.ico if icon exists
        extra = ["--windowed"]  # No console window
    else:
        name = "excel-engine-gui"
        icon_flag = []
        extra = []
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", name,
        "--onedir",  # onedir is more reliable than onefile for streamlit
        "--noconfirm",
        # Add Streamlit data files
        "--collect-all", "streamlit",
        "--collect-all", "altair",
        # Add excel_engine package
        "--collect-all", "excel_engine",
        # Add the streamlit app file as data
        "--add-data", f"{gui_dir / 'app.py'}{os.pathsep}.",
        *icon_flag,
        *extra,
        str(gui_dir / "run_app.py"),
    ]
    
    if clean:
        cmd.insert(3, "--clean")
    
    print(f"Building for {system}...")
    print(f"Command: {' '.join(cmd)}")
    print()
    
    subprocess.run(cmd, check=True, cwd=str(project_dir))
    
    print()
    print(f"✅ Build complete!")
    print(f"   Output: {project_dir / 'dist' / name}")

if __name__ == "__main__":
    build()
