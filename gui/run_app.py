"""
Excel Engine GUI — PyInstaller entry point.

This wrapper launches the Streamlit app programmatically,
which is required for PyInstaller bundling since Streamlit
normally runs via `streamlit run app.py`.
"""
import sys
import os
from pathlib import Path

def main():
    # When running as a PyInstaller bundle, files are in a temp dir
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        bundle_dir = Path(sys._MEIPASS)
        app_path = bundle_dir / "app.py"
    else:
        # Running as normal Python script
        app_path = Path(__file__).parent / "app.py"
    
    # Set up streamlit arguments
    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.headless", "false",
        "--browser.gatherUsageStats", "false",
        "--global.developmentMode", "false",
    ]
    
    # Import and run streamlit
    from streamlit.web.cli import main as st_main
    st_main()

if __name__ == "__main__":
    main()
