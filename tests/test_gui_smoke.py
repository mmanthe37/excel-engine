"""
GUI smoke tests — verify the Streamlit app loads without errors.

Uses streamlit.testing.v1.AppTest to test the GUI at gui/app.py.
"""

import sys
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest


# Ensure the project root is on sys.path so gui/app.py can find excel_engine
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

_APP_PATH = str(_PROJECT_ROOT / "gui" / "app.py")


class TestGuiSmoke:
    def test_app_loads_without_error(self):
        """The app script runs without raising exceptions."""
        at = AppTest.from_file(_APP_PATH, default_timeout=30)
        at.run()
        assert not at.exception, f"App raised: {at.exception}"

    def test_title_rendered(self):
        """Main title is present."""
        at = AppTest.from_file(_APP_PATH, default_timeout=30)
        at.run()
        titles = [t.value for t in at.title]
        assert any("Excel Engine" in t for t in titles)

    def test_sidebar_configuration_exists(self):
        """Sidebar should have a configuration header."""
        at = AppTest.from_file(_APP_PATH, default_timeout=30)
        at.run()
        sidebar = at.sidebar
        # Sidebar renders radio, checkbox, slider widgets
        assert len(sidebar.radio) > 0 or len(sidebar.checkbox) > 0

    def test_file_uploaders_exist(self):
        """Workbook and instruction uploaders should be present."""
        at = AppTest.from_file(_APP_PATH, default_timeout=30)
        at.run()
        uploaders = at.file_uploader
        assert len(uploaders) >= 2, "Expected at least 2 file uploaders"

    def test_run_button_exists(self):
        """The Run button should be present."""
        at = AppTest.from_file(_APP_PATH, default_timeout=30)
        at.run()
        buttons = at.button
        assert len(buttons) >= 1, "Expected at least one button"

    def test_no_result_on_initial_load(self):
        """No results section on first load (nothing has been run)."""
        at = AppTest.from_file(_APP_PATH, default_timeout=30)
        at.run()
        # success/warning blocks should not appear on first load
        assert len(at.success) == 0
