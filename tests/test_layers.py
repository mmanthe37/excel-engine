"""
Tests for individual layer modules.
"""

import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock

from excel_engine.layers.applescript_layer import AppleScriptLayer
from excel_engine.layers.system_events import SystemEventsLayer
from excel_engine.layers.vba_layer import VBALayer
from excel_engine.layers.pyautogui_layer import PyAutoGUILayer
from excel_engine.utils.path_handler import PathHandler
from excel_engine.utils.excel_constants import ExcelConstants
from excel_engine.utils.mac_utils import MacUtils


class TestPathHandler:
    def setup_method(self):
        self.handler = PathHandler()

    def test_has_unsafe_chars(self):
        assert self.handler.has_unsafe_chars(Path("/path/with:colon/file.xlsx"))
        assert not self.handler.has_unsafe_chars(Path("/normal/path/file.xlsx"))

    def test_sanitize_filename(self):
        assert PathHandler.sanitize_filename("file:name.xlsx") == "file_name.xlsx"
        assert PathHandler.sanitize_filename("normal.xlsx") == "normal.xlsx"

    def test_to_hfs(self):
        hfs = PathHandler.to_hfs("/Users/me/Desktop/file.xlsx")
        assert hfs.startswith("Macintosh HD:")
        assert "Users" in hfs
        assert "Desktop" in hfs

    def test_ensure_xlsx_extension(self):
        assert PathHandler.ensure_xlsx_extension(Path("file.xlsx")).suffix == ".xlsx"
        assert PathHandler.ensure_xlsx_extension(Path("file.xlsm")).suffix == ".xlsm"
        assert PathHandler.ensure_xlsx_extension(Path("file.csv")).suffix == ".xlsx"


class TestExcelConstants:
    def test_subtotal_functions(self):
        assert ExcelConstants.XL_SUM == -4157
        assert ExcelConstants.XL_AVERAGE == -4106
        assert ExcelConstants.SUBTOTAL_FUNCTIONS["sum"] == -4157
        assert ExcelConstants.SUBTOTAL_FUNCTIONS["average"] == -4106

    def test_table_styles(self):
        assert "TableStyleMedium5" in ExcelConstants.TABLE_STYLES
        assert "TableStyleLight1" in ExcelConstants.TABLE_STYLES
        assert len(ExcelConstants.TABLE_STYLES) > 30

    def test_chart_types(self):
        assert ExcelConstants.CHART_TYPES["bar"] == "BarChart"
        assert ExcelConstants.CHART_TYPES["line"] == "LineChart"
        assert ExcelConstants.CHART_TYPES["pie"] == "PieChart"

    def test_number_formats(self):
        assert ExcelConstants.NUMBER_FORMATS["currency"] == "$#,##0.00"
        assert ExcelConstants.NUMBER_FORMATS["percentage"] == "0.00%"

    def test_border_styles(self):
        assert "thin" in ExcelConstants.BORDER_STYLES
        assert "medium" in ExcelConstants.BORDER_STYLES
        assert "thick" in ExcelConstants.BORDER_STYLES

    def test_ribbon_tabs(self):
        assert ExcelConstants.RIBBON_TABS["insert"] == "Insert"
        assert ExcelConstants.RIBBON_TABS["data"] == "Data"


class TestAppleScriptLayer:
    @patch("excel_engine.utils.mac_utils.MacUtils.run_applescript")
    def test_activate_sheet(self, mock_run):
        mock_run.return_value = ""
        layer = AppleScriptLayer()
        layer.activate_sheet("Sales Data")
        mock_run.assert_called_once()
        call_args = mock_run.call_args[0][0]
        assert "Sales Data" in call_args
        assert "activate object sheet" in call_args

    @patch("excel_engine.utils.mac_utils.MacUtils.run_applescript")
    def test_set_cell_formula(self, mock_run):
        mock_run.return_value = ""
        layer = AppleScriptLayer()
        layer.set_cell_formula("G3", '=CONCAT([@LAST],", ",[@FIRST])')
        call_args = mock_run.call_args[0][0]
        assert "set formula" in call_args
        assert "G3" in call_args

    @patch("excel_engine.utils.mac_utils.MacUtils.run_applescript")
    def test_sort_range(self, mock_run):
        mock_run.return_value = ""
        layer = AppleScriptLayer()
        layer.sort_range(
            "A1:G50",
            keys=[
                {"range": "A1:A50", "order": "ascending"},
                {"range": "B1:B50", "order": "descending"},
            ],
        )
        call_args = mock_run.call_args[0][0]
        assert "sort range" in call_args
        assert "key1" in call_args
        assert "key2" in call_args
        assert "sort ascending" in call_args
        assert "sort descending" in call_args

    @patch("excel_engine.utils.mac_utils.MacUtils.run_applescript")
    def test_freeze_panes(self, mock_run):
        mock_run.return_value = ""
        layer = AppleScriptLayer()
        layer.freeze_panes("A2")
        call_args = mock_run.call_args[0][0]
        assert "freeze panes" in call_args

    @patch("excel_engine.utils.mac_utils.MacUtils.run_applescript")
    def test_save(self, mock_run):
        mock_run.return_value = ""
        layer = AppleScriptLayer()
        layer.save()
        call_args = mock_run.call_args[0][0]
        assert "save workbook 1" in call_args

    @patch("excel_engine.utils.mac_utils.MacUtils.run_applescript")
    def test_save_as_xlsx(self, mock_run):
        mock_run.return_value = ""
        layer = AppleScriptLayer()
        layer.save_as_xlsx("/Users/me/Desktop", "output")
        call_args = mock_run.call_args[0][0]
        assert "Excel XML file format" in call_args

    @patch("excel_engine.utils.mac_utils.MacUtils.run_applescript")
    def test_escape_quotes(self, mock_run):
        mock_run.return_value = ""
        layer = AppleScriptLayer()
        layer.set_cell_value("A1", 'He said "hello"')
        call_args = mock_run.call_args[0][0]
        assert '\\"hello\\"' in call_args


class TestVBALayer:
    def test_pivot_field_code(self):
        code = VBALayer._pivot_field_code(["Sales", "Region"], "xlRowField")
        assert "Sales" in code
        assert "Region" in code
        assert "xlRowField" in code

    def test_pivot_value_code(self):
        code = VBALayer._pivot_value_code([
            {"name": "Amount", "function": "xlSum", "number_format": "$#,##0"},
        ])
        assert "Amount" in code
        assert "xlSum" in code
        assert "$#,##0" in code

    def test_generate_clean_sub(self):
        layer = VBALayer()
        sub = layer.generate_clean_sub("PivotTable")
        assert "CleanExistingObjects" in sub
        assert "PivotTable" in sub


class TestPyAutoGUILayer:
    def test_retina_adjustment(self):
        layer = PyAutoGUILayer(retina=True)
        x, y = layer._adjust_coords(2000, 1600)
        assert x == 1000
        assert y == 800

    def test_no_retina_adjustment(self):
        layer = PyAutoGUILayer(retina=False)
        x, y = layer._adjust_coords(2000, 1600)
        assert x == 2000
        assert y == 1600


class TestMacUtils:
    def test_retina_to_logical(self):
        x, y = MacUtils.retina_to_logical(2880, 1800)
        assert x == 1440
        assert y == 900

    def test_logical_to_retina(self):
        x, y = MacUtils.logical_to_retina(1440, 900)
        assert x == 2880
        assert y == 1800
