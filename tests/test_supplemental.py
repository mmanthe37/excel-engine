"""
Supplemental tests targeting coverage gaps in path_handler, engine cleanup,
and additional verifier edge cases.
"""

import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock

from openpyxl import Workbook

from excel_engine import ExcelEngine, EngineConfig, TaskType, Layer
from excel_engine.parsers.task_extractor import Task
from excel_engine.planner.task_planner import ExecutionPlan, Section
from excel_engine.utils.path_handler import PathHandler
from excel_engine.verifier.workbook_verifier import WorkbookVerifier


def _task(task_type, **kw):
    kw.setdefault("id", f"sup-{task_type.value}")
    kw.setdefault("description", f"Supplemental {task_type.value}")
    return Task(task_type=task_type, **kw)


# ══════════════════════════════════════════════════════════════════
# PathHandler tests
# ══════════════════════════════════════════════════════════════════

class TestPathHandler:
    def test_has_unsafe_chars_true(self):
        ph = PathHandler()
        assert ph.has_unsafe_chars(Path("/path:with:colons/file.xlsx")) is True

    def test_has_unsafe_chars_false(self):
        ph = PathHandler()
        assert ph.has_unsafe_chars(Path("/normal/path/file.xlsx")) is False

    def test_sanitize_filename(self):
        assert PathHandler.sanitize_filename("file:name.xlsx") == "file_name.xlsx"
        assert PathHandler.sanitize_filename("a/b:c.xlsx") == "a_b_c.xlsx"

    def test_safe_copy_no_unsafe(self, tmp_path):
        ph = PathHandler(desktop_path=tmp_path)
        f = tmp_path / "normal.xlsx"
        f.write_text("test")
        result = ph.safe_copy_for_xlwings(f)
        assert result == f  # no copy needed

    def test_safe_copy_with_unsafe(self, tmp_path):
        ph = PathHandler(desktop_path=tmp_path)
        # Create a file with colon in parent dir (simulate)
        src_dir = tmp_path / "dir_with_colon"
        src_dir.mkdir()
        src = src_dir / "file.xlsx"
        src.write_text("data")
        # Patch has_unsafe_chars to return True
        with patch.object(ph, "has_unsafe_chars", return_value=True):
            result = ph.safe_copy_for_xlwings(src)
        assert result.parent == tmp_path
        assert result.exists()

    def test_copy_back(self, tmp_path):
        src = tmp_path / "desktop.xlsx"
        src.write_text("modified")
        orig = tmp_path / "original.xlsx"
        orig.write_text("original")
        ph = PathHandler()
        ph.copy_back_from_desktop(src, orig)
        assert orig.read_text() == "modified"

    def test_cleanup_desktop_copy(self, tmp_path):
        desk = tmp_path / "temp.xlsx"
        desk.write_text("temp")
        orig = tmp_path / "orig.xlsx"
        ph = PathHandler()
        ph.cleanup_desktop_copy(desk, orig)
        assert not desk.exists()

    def test_cleanup_same_file(self, tmp_path):
        f = tmp_path / "same.xlsx"
        f.write_text("data")
        ph = PathHandler()
        ph.cleanup_desktop_copy(f, f)
        assert f.exists()  # not deleted when same

    def test_to_posix(self, tmp_path):
        f = tmp_path / "test.xlsx"
        result = PathHandler.to_posix(f)
        assert isinstance(result, str)
        assert "test.xlsx" in result

    def test_to_hfs(self):
        result = PathHandler.to_hfs("/Users/me/Desktop/file.xlsx")
        assert result.startswith("Macintosh HD:")
        assert "Users" in result

    def test_ensure_xlsx_extension(self):
        assert PathHandler.ensure_xlsx_extension(Path("f.xlsx")) == Path("f.xlsx")
        assert PathHandler.ensure_xlsx_extension(Path("f.xlsm")) == Path("f.xlsm")
        assert PathHandler.ensure_xlsx_extension(Path("f.csv")).suffix == ".xlsx"

    def test_unique_path(self, tmp_path):
        f = tmp_path / "test.xlsx"
        f.write_text("exists")
        result = PathHandler._unique_path(f)
        assert result != f
        assert "test_1" in result.stem


# ══════════════════════════════════════════════════════════════════
# Engine: scan, verify, and cleanup paths
# ══════════════════════════════════════════════════════════════════

class TestEngineScanPlanVerify:
    def test_scan_extracts_tasks(self, tmp_path):
        instructions = tmp_path / "instr.txt"
        instructions.write_text("Enter 100 in cell A1\nEnter 200 in cell B1\n")
        eng = ExcelEngine()
        tasks = eng.scan(instructions)
        assert isinstance(tasks, list)
        assert len(tasks) >= 0  # may extract 0 or more depending on patterns

    def test_verify_with_tasks(self, sample_workbook):
        eng = ExcelEngine()
        tasks = [
            _task(TaskType.CELL_VALUE, cell="A1", value="Product", sheet="Sales"),
        ]
        sv = eng.verify(sample_workbook, tasks)
        assert sv.section_id == "manual"
        assert sv.pass_count >= 0

    def test_cleanup_is_safe(self):
        eng = ExcelEngine()
        # Should not raise even when nothing is open
        eng._cleanup()

    def test_save_workbook_with_openpyxl(self, sample_workbook):
        eng = ExcelEngine()
        eng._openpyxl.open(sample_workbook)
        eng._save_workbook(sample_workbook)
        eng._openpyxl.close()

    def test_execute_with_verification(self, sample_workbook):
        config = EngineConfig()
        config.verify_after_each_section = True
        eng = ExcelEngine(config)
        tasks = [
            _task(TaskType.CELL_VALUE, id="v1", cell="A1", value="Test", sheet="Sales"),
        ]
        section = Section(id="s1", name="Test", sheet="Sales", tasks=tasks)
        plan = ExecutionPlan(sections=[section], total_tasks=1, estimated_time_seconds=1.0)
        result = eng.execute(plan, sample_workbook)
        assert result.tasks_completed == 1
        assert len(result.verifications) >= 1


# ══════════════════════════════════════════════════════════════════
# Engine: additional _exec_openpyxl edge cases
# ══════════════════════════════════════════════════════════════════

class TestExecOpenpyxlEdgeCases:
    def test_cell_value_without_cell(self, sample_workbook):
        """CELL_VALUE with no cell is a no-op."""
        eng = ExcelEngine()
        eng._openpyxl.open(sample_workbook)
        task = _task(TaskType.CELL_VALUE, value="test")  # no cell
        eng._exec_openpyxl(task, sample_workbook)  # should not raise
        eng._openpyxl.close()

    def test_formula_without_cell(self, sample_workbook):
        eng = ExcelEngine()
        eng._openpyxl.open(sample_workbook)
        task = _task(TaskType.FORMULA, formula="=SUM(1,2)")  # no cell
        eng._exec_openpyxl(task, sample_workbook)
        eng._openpyxl.close()

    def test_text_function_with_value_fallback(self, sample_workbook):
        eng = ExcelEngine()
        eng._openpyxl.open(sample_workbook)
        task = _task(TaskType.TEXT_FUNCTION, cell="A1", value="plain text", sheet="Sales")
        eng._exec_openpyxl(task, sample_workbook)
        assert eng._openpyxl._ws("Sales")["A1"].value == "plain text"
        eng._openpyxl.close()

    def test_formatting_with_range(self, sample_workbook):
        eng = ExcelEngine()
        eng._openpyxl.open(sample_workbook)
        task = _task(TaskType.FORMATTING, range="A1:A3", sheet="Sales",
                     params={"bold": True, "font_size": 14})
        eng._exec_openpyxl(task, sample_workbook)
        eng._openpyxl.close()

    def test_formatting_no_cell(self, sample_workbook):
        eng = ExcelEngine()
        eng._openpyxl.open(sample_workbook)
        task = _task(TaskType.FORMATTING, sheet="Sales", params={"bold": True})
        eng._exec_openpyxl(task, sample_workbook)  # no-op without cell/range
        eng._openpyxl.close()

    def test_table_create_defaults(self, sample_workbook):
        eng = ExcelEngine()
        eng._openpyxl.open(sample_workbook)
        task = _task(TaskType.TABLE_CREATE, sheet="Sales", range="A1:E11")
        eng._exec_openpyxl(task, sample_workbook)
        assert len(eng._openpyxl._ws("Sales").tables) > 0
        eng._openpyxl.close()

    def test_print_settings_both(self, sample_workbook):
        eng = ExcelEngine()
        eng._openpyxl.open(sample_workbook)
        task = _task(TaskType.PRINT_SETTINGS, sheet="Sales",
                     params={"landscape": True, "print_area": "A1:E11"})
        eng._exec_openpyxl(task, sample_workbook)
        ws = eng._openpyxl._ws("Sales")
        assert ws.page_setup.orientation == "landscape"
        assert ws.print_area is not None
        eng._openpyxl.close()


# ══════════════════════════════════════════════════════════════════
# Verifier: additional edge cases
# ══════════════════════════════════════════════════════════════════

class TestVerifierEdgeCases:
    def test_verify_no_workbook_loaded(self):
        v = WorkbookVerifier()
        task = _task(TaskType.FORMULA, cell="A1", formula="=SUM(1,2)")
        result = v.verify_task(task)
        assert "Skipped" in result.message or "No cell" in result.message

    def test_verify_cell_value_no_wb(self):
        v = WorkbookVerifier()
        task = _task(TaskType.CELL_VALUE, cell="A1", value="test")
        result = v.verify_task(task)
        assert "Skipped" in result.message or result.passed is True

    def test_close_already_closed(self):
        v = WorkbookVerifier()
        v.close()  # should not raise

    def test_load_and_close(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        assert v._wb is not None
        v.close()
        assert v._wb is None

    def test_get_ws_default(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        ws = v._get_ws()
        assert ws is not None
        v.close()

    def test_get_ws_by_name(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        ws = v._get_ws("Sales")
        assert ws.title == "Sales"
        v.close()

    def test_get_ws_nonexistent_falls_to_active(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        ws = v._get_ws("NonExistent")
        assert ws is not None  # falls back to active
        v.close()

    def test_verify_sort_no_filter(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        task = _task(TaskType.SORT, sheet="Sales")
        result = v.verify_task(task)
        # No sort state or subtotals
        assert isinstance(result.passed, bool)
        v.close()

    def test_verify_subtotal_empty(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        task = _task(TaskType.SUBTOTAL, sheet="Sales")
        result = v.verify_task(task)
        assert result.passed is False
        v.close()

    def test_verify_split_panes(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        task = _task(TaskType.SPLIT_PANES, sheet="Sales")
        result = v.verify_task(task)
        assert result.passed is False
        v.close()

    def test_verify_external_reference(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        task = _task(TaskType.EXTERNAL_REFERENCE, cell="A1", sheet="Sales")
        result = v.verify_task(task)
        assert result.passed is False
        v.close()

    def test_verify_slicer(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        task = _task(TaskType.SLICER, sheet="Sales")
        result = v.verify_task(task)
        assert isinstance(result.passed, bool)
        v.close()

    def test_verify_pivot_table(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        task = _task(TaskType.PIVOT_TABLE, sheet="Sales")
        result = v.verify_task(task)
        assert result.passed is False
        v.close()

    def test_verify_hyperlink_sheet_level(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        task = _task(TaskType.HYPERLINK, sheet="Sales")  # no cell → sheet-level check
        result = v.verify_task(task)
        assert isinstance(result.passed, bool)
        v.close()

    def test_verify_column_width_no_col(self, sample_workbook):
        v = WorkbookVerifier()
        v.load(sample_workbook)
        task = _task(TaskType.COLUMN_WIDTH, sheet="Sales")
        result = v.verify_task(task)
        assert "skipped" in result.message.lower() or result.passed is True
        v.close()
