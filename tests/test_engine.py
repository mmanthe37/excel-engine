"""
Tests for the main ExcelEngine orchestrator.
"""

import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from excel_engine.engine import ExcelEngine, EngineResult
from excel_engine.config import EngineConfig, TaskType, Layer
from excel_engine.parsers.task_extractor import Task


class TestEngineConfig:
    def test_default_config(self):
        config = EngineConfig()
        assert config.scan_timeout == 120.0
        assert config.max_retries == 3
        assert config.sam_fingerprint_protected is True
        assert config.retina_display is True

    def test_layer_order(self):
        config = EngineConfig()
        assert config.layer_order[0] == Layer.OPENPYXL
        assert config.layer_order[-1] == Layer.PYAUTOGUI

    def test_get_layers_for_task(self):
        config = EngineConfig()
        layers = config.get_layers_for_task(TaskType.FORMULA)
        assert Layer.OPENPYXL in layers
        assert layers[0] == Layer.OPENPYXL  # preferred

    def test_slicer_requires_system_events(self):
        config = EngineConfig()
        layers = config.get_layers_for_task(TaskType.SLICER)
        assert layers == [Layer.SYSTEM_EVENTS]

    def test_calculated_column_prefers_live(self):
        config = EngineConfig()
        layers = config.get_layers_for_task(TaskType.CALCULATED_COLUMN)
        assert Layer.XLWINGS in layers
        assert Layer.OPENPYXL not in layers  # would produce #REF!

    def test_histogram_requires_ui(self):
        config = EngineConfig()
        layers = config.get_layers_for_task(TaskType.CHART_HISTOGRAM)
        assert layers == [Layer.SYSTEM_EVENTS]

    def test_scatter_starts_with_openpyxl(self):
        config = EngineConfig()
        layers = config.get_layers_for_task(TaskType.CHART_SCATTER)
        assert layers[0] == Layer.OPENPYXL

    def test_area_starts_with_openpyxl(self):
        config = EngineConfig()
        layers = config.get_layers_for_task(TaskType.CHART_AREA)
        assert layers[0] == Layer.OPENPYXL

    def test_combo_has_openpyxl_fallback(self):
        config = EngineConfig()
        layers = config.get_layers_for_task(TaskType.CHART_COMBO)
        assert Layer.OPENPYXL in layers

    def test_sparkline_needs_live(self):
        config = EngineConfig()
        layers = config.get_layers_for_task(TaskType.SPARKLINE)
        assert Layer.XLWINGS in layers
        assert Layer.OPENPYXL not in layers

    def test_pivot_requires_vba(self):
        config = EngineConfig()
        layers = config.get_layers_for_task(TaskType.PIVOT_TABLE)
        assert layers == [Layer.VBA]


class TestExcelEngine:
    def test_engine_init(self):
        engine = ExcelEngine()
        assert engine.openpyxl is not None
        assert engine.applescript is not None
        assert engine.config.verify_after_each_section is True

    def test_engine_custom_config(self):
        config = EngineConfig(max_retries=5, retina_display=False)
        engine = ExcelEngine(config=config)
        assert engine.config.max_retries == 5
        assert engine.pyautogui.retina is False

    def test_scan_with_text(self):
        engine = ExcelEngine()
        tasks = engine.extractor.extract(
            "In cell A1, enter the formula =SUM(B1:B10)"
        )
        assert len(tasks) > 0
        assert any(t.task_type == TaskType.FORMULA for t in tasks)

    def test_plan_empty_tasks(self):
        engine = ExcelEngine()
        plan = engine.plan([])
        assert plan.section_count == 0
        assert plan.total_tasks == 0

    def test_plan_with_tasks(self):
        engine = ExcelEngine()
        tasks = [
            Task(id="t1", task_type=TaskType.FORMULA, description="test formula",
                 sheet="Sheet1", cell="A1", formula="=SUM(B1:B10)"),
            Task(id="t2", task_type=TaskType.TABLE_CREATE, description="create table",
                 sheet="Sheet1", range="A1:D10"),
        ]
        plan = engine.plan(tasks)
        assert plan.total_tasks == 2
        assert plan.section_count >= 1

    def test_engine_result_summary(self):
        result = EngineResult(
            success=True,
            workbook_path=Path("test.xlsx"),
            sections_completed=3,
            sections_total=3,
            tasks_completed=10,
            tasks_total=10,
            elapsed_seconds=45.2,
        )
        summary = result.summary()
        assert "SUCCESS" in summary
        assert "10/10" in summary


# ── Recovery tests ──

from excel_engine.recovery import ErrorClassifier, RecoveryStrategy, TaskError


class TestErrorClassifier:
    def test_transient_timeout(self):
        err = RuntimeError("AppleEvent timed out after 30s")
        assert ErrorClassifier.classify(err, Layer.APPLESCRIPT) == "transient"

    def test_transient_busy(self):
        err = RuntimeError("Excel is busy")
        assert ErrorClassifier.classify(err, Layer.XLWINGS) == "transient"

    def test_transient_not_responding(self):
        err = RuntimeError("Application not responding")
        assert ErrorClassifier.classify(err, Layer.APPLESCRIPT) == "transient"

    def test_permanent_file_not_found(self):
        err = FileNotFoundError("No such file")
        assert ErrorClassifier.classify(err, Layer.OPENPYXL) == "permanent"

    def test_permanent_no_such_sheet(self):
        err = KeyError("No such sheet 'Data'")
        assert ErrorClassifier.classify(err, Layer.OPENPYXL) == "permanent"

    def test_permanent_not_supported(self):
        err = ValueError("Feature not supported in this layer")
        assert ErrorClassifier.classify(err, Layer.OPENPYXL) == "permanent"

    def test_layer_incompatible_not_implemented(self):
        err = NotImplementedError("sparklines not available")
        assert ErrorClassifier.classify(err, Layer.OPENPYXL) == "layer_incompatible"

    def test_layer_incompatible_attribute_error(self):
        err = AttributeError("no attribute 'create_sparkline'")
        assert ErrorClassifier.classify(err, Layer.OPENPYXL) == "layer_incompatible"

    def test_unknown_error_defaults_permanent(self):
        err = ZeroDivisionError("boom")
        assert ErrorClassifier.classify(err, Layer.OPENPYXL) == "permanent"


class TestRecoveryStrategy:
    def test_should_retry_transient(self):
        rs = RecoveryStrategy(max_retries=3)
        assert rs.should_retry("transient", 0) is True
        assert rs.should_retry("transient", 2) is True
        assert rs.should_retry("transient", 3) is False

    def test_should_not_retry_permanent(self):
        rs = RecoveryStrategy(max_retries=3)
        assert rs.should_retry("permanent", 0) is False

    def test_should_not_retry_layer_incompatible(self):
        rs = RecoveryStrategy(max_retries=3)
        assert rs.should_retry("layer_incompatible", 0) is False

    def test_get_delay_exponential(self):
        rs = RecoveryStrategy(base_delay=1.0, max_delay=30.0)
        d0 = rs.get_delay(0)
        d1 = rs.get_delay(1)
        d2 = rs.get_delay(2)
        # base * 2^attempt + small jitter
        assert 1.0 <= d0 <= 1.2
        assert 2.0 <= d1 <= 2.3
        assert 4.0 <= d2 <= 4.5

    def test_get_delay_capped(self):
        rs = RecoveryStrategy(base_delay=1.0, max_delay=5.0)
        d10 = rs.get_delay(10)  # 1 * 2^10 = 1024, capped to 5
        assert d10 <= 5.6  # 5 + 10% jitter

    def test_get_delay_custom_base(self):
        rs = RecoveryStrategy(base_delay=2.0, max_delay=30.0)
        d0 = rs.get_delay(0)
        assert 2.0 <= d0 <= 2.3


class TestEngineResultWithErrors:
    def test_summary_includes_failed_tasks(self):
        result = EngineResult(
            success=False,
            workbook_path=Path("test.xlsx"),
            sections_completed=1,
            sections_total=2,
            tasks_completed=5,
            tasks_total=8,
            failed_tasks=["t3", "t5", "t7"],
            elapsed_seconds=30.0,
        )
        summary = result.summary()
        assert "FAILED" in summary
        assert "Failed tasks (3)" in summary
        assert "t3" in summary

    def test_summary_includes_task_error_breakdown(self):
        import time as _time
        errors = [
            TaskError("t1", TaskType.FORMULA, Layer.OPENPYXL, "transient", "busy", _time.time()),
            TaskError("t1", TaskType.FORMULA, Layer.OPENPYXL, "transient", "busy", _time.time()),
            TaskError("t1", TaskType.FORMULA, Layer.XLWINGS, "permanent", "no sheet", _time.time()),
        ]
        result = EngineResult(
            success=False,
            workbook_path=Path("test.xlsx"),
            sections_completed=0,
            sections_total=1,
            tasks_completed=0,
            tasks_total=1,
            task_errors=errors,
            elapsed_seconds=10.0,
        )
        summary = result.summary()
        assert "Task errors: 3" in summary
        assert "transient=2" in summary
        assert "permanent=1" in summary

    def test_empty_errors_no_extra_lines(self):
        result = EngineResult(
            success=True,
            workbook_path=Path("test.xlsx"),
            sections_completed=1,
            sections_total=1,
            tasks_completed=5,
            tasks_total=5,
            elapsed_seconds=5.0,
        )
        summary = result.summary()
        assert "Failed tasks" not in summary
        assert "Task errors" not in summary
