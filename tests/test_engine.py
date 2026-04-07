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
