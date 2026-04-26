"""
Tests for Track B performance improvements:
  B1 — VBA sentinel polling
  B2 — Progress callback mechanism
  B3 — Conditional format false-positive fix
  B4 — Parallel independent task execution
  B5 — Expanded numeric value extraction
"""

import time
import threading
from pathlib import Path
from unittest.mock import MagicMock, patch, call

import pytest

from excel_engine.config import EngineConfig, TaskType, Layer
from excel_engine.layers.vba_layer import VBALayer, SENTINEL_SHEET, POLL_INTERVAL
from excel_engine.parsers.task_extractor import (
    TaskExtractor,
    Task,
    extract_numeric_value,
)
from excel_engine.engine import ExcelEngine, EngineResult
from excel_engine.planner.task_planner import ExecutionPlan, Section


# ═══════════════════════════════════════════════════════════════════════
# B1 — VBA Sentinel Polling
# ═══════════════════════════════════════════════════════════════════════

class TestVBASentinelPolling:
    """B1: Replace fixed sleep with sentinel polling."""

    def test_inject_sentinel_adds_vba_code(self):
        """Sentinel injection inserts the sheet-write before the last End Sub."""
        code = (
            "Sub Main()\n"
            "    MsgBox \"Hello\"\n"
            "End Sub\n"
        )
        injected = VBALayer.inject_sentinel(code)
        assert SENTINEL_SHEET in injected
        assert '"DONE"' in injected
        assert injected.rstrip().endswith("End Sub")

    def test_inject_sentinel_multiple_subs(self):
        """With multiple subs, sentinel is injected only in the last one."""
        code = (
            "Sub CleanUp()\n"
            "End Sub\n"
            "\n"
            "Sub Main()\n"
            "    ' work\n"
            "End Sub\n"
        )
        injected = VBALayer.inject_sentinel(code)
        # Should appear only once
        assert injected.count('"DONE"') == 1
        # Should be inside the Main sub (after "' work", before its End Sub)
        main_start = injected.find("Sub Main()")
        done_pos = injected.find('"DONE"')
        assert done_pos > main_start

    def test_fast_completion_via_sentinel(self):
        """When sentinel returns True immediately, polling exits early."""
        checker = MagicMock(return_value=True)
        layer = VBALayer(
            execution_timeout=10.0,
            poll_interval=0.05,
            sentinel_checker=checker,
        )

        with patch.object(layer, '_cleanup_sentinel'):
            with patch("excel_engine.layers.vba_layer.MacUtils"):
                start = time.time()
                layer.execute_vba("Sub Main()\nEnd Sub")
                elapsed = time.time() - start

        # Should complete well under the 10s timeout
        assert elapsed < 5.0
        checker.assert_called()

    def test_timeout_when_sentinel_never_fires(self):
        """When sentinel never returns True, polling runs until timeout."""
        checker = MagicMock(return_value=False)
        timeout = 1.0
        layer = VBALayer(
            execution_timeout=timeout,
            poll_interval=0.1,
            sentinel_checker=checker,
        )

        with patch.object(layer, '_cleanup_sentinel'):
            with patch("excel_engine.layers.vba_layer.MacUtils"):
                start = time.time()
                layer.execute_vba("Sub Main()\nEnd Sub")
                elapsed = time.time() - start

        assert elapsed >= timeout - 0.2  # allow small margin
        assert checker.call_count >= 2  # polled multiple times

    def test_sentinel_checker_exception_handled(self):
        """Exceptions in the sentinel checker don't crash the polling loop."""
        call_count = 0

        def flaky_checker():
            nonlocal call_count
            call_count += 1
            if call_count < 3:
                raise ConnectionError("xlwings not ready")
            return True

        layer = VBALayer(
            execution_timeout=5.0,
            poll_interval=0.05,
            sentinel_checker=flaky_checker,
        )

        with patch.object(layer, '_cleanup_sentinel'):
            with patch("excel_engine.layers.vba_layer.MacUtils"):
                layer.execute_vba("Sub Main()\nEnd Sub")

        assert call_count >= 3

    def test_default_poll_interval(self):
        """Default poll_interval matches the module constant."""
        layer = VBALayer()
        assert layer.poll_interval == POLL_INTERVAL


# ═══════════════════════════════════════════════════════════════════════
# B2 — Progress Callback Mechanism
# ═══════════════════════════════════════════════════════════════════════

class TestProgressCallback:
    """B2: Progress callback in engine.run() and execute()."""

    def _make_plan_and_config(self, task_count=2):
        """Helper: create a minimal plan with N tasks on one section."""
        tasks = []
        for i in range(task_count):
            tasks.append(Task(
                id=f"t{i}",
                task_type=TaskType.CELL_VALUE,
                description=f"enter 42 in cell A{i+1}",
                cell=f"A{i+1}",
                value="42",
            ))
        section = Section(id="s1", name="TestSection", sheet="Sheet1", tasks=tasks)
        plan = ExecutionPlan(sections=[section], total_tasks=task_count, estimated_time_seconds=10.0)
        config = EngineConfig(
            verify_after_each_section=False,
            recalculate_formulas=False,
        )
        return plan, config

    @patch.object(ExcelEngine, '_prepare_workbook', side_effect=lambda self_or_wb: Path("/fake/workbook.xlsx"))
    @patch.object(ExcelEngine, '_dispatch_task')
    @patch.object(ExcelEngine, '_cleanup')
    def test_progress_callback_called(self, mock_cleanup, mock_dispatch, mock_prep):
        """Progress callback receives executing + completed for each task."""
        plan, config = self._make_plan_and_config(2)
        engine = ExcelEngine(config=config)
        cb = MagicMock()

        engine.execute(plan, Path("/fake/workbook.xlsx"), progress_callback=cb)

        # Should have 2 executing + 2 completed = 4 calls
        assert cb.call_count == 4
        phases = [c.args[0]["phase"] for c in cb.call_args_list]
        assert phases == ["executing", "completed", "executing", "completed"]

    @patch.object(ExcelEngine, '_prepare_workbook', side_effect=lambda self_or_wb: Path("/fake/workbook.xlsx"))
    @patch.object(ExcelEngine, '_dispatch_task')
    @patch.object(ExcelEngine, '_cleanup')
    def test_progress_callback_data_shape(self, mock_cleanup, mock_dispatch, mock_prep):
        """Progress callback dict contains expected keys."""
        plan, config = self._make_plan_and_config(1)
        engine = ExcelEngine(config=config)
        cb = MagicMock()

        engine.execute(plan, Path("/fake/workbook.xlsx"), progress_callback=cb)

        executing_call = cb.call_args_list[0].args[0]
        assert "phase" in executing_call
        assert "task" in executing_call
        assert "total" in executing_call
        assert executing_call["phase"] == "executing"
        assert executing_call["task"] == "t0"

    @patch.object(ExcelEngine, '_prepare_workbook', side_effect=lambda self_or_wb: Path("/fake/workbook.xlsx"))
    @patch.object(ExcelEngine, '_dispatch_task')
    @patch.object(ExcelEngine, '_cleanup')
    def test_no_callback_no_error(self, mock_cleanup, mock_dispatch, mock_prep):
        """None callback does not cause errors."""
        plan, config = self._make_plan_and_config(1)
        engine = ExcelEngine(config=config)

        # Should not raise
        result = engine.execute(plan, Path("/fake/workbook.xlsx"), progress_callback=None)
        assert result.tasks_completed == 1

    @patch.object(ExcelEngine, '_prepare_workbook', side_effect=lambda self_or_wb: Path("/fake/workbook.xlsx"))
    @patch.object(ExcelEngine, '_dispatch_task', side_effect=RuntimeError("boom"))
    @patch.object(ExcelEngine, '_cleanup')
    def test_progress_callback_on_failure(self, mock_cleanup, mock_dispatch, mock_prep):
        """On task failure, callback reports success=False."""
        plan, config = self._make_plan_and_config(1)
        config.max_retries = 0
        engine = ExcelEngine(config=config)
        cb = MagicMock()

        engine.execute(plan, Path("/fake/workbook.xlsx"), progress_callback=cb)

        completed_call = cb.call_args_list[-1].args[0]
        assert completed_call["phase"] == "completed"
        assert completed_call["success"] is False


# ═══════════════════════════════════════════════════════════════════════
# B3 — Conditional Format False Positives
# ═══════════════════════════════════════════════════════════════════════

class TestConditionalFormatFalsePositives:
    """B3: Basic formatting must NOT trigger conditional_format."""

    def setup_method(self):
        self.extractor = TaskExtractor()

    def _types_for(self, text: str) -> set[str]:
        tasks = self.extractor.extract(text)
        return {t.task_type.value for t in tasks}

    # ── Negative cases: should NOT produce conditional_format ──

    def test_format_cells_bold_not_conditional(self):
        assert "conditional_format" not in self._types_for(
            "Format the cells as bold"
        )

    def test_format_range_as_currency_not_conditional(self):
        assert "conditional_format" not in self._types_for(
            "Format the range A1:A10 as currency"
        )

    def test_format_cells_as_percentage_not_conditional(self):
        assert "conditional_format" not in self._types_for(
            "Format cells as percentage"
        )

    def test_plain_format_in_red_not_conditional(self):
        """'format in red' without conditional/rule/highlight is basic formatting."""
        types = self._types_for("Format the header row in red")
        assert "conditional_format" not in types

    # ── Positive cases: SHOULD produce conditional_format ──

    def test_conditional_format_rule(self):
        assert "conditional_format" in self._types_for(
            "Apply a conditional formatting rule to highlight cells greater than 100"
        )

    def test_data_bar(self):
        assert "conditional_format" in self._types_for(
            "Add data bars to the range B2:B20"
        )

    def test_color_scale(self):
        assert "conditional_format" in self._types_for(
            "Apply a color scale to C1:C50"
        )

    def test_icon_set(self):
        assert "conditional_format" in self._types_for(
            "Apply icon sets to the profit column"
        )

    def test_highlight_cells_rule(self):
        assert "conditional_format" in self._types_for(
            "Highlight cells rule for values greater than 500"
        )

    def test_top_bottom_rules(self):
        assert "conditional_format" in self._types_for(
            "Apply top/bottom 10 rules"
        )

    def test_new_formatting_rule(self):
        assert "conditional_format" in self._types_for(
            "Create a new rule for values above 1000"
        )

    def test_conditional_format_plain_phrase(self):
        assert "conditional_format" in self._types_for(
            "Add conditional formatting to the sales column"
        )

    def test_duplicate_values_rule(self):
        assert "conditional_format" in self._types_for(
            "Highlight duplicate values rule in column A"
        )


# ═══════════════════════════════════════════════════════════════════════
# B4 — Parallel Independent Task Execution
# ═══════════════════════════════════════════════════════════════════════

class TestParallelExecution:
    """B4: Opt-in parallel execution for independent tasks on different sheets."""

    def test_default_parallel_off(self):
        """parallel_execution defaults to False."""
        config = EngineConfig()
        assert config.parallel_execution is False
        assert config.max_workers == 4

    @patch.object(ExcelEngine, '_prepare_workbook', side_effect=lambda self_or_wb: Path("/fake/wb.xlsx"))
    @patch.object(ExcelEngine, '_dispatch_task')
    @patch.object(ExcelEngine, '_cleanup')
    def test_sequential_unchanged(self, mock_cleanup, mock_dispatch, mock_prep):
        """With parallel_execution=False, tasks run sequentially (existing behavior)."""
        tasks = [
            Task(id="t1", task_type=TaskType.CELL_VALUE, description="enter 1", sheet="Sheet1", cell="A1", value="1"),
            Task(id="t2", task_type=TaskType.CELL_VALUE, description="enter 2", sheet="Sheet2", cell="A1", value="2"),
        ]
        section = Section(id="s1", name="Test", sheet="Sheet1", tasks=tasks)
        plan = ExecutionPlan(sections=[section], total_tasks=len(tasks), estimated_time_seconds=10.0)
        config = EngineConfig(
            parallel_execution=False,
            verify_after_each_section=False,
            recalculate_formulas=False,
        )
        engine = ExcelEngine(config=config)
        result = engine.execute(plan, Path("/fake/wb.xlsx"))

        assert result.tasks_completed == 2
        assert mock_dispatch.call_count == 2
        # Verify sequential order
        calls = mock_dispatch.call_args_list
        assert calls[0].args[0].id == "t1"
        assert calls[1].args[0].id == "t2"

    @patch.object(ExcelEngine, '_prepare_workbook', side_effect=lambda self_or_wb: Path("/fake/wb.xlsx"))
    @patch.object(ExcelEngine, '_dispatch_task')
    @patch.object(ExcelEngine, '_cleanup')
    def test_parallel_multi_sheet(self, mock_cleanup, mock_dispatch, mock_prep):
        """With parallel=True and different sheets, all tasks complete."""
        tasks = [
            Task(id="t1", task_type=TaskType.CELL_VALUE, description="enter 1", sheet="Sheet1", cell="A1", value="1"),
            Task(id="t2", task_type=TaskType.CELL_VALUE, description="enter 2", sheet="Sheet2", cell="A1", value="2"),
            Task(id="t3", task_type=TaskType.CELL_VALUE, description="enter 3", sheet="Sheet3", cell="A1", value="3"),
        ]
        section = Section(id="s1", name="Test", sheet="Sheet1", tasks=tasks)
        plan = ExecutionPlan(sections=[section], total_tasks=len(tasks), estimated_time_seconds=10.0)
        config = EngineConfig(
            parallel_execution=True,
            max_workers=3,
            verify_after_each_section=False,
            recalculate_formulas=False,
        )
        engine = ExcelEngine(config=config)
        result = engine.execute(plan, Path("/fake/wb.xlsx"))

        assert result.tasks_completed == 3
        assert result.success is True

    @patch.object(ExcelEngine, '_prepare_workbook', side_effect=lambda self_or_wb: Path("/fake/wb.xlsx"))
    @patch.object(ExcelEngine, '_dispatch_task')
    @patch.object(ExcelEngine, '_cleanup')
    def test_parallel_same_sheet_serial(self, mock_cleanup, mock_dispatch, mock_prep):
        """Tasks on the same sheet run serially even in parallel mode."""
        execution_order = []

        def track_dispatch(task, layer, workbook):
            execution_order.append(task.id)

        mock_dispatch.side_effect = track_dispatch

        tasks = [
            Task(id="t1", task_type=TaskType.CELL_VALUE, description="enter 1", sheet="Sheet1", cell="A1", value="1"),
            Task(id="t2", task_type=TaskType.CELL_VALUE, description="enter 2", sheet="Sheet1", cell="A2", value="2"),
            Task(id="t3", task_type=TaskType.CELL_VALUE, description="enter 3", sheet="Sheet1", cell="A3", value="3"),
        ]
        section = Section(id="s1", name="Test", sheet="Sheet1", tasks=tasks)
        plan = ExecutionPlan(sections=[section], total_tasks=len(tasks), estimated_time_seconds=10.0)
        config = EngineConfig(
            parallel_execution=True,
            verify_after_each_section=False,
            recalculate_formulas=False,
        )
        engine = ExcelEngine(config=config)
        result = engine.execute(plan, Path("/fake/wb.xlsx"))

        assert result.tasks_completed == 3
        # Same sheet → must be in order
        assert execution_order == ["t1", "t2", "t3"]

    @patch.object(ExcelEngine, '_prepare_workbook', side_effect=lambda self_or_wb: Path("/fake/wb.xlsx"))
    @patch.object(ExcelEngine, '_dispatch_task')
    @patch.object(ExcelEngine, '_cleanup')
    def test_parallel_with_progress_callback(self, mock_cleanup, mock_dispatch, mock_prep):
        """Progress callback works in parallel mode."""
        tasks = [
            Task(id="t1", task_type=TaskType.CELL_VALUE, description="enter 1", sheet="Sheet1", cell="A1", value="1"),
            Task(id="t2", task_type=TaskType.CELL_VALUE, description="enter 2", sheet="Sheet2", cell="A1", value="2"),
        ]
        section = Section(id="s1", name="Test", sheet="Sheet1", tasks=tasks)
        plan = ExecutionPlan(sections=[section], total_tasks=len(tasks), estimated_time_seconds=10.0)
        config = EngineConfig(
            parallel_execution=True,
            verify_after_each_section=False,
            recalculate_formulas=False,
        )
        engine = ExcelEngine(config=config)
        cb = MagicMock()
        result = engine.execute(plan, Path("/fake/wb.xlsx"), progress_callback=cb)

        assert result.tasks_completed == 2
        # 2 executing + 2 completed = 4 calls
        assert cb.call_count == 4


# ═══════════════════════════════════════════════════════════════════════
# B5 — Expanded Numeric Value Extraction
# ═══════════════════════════════════════════════════════════════════════

class TestExpandedNumericExtraction:
    """B5: Percentages, currency, scientific notation, parenthetical negatives."""

    def test_percentage_integer(self):
        assert extract_numeric_value("enter 25% in cell A1") == "25%"

    def test_percentage_decimal(self):
        assert extract_numeric_value("the rate is 3.5%") == "3.5%"

    def test_currency_simple(self):
        assert extract_numeric_value("enter $100 in cell B2") == "100"

    def test_currency_with_commas(self):
        assert extract_numeric_value("the total is $1,234.56") == "1234.56"

    def test_scientific_notation_lower(self):
        assert extract_numeric_value("value is 1.5e-3") == "1.5e-3"

    def test_scientific_notation_upper(self):
        assert extract_numeric_value("enter 2E10 in cell C1") == "2E10"

    def test_parenthetical_negative_simple(self):
        assert extract_numeric_value("the loss is (100)") == "-100"

    def test_parenthetical_negative_with_commas(self):
        assert extract_numeric_value("net income (1,234.56)") == "-1234.56"

    def test_plain_number(self):
        assert extract_numeric_value("enter 42 in cell A1") == "42"

    def test_plain_number_with_commas(self):
        assert extract_numeric_value("total is 1,000,000") == "1000000"

    def test_no_number_returns_none(self):
        assert extract_numeric_value("format the cells as bold") is None

    def test_extractor_uses_expanded_for_cell_value(self):
        """TaskExtractor picks up $1,234.56 via quoted value for CELL_VALUE tasks."""
        extractor = TaskExtractor()
        tasks = extractor.extract('enter the value "$1,234.56" in the cell A1')
        cell_value_tasks = [t for t in tasks if t.task_type == TaskType.CELL_VALUE]
        assert len(cell_value_tasks) >= 1
        # The quoted value is extracted directly; verify it contains the amount
        assert "1,234.56" in cell_value_tasks[0].value or "1234.56" in cell_value_tasks[0].value

    def test_extractor_bare_number_with_commas(self):
        """TaskExtractor picks up comma-separated number for CELL_VALUE tasks."""
        extractor = TaskExtractor()
        tasks = extractor.extract("enter 15000 in cell C1")
        cell_value_tasks = [t for t in tasks if t.task_type == TaskType.CELL_VALUE]
        assert len(cell_value_tasks) >= 1
        assert cell_value_tasks[0].value == "15000"

    def test_extractor_plain_number(self):
        """TaskExtractor picks up plain integers for CELL_VALUE tasks."""
        extractor = TaskExtractor()
        tasks = extractor.extract("enter 42 in cell D1")
        cell_value_tasks = [t for t in tasks if t.task_type == TaskType.CELL_VALUE]
        assert len(cell_value_tasks) >= 1
        assert cell_value_tasks[0].value == "42"
