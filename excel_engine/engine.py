"""
Engine — Main orchestrator implementing Engine Protocol v2.0.

Coordinates all 6 layers, parsers, planner, and verifier to autonomously
complete Excel worksheet assignments.

Protocol v2.0:
  1. SCAN  — Read instructions, identify ALL tasks (2 min max)
  2. GROUP — Cluster tasks into logical sections (by sheet/dependency)
  3. For each section:
     a. Quick-plan: What needs doing, what order (30 sec)
     b. Verify plan is sound (no conflicts/missing prereqs)
     c. Execute aggressively
     d. Verify completion
     e. Move to next section
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

from excel_engine.config import EngineConfig, TaskType, Layer, TASK_LAYER_MAP
from excel_engine.layers.openpyxl_layer import OpenpyxlLayer
from excel_engine.layers.xlwings_layer import XlwingsLayer
from excel_engine.layers.applescript_layer import AppleScriptLayer
from excel_engine.layers.system_events import SystemEventsLayer
from excel_engine.layers.vba_layer import VBALayer
from excel_engine.layers.pyautogui_layer import PyAutoGUILayer
from excel_engine.parsers.instruction_parser import InstructionParser
from excel_engine.parsers.task_extractor import Task, TaskExtractor
from excel_engine.planner.task_planner import TaskPlanner, ExecutionPlan, Section
from excel_engine.verifier.workbook_verifier import WorkbookVerifier, SectionVerification
from excel_engine.utils.path_handler import PathHandler
from excel_engine.utils.mac_utils import MacUtils

logger = logging.getLogger(__name__)


@dataclass
class EngineResult:
    """Result of running the engine on an assignment."""
    success: bool
    workbook_path: Path
    sections_completed: int
    sections_total: int
    tasks_completed: int
    tasks_total: int
    verifications: list[SectionVerification] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    elapsed_seconds: float = 0.0

    def summary(self) -> str:
        status = "✓ SUCCESS" if self.success else "✗ FAILED"
        lines = [
            f"{status}: {self.workbook_path.name}",
            f"  Sections: {self.sections_completed}/{self.sections_total}",
            f"  Tasks: {self.tasks_completed}/{self.tasks_total}",
            f"  Time: {self.elapsed_seconds:.1f}s",
        ]
        if self.errors:
            lines.append(f"  Errors ({len(self.errors)}):")
            for err in self.errors[:5]:
                lines.append(f"    - {err}")
        return "\n".join(lines)


class ExcelEngine:
    """
    Main orchestrator — coordinates all layers to complete Excel assignments.

    Usage:
        engine = ExcelEngine()
        result = engine.run(
            workbook=Path("assignment.xlsx"),
            instructions=Path("instructions.docx"),
        )
        print(result.summary())
    """

    def __init__(self, config: Optional[EngineConfig] = None) -> None:
        self.config = config or EngineConfig()

        # Initialize layers
        self.openpyxl = OpenpyxlLayer()
        self.xlwings = XlwingsLayer()
        self.applescript = AppleScriptLayer(timeout=self.config.applescript_timeout)
        self.system_events = SystemEventsLayer(delay=self.config.ui_delay_between_actions)
        self.vba = VBALayer(
            execution_timeout=self.config.vba_execution_timeout,
            split_threshold=self.config.vba_split_threshold,
        )
        self.pyautogui = PyAutoGUILayer(retina=self.config.retina_display)

        # Initialize components
        self.parser = InstructionParser()
        self.extractor = TaskExtractor()
        self.planner = TaskPlanner(config=self.config)
        self.verifier = WorkbookVerifier()
        self.path_handler = PathHandler(desktop_path=self.config.desktop_path)

        # Layer lookup
        self._layers = {
            Layer.OPENPYXL: self.openpyxl,
            Layer.XLWINGS: self.xlwings,
            Layer.APPLESCRIPT: self.applescript,
            Layer.SYSTEM_EVENTS: self.system_events,
            Layer.VBA: self.vba,
            Layer.PYAUTOGUI: self.pyautogui,
        }

    # ── Main Entry Point ──

    def run(
        self,
        workbook: Path,
        instructions: Optional[Path] = None,
        instruction_text: Optional[str] = None,
        tasks: Optional[list[Task]] = None,
    ) -> EngineResult:
        """
        Run the engine on a workbook with instructions.

        Provide EITHER:
          - instructions: path to instruction file (.docx, .rtfd, .pdf, .txt)
          - instruction_text: raw instruction text
          - tasks: pre-extracted Task list
        """
        start_time = time.time()
        workbook = Path(workbook).resolve()

        result = EngineResult(
            success=False,
            workbook_path=workbook,
            sections_completed=0,
            sections_total=0,
            tasks_completed=0,
            tasks_total=0,
        )

        try:
            # ── Phase 1: SCAN ──
            logger.info("=" * 60)
            logger.info("PHASE 1: SCAN — Parsing instructions")
            logger.info("=" * 60)

            if tasks is None:
                if instructions:
                    text = self.parser.parse(instructions)
                elif instruction_text:
                    text = instruction_text
                else:
                    raise ValueError("Provide instructions, instruction_text, or tasks")

                tasks = self.extractor.extract(text)

            logger.info("Extracted %d tasks", len(tasks))

            # ── Phase 2: GROUP ──
            logger.info("=" * 60)
            logger.info("PHASE 2: GROUP — Planning execution")
            logger.info("=" * 60)

            plan = self.planner.plan(tasks)
            result.sections_total = plan.section_count
            result.tasks_total = plan.total_tasks

            logger.info("\n%s", plan.summary())

            # ── Phase 3: EXECUTE ──
            logger.info("=" * 60)
            logger.info("PHASE 3: EXECUTE — Running %d sections", plan.section_count)
            logger.info("=" * 60)

            for section in plan.sections:
                section_ok = self._execute_section(section, workbook, result)
                if section_ok:
                    result.sections_completed += 1

            result.success = (
                result.sections_completed == result.sections_total
                and result.tasks_completed == result.tasks_total
            )

        except Exception as e:
            logger.exception("Engine failed: %s", e)
            result.errors.append(str(e))

        finally:
            result.elapsed_seconds = time.time() - start_time
            self._cleanup()

        logger.info("\n%s", result.summary())
        return result

    # ── Section Execution ──

    def _execute_section(
        self, section: Section, workbook: Path, result: EngineResult
    ) -> bool:
        """Execute all tasks in a section."""
        logger.info("-" * 40)
        logger.info("Section %s: %s (%d tasks)", section.id, section.name, section.task_count)
        logger.info("-" * 40)

        section_ok = True

        for task in section.tasks:
            task_ok = self._execute_task(task, workbook)
            if task_ok:
                task.completed = True
                result.tasks_completed += 1
            else:
                section_ok = False
                result.errors.append(
                    f"Task {task.id} ({task.task_type.value}) failed"
                )

        # Verify section if configured
        if self.config.verify_after_each_section:
            try:
                self._save_workbook(workbook)
                self.verifier.load(workbook)
                verification = self.verifier.verify_section(section.id, section.tasks)
                result.verifications.append(verification)
                self.verifier.close()

                if not verification.all_passed:
                    logger.warning(
                        "Verification issues in section %s: %d failed",
                        section.id, verification.fail_count,
                    )
            except Exception as e:
                logger.warning("Verification error for section %s: %s", section.id, e)

        section.completed = section_ok
        return section_ok

    # ── Task Execution ──

    def _execute_task(self, task: Task, workbook: Path) -> bool:
        """Execute a single task using the appropriate layer(s)."""
        layers_to_try = self.config.get_layers_for_task(task.task_type)

        if not layers_to_try:
            logger.warning("No layer can handle task type: %s", task.task_type.value)
            return False

        for layer_enum in layers_to_try:
            try:
                logger.info(
                    "  Task %s: %s via Layer %s",
                    task.id, task.task_type.value, layer_enum.name,
                )
                self._dispatch_task(task, layer_enum, workbook)
                return True
            except Exception as e:
                logger.warning(
                    "  Layer %s failed for %s: %s — trying next",
                    layer_enum.name, task.id, e,
                )

        logger.error("All layers failed for task %s", task.id)
        return False

    def _dispatch_task(
        self, task: Task, layer: Layer, workbook: Path
    ) -> None:
        """Dispatch a task to a specific layer."""
        dispatch_map = {
            Layer.OPENPYXL: self._exec_openpyxl,
            Layer.XLWINGS: self._exec_xlwings,
            Layer.APPLESCRIPT: self._exec_applescript,
            Layer.SYSTEM_EVENTS: self._exec_system_events,
            Layer.VBA: self._exec_vba,
            Layer.PYAUTOGUI: self._exec_pyautogui,
        }

        handler = dispatch_map.get(layer)
        if handler:
            handler(task, workbook)
        else:
            raise ValueError(f"Unknown layer: {layer}")

    # ── Layer Dispatchers ──

    def _exec_openpyxl(self, task: Task, workbook: Path) -> None:
        """Execute a task via Layer 1 (openpyxl)."""
        if not self.openpyxl.wb:
            self.openpyxl.open(workbook)

        tt = task.task_type

        if tt == TaskType.FORMULA:
            if task.cell and task.formula:
                self.openpyxl.set_formula(task.cell, task.formula, sheet=task.sheet)

        elif tt == TaskType.CELL_VALUE:
            if task.cell and task.value:
                self.openpyxl.set_value(task.cell, task.value, sheet=task.sheet)

        elif tt == TaskType.TABLE_CREATE:
            name = task.params.get("name", "Table1")
            ref = task.range or task.params.get("ref", "A1:A1")
            style = task.style or task.params.get("style", "TableStyleMedium5")
            self.openpyxl.create_table(name, ref, style=style, sheet=task.sheet)

        elif tt == TaskType.NUMBER_FORMAT:
            ref = task.range or task.cell
            fmt = task.params.get("format", "$#,##0.00")
            if ref:
                self.openpyxl.set_number_format(ref, fmt, sheet=task.sheet)

        elif tt == TaskType.FONT:
            ref = task.range or task.cell
            if ref:
                self.openpyxl.set_font(
                    ref, sheet=task.sheet,
                    bold=task.params.get("bold", False),
                    italic=task.params.get("italic", False),
                    size=task.params.get("size", 11),
                    color=task.params.get("color", "000000"),
                    name=task.params.get("font_name", "Calibri"),
                )

        elif tt == TaskType.FILL:
            ref = task.range or task.cell
            color = task.params.get("color", "FFFF00")
            if ref:
                self.openpyxl.set_fill(ref, color, sheet=task.sheet)

        elif tt == TaskType.ALIGNMENT:
            ref = task.range or task.cell
            if ref:
                self.openpyxl.set_alignment(
                    ref, sheet=task.sheet,
                    horizontal=task.params.get("horizontal", "general"),
                    vertical=task.params.get("vertical", "bottom"),
                    wrap_text=task.params.get("wrap_text", False),
                )

        elif tt == TaskType.BORDER:
            ref = task.range or task.cell
            if ref:
                self.openpyxl.set_border(
                    ref, sheet=task.sheet,
                    style=task.params.get("style", "thin"),
                )

        elif tt == TaskType.COLUMN_WIDTH:
            col = task.params.get("column", "A")
            width = task.params.get("width", 12)
            self.openpyxl.set_column_width(col, width, sheet=task.sheet)

        elif tt == TaskType.CONDITIONAL_FORMAT:
            ref = task.range or "A1:A1"
            operator = task.params.get("operator", "greaterThan")
            formula = task.params.get("formula", ["0"])
            self.openpyxl.add_conditional_format_cell_is(
                ref, operator, formula, sheet=task.sheet,
            )

        elif tt == TaskType.CHART_BAR:
            self.openpyxl.add_bar_chart(
                sheet=task.sheet,
                title=task.params.get("title", ""),
                data_range=task.params.get("data_range", task.range or "B1:B10"),
                cats_range=task.params.get("cats_range"),
                anchor=task.params.get("anchor", "E2"),
            )

        elif tt == TaskType.CHART_LINE:
            self.openpyxl.add_line_chart(
                sheet=task.sheet,
                title=task.params.get("title", ""),
                data_range=task.params.get("data_range", task.range or "B1:B10"),
                cats_range=task.params.get("cats_range"),
                anchor=task.params.get("anchor", "E2"),
            )

        elif tt == TaskType.CHART_PIE:
            self.openpyxl.add_pie_chart(
                sheet=task.sheet,
                title=task.params.get("title", ""),
                data_range=task.params.get("data_range", task.range or "B1:B10"),
                cats_range=task.params.get("cats_range"),
                anchor=task.params.get("anchor", "E2"),
            )

        elif tt == TaskType.NAMED_RANGE:
            name = task.params.get("name", "NamedRange1")
            sheet = task.sheet or self.openpyxl.wb.active.title
            range_str = task.range or "$A$1"
            self.openpyxl.create_named_range(name, sheet, range_str)

        elif tt == TaskType.DATA_VALIDATION:
            ref = task.range or task.cell or "A1"
            self.openpyxl.add_data_validation(
                ref, sheet=task.sheet,
                validation_type=task.params.get("type", "list"),
                formula1=task.params.get("formula1"),
            )

        elif tt == TaskType.FREEZE_PANES:
            cell = task.cell or "A2"
            self.openpyxl.freeze_panes(cell, sheet=task.sheet)

        elif tt == TaskType.AUTOFILTER:
            ref = task.range or task.params.get("ref", "A1:A1")
            self.openpyxl.set_autofilter(ref, sheet=task.sheet)

        elif tt == TaskType.SHEET_CREATE:
            name = task.params.get("name", task.sheet or "NewSheet")
            self.openpyxl.create_sheet(name)

        elif tt == TaskType.SHEET_RENAME:
            old = task.params.get("old_name", "Sheet1")
            new = task.params.get("new_name", task.sheet or "Renamed")
            self.openpyxl.rename_sheet(old, new)

        elif tt == TaskType.MERGE_CELLS:
            ref = task.range
            if ref:
                self.openpyxl.merge_cells(ref, sheet=task.sheet)

        elif tt == TaskType.PRINT_SETTINGS:
            if task.params.get("landscape"):
                self.openpyxl.set_page_orientation(True, sheet=task.sheet)
            if task.params.get("print_area"):
                self.openpyxl.set_print_area(task.params["print_area"], sheet=task.sheet)

        else:
            raise NotImplementedError(
                f"openpyxl handler not implemented for {tt.value}"
            )

    def _exec_xlwings(self, task: Task, workbook: Path) -> None:
        """Execute a task via Layer 2 (xlwings)."""
        safe_path = self.path_handler.safe_copy_for_xlwings(workbook)

        if not self.xlwings._wb:
            MacUtils.launch_excel()
            MacUtils.open_workbook_in_excel(safe_path)
            MacUtils.wait_for_excel_ready()
            self.xlwings.connect(safe_path)

        tt = task.task_type

        if tt == TaskType.FORMULA or tt == TaskType.CALCULATED_COLUMN:
            if task.cell and task.formula:
                self.xlwings.set_formula(task.cell, task.formula, sheet=task.sheet)

        elif tt == TaskType.CELL_VALUE:
            if task.cell and task.value:
                self.xlwings.write_cell(task.cell, task.value, sheet=task.sheet)

        elif tt == TaskType.SUBTOTAL:
            range_str = task.range or task.params.get("range", "A1:A1")
            self.xlwings.add_subtotal(
                range_str,
                group_by=task.params.get("group_by", 1),
                function=task.params.get("function", -4157),
                total_list=task.params.get("total_list", [2]),
                sheet=task.sheet,
            )

        elif tt == TaskType.SPLIT_PANES:
            row = task.params.get("row", 2)
            col = task.params.get("col", 1)
            self.xlwings.split_panes(row, col, sheet=task.sheet)

        elif tt == TaskType.ADVANCED_FILTER:
            self.xlwings.advanced_filter(
                list_range=task.params.get("list_range", task.range or "A1:A1"),
                criteria_range=task.params.get("criteria_range", ""),
                copy_to_range=task.params.get("copy_to_range"),
                unique=task.params.get("unique", False),
                sheet=task.sheet,
            )

        elif tt == TaskType.NUMBER_FORMAT:
            ref = task.range or task.cell
            if ref:
                self.xlwings.set_number_format(
                    ref, task.params.get("format", "$#,##0.00"), sheet=task.sheet,
                )

        elif tt == TaskType.COLUMN_WIDTH:
            col = task.params.get("column", "A")
            width = task.params.get("width", 12)
            self.xlwings.set_column_width(col, width, sheet=task.sheet)

        elif tt == TaskType.TABLE_STYLE:
            table_name = task.params.get("table_name", "Table1")
            style = task.style or task.params.get("style", "TableStyleMedium5")
            sheet = task.sheet or "Sheet1"
            self.xlwings.set_table_style(table_name, style, sheet)

        elif tt == TaskType.SAVE:
            self.xlwings.save()

        else:
            raise NotImplementedError(
                f"xlwings handler not implemented for {tt.value}"
            )

        # Copy back if we used a desktop copy
        if safe_path != workbook:
            self.path_handler.copy_back_from_desktop(safe_path, workbook)

    def _exec_applescript(self, task: Task, workbook: Path) -> None:
        """Execute a task via Layer 3 (AppleScript)."""
        tt = task.task_type

        if tt == TaskType.FORMULA or tt == TaskType.CALCULATED_COLUMN:
            if task.cell and task.formula:
                self.applescript.set_cell_formula(
                    task.cell, task.formula, sheet=task.sheet,
                )

        elif tt == TaskType.CELL_VALUE:
            if task.cell:
                self.applescript.set_cell_value(
                    task.cell, str(task.value or ""), sheet=task.sheet,
                )

        elif tt == TaskType.SORT:
            range_str = task.range or task.params.get("range", "A1:A1")
            keys = task.params.get("keys", [])
            self.applescript.sort_range(range_str, keys, sheet=task.sheet)

        elif tt == TaskType.AUTOFILTER:
            range_str = task.range or task.params.get("range", "A1:A1")
            field_num = task.params.get("field", 1)
            criteria = task.params.get("criteria", "*")
            self.applescript.set_autofilter(
                range_str, field_num, criteria, sheet=task.sheet,
            )

        elif tt == TaskType.FREEZE_PANES:
            cell = task.cell or "A2"
            self.applescript.freeze_panes(cell, sheet=task.sheet)

        elif tt == TaskType.SAVE:
            self.applescript.save()

        elif tt == TaskType.SAVE_AS:
            path = task.params.get("path", str(self.config.desktop_path))
            filename = task.params.get("filename", "output")
            self.applescript.save_as_xlsx(path, filename)

        else:
            raise NotImplementedError(
                f"AppleScript handler not implemented for {tt.value}"
            )

    def _exec_system_events(self, task: Task, workbook: Path) -> None:
        """Execute a task via Layer 4 (System Events)."""
        tt = task.task_type

        if tt == TaskType.SLICER:
            fields = task.params.get("fields", [])
            self.system_events.insert_slicer(fields)
            # Configure if specified
            if task.params.get("columns") or task.params.get("caption"):
                self.system_events.configure_slicer(
                    columns=task.params.get("columns"),
                    caption=task.params.get("caption"),
                )

        elif tt == TaskType.CHART_HISTOGRAM:
            self.system_events.insert_histogram(
                data_range=task.range or task.params.get("data_range"),
            )
            # Configure bins if specified
            if any(k in task.params for k in ("bin_width", "overflow", "underflow")):
                self.system_events.configure_histogram_bins(
                    bin_width=task.params.get("bin_width"),
                    overflow=task.params.get("overflow"),
                    underflow=task.params.get("underflow"),
                )

        elif tt == TaskType.TABLE_STYLE:
            style = task.style or task.params.get("style", "TableStyleMedium5")
            self.system_events.apply_table_style_via_ribbon(style)

        else:
            raise NotImplementedError(
                f"System Events handler not implemented for {tt.value}"
            )

    def _exec_vba(self, task: Task, workbook: Path) -> None:
        """Execute a task via Layer 5 (VBA)."""
        tt = task.task_type

        if tt == TaskType.PIVOT_TABLE:
            self.vba.create_pivot_table(
                source_sheet=task.params.get("source_sheet", task.sheet or "Sheet1"),
                source_range=task.params.get("source_range", task.range or "A1:A1"),
                dest_sheet=task.params.get("dest_sheet", "PivotSheet"),
                dest_cell=task.params.get("dest_cell", "A3"),
                table_name=task.params.get("table_name", "PivotTable1"),
                row_fields=task.params.get("row_fields"),
                column_fields=task.params.get("column_fields"),
                value_fields=task.params.get("value_fields"),
                filter_fields=task.params.get("filter_fields"),
            )

        elif tt == TaskType.PIVOT_CHART:
            self.vba.create_pivot_chart(
                pivot_table_name=task.params.get("pivot_table_name", "PivotTable1"),
                chart_type=task.params.get("chart_type", "xlColumnClustered"),
                chart_title=task.params.get("chart_title", ""),
                dest_sheet=task.params.get("dest_sheet"),
            )

        else:
            raise NotImplementedError(
                f"VBA handler not implemented for {tt.value}"
            )

    def _exec_pyautogui(self, task: Task, workbook: Path) -> None:
        """Execute a task via Layer 6 (PyAutoGUI) — last resort."""
        raise NotImplementedError(
            f"PyAutoGUI execution for {task.task_type.value} requires "
            f"screen coordinates — implement per assignment"
        )

    # ── Utility Methods ──

    def _save_workbook(self, workbook: Path) -> None:
        """Save via the best available method."""
        if self.openpyxl.wb:
            self.openpyxl.save(workbook)
        elif self.xlwings._wb:
            self.xlwings.save()
        else:
            try:
                self.applescript.save()
            except Exception:
                logger.warning("Could not save workbook")

    def _cleanup(self) -> None:
        """Clean up all layer connections."""
        try:
            self.openpyxl.close()
        except Exception:
            pass
        try:
            self.xlwings.disconnect()
        except Exception:
            pass
        try:
            self.verifier.close()
        except Exception:
            pass

    # ── Convenience Methods ──

    def scan(self, instructions: Path) -> list[Task]:
        """Phase 1: Scan instructions and extract tasks."""
        text = self.parser.parse(instructions)
        return self.extractor.extract(text)

    def plan(self, tasks: list[Task]) -> ExecutionPlan:
        """Phase 2: Create an execution plan from tasks."""
        return self.planner.plan(tasks)

    def verify(self, workbook: Path, tasks: list[Task]) -> SectionVerification:
        """Verify a set of tasks against a workbook."""
        self.verifier.load(workbook)
        result = self.verifier.verify_section("manual", tasks)
        self.verifier.close()
        return result
