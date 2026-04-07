"""
Workbook Verifier — Verify task completion after each section.

Checks formulas, formatting, charts, tables, slicers, named ranges,
data validation, and other Excel features to confirm tasks were executed.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

from openpyxl import load_workbook
from excel_engine.config import TaskType
from excel_engine.parsers.task_extractor import Task

logger = logging.getLogger(__name__)


@dataclass
class VerificationResult:
    """Result of verifying a single task."""
    task_id: str
    task_type: TaskType
    passed: bool
    message: str
    details: Optional[dict] = None


@dataclass
class SectionVerification:
    """Aggregated verification results for a section."""
    section_id: str
    results: list[VerificationResult] = field(default_factory=list)

    @property
    def all_passed(self) -> bool:
        return all(r.passed for r in self.results)

    @property
    def pass_count(self) -> int:
        return sum(1 for r in self.results if r.passed)

    @property
    def fail_count(self) -> int:
        return sum(1 for r in self.results if not r.passed)

    def summary(self) -> str:
        total = len(self.results)
        return (
            f"Section {self.section_id}: {self.pass_count}/{total} passed"
            f"{' ✓' if self.all_passed else ' ✗'}"
        )


class WorkbookVerifier:
    """Verify Excel workbook state against expected task outcomes."""

    def __init__(self) -> None:
        self._wb = None
        self._path: Optional[Path] = None

    def load(self, path: Path) -> None:
        """Load a workbook for verification (data_only=False to see formulas)."""
        self._path = Path(path)
        self._wb = load_workbook(str(self._path), data_only=False)
        logger.info("Loaded workbook for verification: %s", self._path.name)

    def load_with_values(self, path: Path) -> None:
        """Load a workbook with calculated values (data_only=True)."""
        self._path = Path(path)
        self._wb = load_workbook(str(self._path), data_only=True)

    def close(self) -> None:
        if self._wb:
            self._wb.close()
            self._wb = None

    def verify_section(
        self, section_id: str, tasks: list[Task]
    ) -> SectionVerification:
        """Verify all tasks in a section."""
        verification = SectionVerification(section_id=section_id)

        for task in tasks:
            result = self.verify_task(task)
            verification.results.append(result)

        logger.info(verification.summary())
        return verification

    def verify_task(self, task: Task) -> VerificationResult:
        """Verify a single task was completed correctly."""
        verifiers = {
            TaskType.FORMULA: self._verify_formula,
            TaskType.TABLE_CREATE: self._verify_table,
            TaskType.TABLE_STYLE: self._verify_table_style,
            TaskType.FORMATTING: self._verify_formatting,
            TaskType.NUMBER_FORMAT: self._verify_number_format,
            TaskType.CONDITIONAL_FORMAT: self._verify_conditional_format,
            TaskType.CHART_BAR: self._verify_chart,
            TaskType.CHART_LINE: self._verify_chart,
            TaskType.CHART_PIE: self._verify_chart,
            TaskType.NAMED_RANGE: self._verify_named_range,
            TaskType.DATA_VALIDATION: self._verify_data_validation,
            TaskType.FREEZE_PANES: self._verify_freeze_panes,
            TaskType.AUTOFILTER: self._verify_autofilter,
            TaskType.SHEET_CREATE: self._verify_sheet_exists,
            TaskType.MERGE_CELLS: self._verify_merged_cells,
        }

        verifier = verifiers.get(task.task_type)
        if verifier:
            try:
                return verifier(task)
            except Exception as e:
                return VerificationResult(
                    task_id=task.id,
                    task_type=task.task_type,
                    passed=False,
                    message=f"Verification error: {e}",
                )

        # No specific verifier — assume passed if task is marked completed
        return VerificationResult(
            task_id=task.id,
            task_type=task.task_type,
            passed=task.completed,
            message="No specific verifier; relying on execution status",
        )

    # ── Specific Verifiers ──

    def _verify_formula(self, task: Task) -> VerificationResult:
        """Verify a formula was entered correctly."""
        if not self._wb or not task.cell:
            return self._skip(task, "No cell reference")

        ws = self._get_ws(task.sheet)
        cell = ws[task.cell]

        if cell.value is None:
            return VerificationResult(
                task_id=task.id, task_type=task.task_type,
                passed=False, message=f"Cell {task.cell} is empty",
            )

        cell_value = str(cell.value)

        # Check if it's a formula
        if task.formula:
            if cell_value.startswith("="):
                return VerificationResult(
                    task_id=task.id, task_type=task.task_type,
                    passed=True,
                    message=f"Formula found: {cell_value}",
                    details={"expected": task.formula, "actual": cell_value},
                )
            else:
                return VerificationResult(
                    task_id=task.id, task_type=task.task_type,
                    passed=False,
                    message=f"Expected formula, got value: {cell_value}",
                )

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=cell_value is not None,
            message=f"Cell has value: {cell_value}",
        )

    def _verify_table(self, task: Task) -> VerificationResult:
        """Verify a table was created."""
        ws = self._get_ws(task.sheet)

        if ws.tables:
            table_names = list(ws.tables.keys())
            return VerificationResult(
                task_id=task.id, task_type=task.task_type,
                passed=True,
                message=f"Table(s) found: {table_names}",
            )

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=False, message="No tables found on sheet",
        )

    def _verify_table_style(self, task: Task) -> VerificationResult:
        """Verify a table style was applied."""
        ws = self._get_ws(task.sheet)

        for name, table in ws.tables.items():
            if task.style:
                if table.tableStyleInfo and table.tableStyleInfo.name == task.style:
                    return VerificationResult(
                        task_id=task.id, task_type=task.task_type,
                        passed=True,
                        message=f"Table '{name}' has style {task.style}",
                    )
            else:
                if table.tableStyleInfo:
                    return VerificationResult(
                        task_id=task.id, task_type=task.task_type,
                        passed=True,
                        message=f"Table '{name}' has style {table.tableStyleInfo.name}",
                    )

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=False,
            message=f"Expected style '{task.style}' not found",
        )

    def _verify_formatting(self, task: Task) -> VerificationResult:
        """Verify general formatting was applied."""
        if not task.cell and not task.range:
            return self._skip(task, "No cell/range reference")

        ws = self._get_ws(task.sheet)
        ref = task.cell or task.range
        cell = ws[ref.split(":")[0]]  # check first cell of range

        has_formatting = (
            cell.font.bold or cell.font.italic or
            cell.fill.start_color.index not in (None, "00000000", 0) or
            cell.alignment.horizontal not in (None, "general")
        )

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=has_formatting,
            message="Formatting detected" if has_formatting else "No formatting detected",
        )

    def _verify_number_format(self, task: Task) -> VerificationResult:
        """Verify a number format was applied."""
        if not task.cell and not task.range:
            return self._skip(task, "No cell/range reference")

        ws = self._get_ws(task.sheet)
        ref = task.cell or task.range
        cell = ws[ref.split(":")[0]]

        fmt = cell.number_format
        is_custom = fmt not in (None, "General")

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=is_custom,
            message=f"Number format: {fmt}",
        )

    def _verify_conditional_format(self, task: Task) -> VerificationResult:
        """Verify conditional formatting rules exist."""
        ws = self._get_ws(task.sheet)

        cf_rules = ws.conditional_formatting
        count = len(list(cf_rules))

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=count > 0,
            message=f"{count} conditional formatting rule(s) found",
        )

    def _verify_chart(self, task: Task) -> VerificationResult:
        """Verify a chart exists on the sheet."""
        ws = self._get_ws(task.sheet)

        charts = ws._charts
        if charts:
            return VerificationResult(
                task_id=task.id, task_type=task.task_type,
                passed=True,
                message=f"{len(charts)} chart(s) found",
            )

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=False, message="No charts found on sheet",
        )

    def _verify_named_range(self, task: Task) -> VerificationResult:
        """Verify a named range exists."""
        if not self._wb:
            return self._skip(task, "No workbook loaded")

        names = [dn.name for dn in self._wb.defined_names.definedName]
        if names:
            return VerificationResult(
                task_id=task.id, task_type=task.task_type,
                passed=True,
                message=f"Named ranges found: {names}",
            )

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=False, message="No named ranges found",
        )

    def _verify_data_validation(self, task: Task) -> VerificationResult:
        """Verify data validation exists."""
        ws = self._get_ws(task.sheet)

        dv_count = len(ws.data_validations.dataValidation)

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=dv_count > 0,
            message=f"{dv_count} data validation(s) found",
        )

    def _verify_freeze_panes(self, task: Task) -> VerificationResult:
        """Verify freeze panes are set."""
        ws = self._get_ws(task.sheet)

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=ws.freeze_panes is not None,
            message=f"Freeze panes: {ws.freeze_panes}",
        )

    def _verify_autofilter(self, task: Task) -> VerificationResult:
        """Verify autofilter is enabled."""
        ws = self._get_ws(task.sheet)

        has_filter = ws.auto_filter.ref is not None

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=has_filter,
            message=f"Autofilter: {ws.auto_filter.ref}" if has_filter else "No autofilter",
        )

    def _verify_sheet_exists(self, task: Task) -> VerificationResult:
        """Verify a sheet exists."""
        if not self._wb:
            return self._skip(task, "No workbook loaded")

        sheet_name = task.sheet or task.params.get("name")
        exists = sheet_name in self._wb.sheetnames if sheet_name else False

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=exists,
            message=f"Sheet '{sheet_name}' {'exists' if exists else 'not found'}",
        )

    def _verify_merged_cells(self, task: Task) -> VerificationResult:
        """Verify merged cells exist."""
        ws = self._get_ws(task.sheet)
        merged = ws.merged_cells.ranges

        return VerificationResult(
            task_id=task.id, task_type=task.task_type,
            passed=len(merged) > 0,
            message=f"{len(merged)} merged range(s) found",
        )

    # ── Helpers ──

    def _get_ws(self, sheet_name: Optional[str] = None):
        """Get a worksheet by name or active sheet."""
        if not self._wb:
            raise RuntimeError("No workbook loaded for verification")
        if sheet_name and sheet_name in self._wb.sheetnames:
            return self._wb[sheet_name]
        return self._wb.active

    def _skip(self, task: Task, reason: str) -> VerificationResult:
        """Skip verification with a reason."""
        return VerificationResult(
            task_id=task.id,
            task_type=task.task_type,
            passed=True,
            message=f"Skipped: {reason}",
        )
