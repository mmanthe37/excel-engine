"""
Task Extractor — Extract structured tasks from parsed instruction text.

Identifies specific Excel operations (formulas, tables, formatting, charts, etc.)
from natural language instructions and maps them to TaskType enums.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Optional

from excel_engine.config import TaskType

logger = logging.getLogger(__name__)


@dataclass
class Task:
    """A single extracted task to be executed."""
    id: str
    task_type: TaskType
    description: str
    sheet: Optional[str] = None
    cell: Optional[str] = None
    range: Optional[str] = None
    value: Optional[str] = None
    formula: Optional[str] = None
    style: Optional[str] = None
    params: dict = field(default_factory=dict)
    depends_on: list[str] = field(default_factory=list)
    completed: bool = False

    @property
    def key(self) -> str:
        """Unique key for dependency resolution."""
        return f"{self.sheet or 'default'}:{self.id}"


# Regex patterns to identify task types from instruction text
_PATTERNS: dict[TaskType, list[re.Pattern]] = {
    TaskType.FORMULA: [
        re.compile(r"(?:enter|type|input|create|add)\s+(?:the\s+)?formula\b", re.I),
        re.compile(r"(?:in\s+cell\s+\w+),?\s*(?:enter|type)\s+=", re.I),
        re.compile(r"=\w+\(", re.I),  # literal formula like =SUM(
    ],
    TaskType.TABLE_CREATE: [
        re.compile(r"(?:create|format\s+as|convert\s+to)\s+(?:a\s+)?(?:an?\s+)?(?:Excel\s+)?table\b", re.I),
    ],
    TaskType.TABLE_STYLE: [
        re.compile(r"(?:apply|use|change)\s+(?:the\s+)?(?:table\s+)?style\s+(\w+)", re.I),
        re.compile(r"TableStyle\w+", re.I),
    ],
    TaskType.TABLE_TOTAL_ROW: [
        re.compile(r"total\s*row", re.I),
        re.compile(r"(?:add|show|enable)\s+(?:the\s+)?totals?\s+row", re.I),
    ],
    TaskType.CALCULATED_COLUMN: [
        re.compile(r"calculated\s+column", re.I),
        re.compile(r"\[@\w+\]", re.I),  # structural reference
        re.compile(r"\[@\[.+?\]\]", re.I),  # structural reference with spaces
    ],
    TaskType.CONDITIONAL_FORMAT: [
        re.compile(r"conditional\s+format", re.I),
        re.compile(r"highlight\s+cells?\s+(?:rules?|that)", re.I),
        re.compile(r"color\s+scale", re.I),
        re.compile(r"data\s+bars?", re.I),
    ],
    TaskType.NUMBER_FORMAT: [
        re.compile(r"(?:format|change)\s+(?:the\s+)?(?:cells?\s+)?(?:as|to)\s+(?:currency|accounting|percentage|number|date|time|text)", re.I),
        re.compile(r"number\s+format", re.I),
        re.compile(r"(?:apply|use)\s+(?:the\s+)?(?:Accounting|Currency|Percentage|Number)\s+format", re.I),
    ],
    TaskType.ALIGNMENT: [
        re.compile(r"(?:center|left|right)\s+align", re.I),
        re.compile(r"(?:align|alignment)\s+(?:to\s+)?(?:center|left|right|top|bottom|middle)", re.I),
        re.compile(r"wrap\s+text", re.I),
        re.compile(r"merge\s+(?:and\s+)?center", re.I),
    ],
    TaskType.COLUMN_WIDTH: [
        re.compile(r"(?:change|set|adjust|resize)\s+(?:the\s+)?column\s+width", re.I),
        re.compile(r"autofit\s+(?:column|width)", re.I),
        re.compile(r"column\s+width\s+(?:to\s+)?(\d+)", re.I),
    ],
    TaskType.FREEZE_PANES: [
        re.compile(r"freeze\s+(?:the\s+)?(?:top\s+)?(?:row|panes?|column)", re.I),
    ],
    TaskType.SPLIT_PANES: [
        re.compile(r"split\s+(?:the\s+)?(?:window|panes?)", re.I),
    ],
    TaskType.AUTOFILTER: [
        re.compile(r"auto\s*filter", re.I),
        re.compile(r"(?:apply|add|enable)\s+(?:a\s+)?filter", re.I),
    ],
    TaskType.ADVANCED_FILTER: [
        re.compile(r"advanced\s+filter", re.I),
    ],
    TaskType.SORT: [
        re.compile(r"sort\s+(?:the\s+)?(?:data|range|table)", re.I),
        re.compile(r"(?:sort\s+)?(?:ascending|descending)\s+(?:by|order)", re.I),
    ],
    TaskType.SUBTOTAL: [
        re.compile(r"subtotal", re.I),
    ],
    TaskType.CHART_BAR: [
        re.compile(r"(?:create|insert|add)\s+(?:a\s+)?(?:bar|column)\s+chart", re.I),
        re.compile(r"(?:clustered|stacked)\s+(?:bar|column)", re.I),
    ],
    TaskType.CHART_LINE: [
        re.compile(r"(?:create|insert|add)\s+(?:a\s+)?line\s+chart", re.I),
    ],
    TaskType.CHART_PIE: [
        re.compile(r"(?:create|insert|add)\s+(?:a\s+)?pie\s+chart", re.I),
    ],
    TaskType.CHART_HISTOGRAM: [
        re.compile(r"histogram", re.I),
    ],
    TaskType.NAMED_RANGE: [
        re.compile(r"(?:create|define|name)\s+(?:a\s+)?(?:named\s+)?range", re.I),
        re.compile(r"name\s+manager", re.I),
    ],
    TaskType.DATA_VALIDATION: [
        re.compile(r"data\s+validation", re.I),
        re.compile(r"(?:drop\s*-?\s*down|dropdown)\s+list", re.I),
    ],
    TaskType.SLICER: [
        re.compile(r"(?:insert|create|add)\s+(?:a\s+)?slicer", re.I),
    ],
    TaskType.PIVOT_TABLE: [
        re.compile(r"pivot\s*table", re.I),
    ],
    TaskType.PIVOT_CHART: [
        re.compile(r"pivot\s*chart", re.I),
    ],
    TaskType.SHEET_CREATE: [
        re.compile(r"(?:insert|add|create)\s+(?:a\s+)?(?:new\s+)?(?:work)?sheet", re.I),
    ],
    TaskType.SHEET_RENAME: [
        re.compile(r"rename\s+(?:the\s+)?(?:work)?sheet", re.I),
    ],
    TaskType.FONT: [
        re.compile(r"(?:change|set|apply)\s+(?:the\s+)?font", re.I),
        re.compile(r"(?:bold|italic|underline)\b", re.I),
        re.compile(r"font\s+(?:size|color|name)", re.I),
    ],
    TaskType.FILL: [
        re.compile(r"(?:fill|background)\s+color", re.I),
        re.compile(r"(?:shade|highlight)\s+(?:the\s+)?cell", re.I),
    ],
    TaskType.BORDER: [
        re.compile(r"(?:add|apply|draw)\s+(?:a\s+)?border", re.I),
        re.compile(r"(?:outside|inside|bottom|top)\s+border", re.I),
    ],
    TaskType.MERGE_CELLS: [
        re.compile(r"merge\s+(?:and\s+center\s+)?cells?", re.I),
    ],
    TaskType.SAVE: [
        re.compile(r"save\s+(?:the\s+)?workbook\b", re.I),
    ],
    TaskType.PRINT_SETTINGS: [
        re.compile(r"(?:set|change)\s+(?:the\s+)?(?:print|page)\s+(?:area|orientation|margins?|header|footer)", re.I),
        re.compile(r"landscape\s+orientation", re.I),
    ],
}

# Pattern to extract cell references
_CELL_REF = re.compile(r"\b([A-Z]{1,3}\d{1,7})\b")
_RANGE_REF = re.compile(r"\b([A-Z]{1,3}\d{1,7}:[A-Z]{1,3}\d{1,7})\b")
_SHEET_REF = re.compile(
    r"(?:(?:on|in|of|to)\s+)?(?:the\s+)?(?:sheet|worksheet)\s+[\"']?(\w[\w\s]*?)[\"']?(?:\s|,|\.|$)",
    re.I,
)
_FORMULA_REF = re.compile(r"(=[A-Z]+\(.*?\))", re.I)


class TaskExtractor:
    """Extract structured Task objects from instruction text."""

    def __init__(self) -> None:
        self._counter = 0

    def extract(self, text: str) -> list[Task]:
        """
        Extract all tasks from instruction text.
        Splits text into instruction lines/paragraphs and identifies tasks.
        """
        self._counter = 0
        tasks: list[Task] = []

        # Split into numbered steps or paragraphs
        lines = self._split_instructions(text)

        for line in lines:
            line_tasks = self._extract_from_line(line)
            tasks.extend(line_tasks)

        # Auto-detect dependencies
        self._resolve_dependencies(tasks)

        logger.info("Extracted %d tasks from instructions", len(tasks))
        return tasks

    def _next_id(self) -> str:
        """Generate a sequential task ID."""
        self._counter += 1
        return f"task_{self._counter:03d}"

    def _split_instructions(self, text: str) -> list[str]:
        """Split instruction text into individual instruction lines."""
        # Try numbered steps first (1. or Step 1: or a.)
        numbered = re.split(r"\n\s*(?:\d+[.)]\s+|Step\s+\d+[:.]\s+|[a-z][.)]\s+)", text)
        if len(numbered) > 2:
            return [s.strip() for s in numbered if s.strip()]

        # Fall back to sentence/paragraph splitting
        paragraphs = text.split("\n")
        result = []
        for p in paragraphs:
            p = p.strip()
            if p and len(p) > 10:
                result.append(p)
        return result

    def _extract_from_line(self, line: str) -> list[Task]:
        """Extract tasks from a single instruction line."""
        tasks = []

        for task_type, patterns in _PATTERNS.items():
            for pattern in patterns:
                if pattern.search(line):
                    task = self._build_task(task_type, line)
                    tasks.append(task)
                    break  # one match per task type per line

        # If no pattern matched but line contains a formula, treat as FORMULA
        if not tasks:
            formula_match = _FORMULA_REF.search(line)
            if formula_match:
                task = self._build_task(TaskType.FORMULA, line)
                task.formula = formula_match.group(1)
                tasks.append(task)

        return tasks

    def _build_task(self, task_type: TaskType, line: str) -> Task:
        """Build a Task object from a matched line."""
        task = Task(
            id=self._next_id(),
            task_type=task_type,
            description=line[:200],
        )

        # Extract sheet reference
        sheet_match = _SHEET_REF.search(line)
        if sheet_match:
            task.sheet = sheet_match.group(1).strip()

        # Extract cell reference
        range_match = _RANGE_REF.search(line)
        if range_match:
            task.range = range_match.group(1)
        else:
            cell_match = _CELL_REF.search(line)
            if cell_match:
                task.cell = cell_match.group(1)

        # Extract formula
        formula_match = _FORMULA_REF.search(line)
        if formula_match:
            task.formula = formula_match.group(1)

        # Extract table style names
        style_match = re.search(r"(TableStyle\w+\d+)", line)
        if style_match:
            task.style = style_match.group(1)

        return task

    def _resolve_dependencies(self, tasks: list[Task]) -> None:
        """
        Auto-detect task dependencies based on ordering rules:
          - Table operations depend on table creation
          - Calculated columns depend on table creation
          - Charts depend on data being present
          - Slicers depend on tables/pivots
          - Save depends on everything
        """
        table_creates = [t for t in tasks if t.task_type == TaskType.TABLE_CREATE]
        pivot_creates = [t for t in tasks if t.task_type == TaskType.PIVOT_TABLE]

        for task in tasks:
            # Table-dependent operations
            if task.task_type in (
                TaskType.TABLE_STYLE, TaskType.TABLE_TOTAL_ROW,
                TaskType.CALCULATED_COLUMN,
            ):
                for tc in table_creates:
                    if tc.sheet == task.sheet or task.sheet is None:
                        task.depends_on.append(tc.id)

            # Slicer depends on table or pivot
            if task.task_type == TaskType.SLICER:
                for tc in table_creates + pivot_creates:
                    task.depends_on.append(tc.id)

            # PivotChart depends on PivotTable
            if task.task_type == TaskType.PIVOT_CHART:
                for pc in pivot_creates:
                    task.depends_on.append(pc.id)

            # Save depends on all other tasks
            if task.task_type == TaskType.SAVE:
                for other in tasks:
                    if other.id != task.id and other.task_type != TaskType.SAVE:
                        task.depends_on.append(other.id)
