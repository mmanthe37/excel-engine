"""
Task Planner — Section-based planning for the Engine Protocol v2.0.

Groups tasks into logical sections by sheet/dependency, orders them
for execution, and validates prerequisite chains.
"""

from __future__ import annotations

import logging
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

from excel_engine.config import TaskType, Layer, TASK_LAYER_MAP, EngineConfig
from excel_engine.parsers.task_extractor import Task
from excel_engine.planner.dependency_graph import DependencyGraph

logger = logging.getLogger(__name__)


@dataclass
class Section:
    """A group of related tasks executed together."""
    id: str
    name: str
    sheet: Optional[str]
    tasks: list[Task] = field(default_factory=list)
    depends_on_sections: list[str] = field(default_factory=list)
    completed: bool = False

    @property
    def task_count(self) -> int:
        return len(self.tasks)


@dataclass
class ExecutionPlan:
    """The complete execution plan with ordered sections."""
    sections: list[Section]
    total_tasks: int
    estimated_time_seconds: float

    @property
    def section_count(self) -> int:
        return len(self.sections)

    def summary(self) -> str:
        lines = [f"Execution Plan: {self.section_count} sections, {self.total_tasks} tasks"]
        for sec in self.sections:
            lines.append(f"  [{sec.id}] {sec.name} ({sec.task_count} tasks)")
            for t in sec.tasks:
                layers = TASK_LAYER_MAP.get(t.task_type, [])
                layer_str = layers[0].name if layers else "?"
                lines.append(f"    - {t.id}: {t.task_type.value} → Layer {layer_str}")
        return "\n".join(lines)


# Time estimates per task type (seconds)
_TIME_ESTIMATES: dict[TaskType, float] = {
    TaskType.FORMULA: 2.0,
    TaskType.TABLE_CREATE: 3.0,
    TaskType.TABLE_STYLE: 2.0,
    TaskType.TABLE_TOTAL_ROW: 2.0,
    TaskType.CALCULATED_COLUMN: 5.0,
    TaskType.FORMATTING: 2.0,
    TaskType.CONDITIONAL_FORMAT: 4.0,
    TaskType.NUMBER_FORMAT: 2.0,
    TaskType.ALIGNMENT: 1.5,
    TaskType.COLUMN_WIDTH: 1.0,
    TaskType.FREEZE_PANES: 2.0,
    TaskType.SPLIT_PANES: 3.0,
    TaskType.AUTOFILTER: 3.0,
    TaskType.ADVANCED_FILTER: 5.0,
    TaskType.SORT: 5.0,
    TaskType.SUBTOTAL: 8.0,
    TaskType.CHART_BAR: 8.0,
    TaskType.CHART_LINE: 8.0,
    TaskType.CHART_PIE: 8.0,
    TaskType.CHART_HISTOGRAM: 20.0,
    TaskType.NAMED_RANGE: 2.0,
    TaskType.DATA_VALIDATION: 3.0,
    TaskType.SLICER: 15.0,
    TaskType.PIVOT_TABLE: 25.0,
    TaskType.PIVOT_CHART: 15.0,
    TaskType.SHEET_CREATE: 2.0,
    TaskType.SHEET_RENAME: 1.0,
    TaskType.CELL_VALUE: 1.0,
    TaskType.SAVE: 3.0,
}

# Task ordering priority within a section (lower = first)
_TASK_ORDER: dict[TaskType, int] = {
    TaskType.SHEET_CREATE: 0,
    TaskType.SHEET_RENAME: 1,
    TaskType.CELL_VALUE: 5,
    TaskType.FORMULA: 10,
    TaskType.NUMBER_FORMAT: 15,
    TaskType.FONT: 18,
    TaskType.FILL: 18,
    TaskType.ALIGNMENT: 18,
    TaskType.BORDER: 18,
    TaskType.COLUMN_WIDTH: 19,
    TaskType.ROW_HEIGHT: 19,
    TaskType.MERGE_CELLS: 20,
    TaskType.CONDITIONAL_FORMAT: 22,
    TaskType.TABLE_CREATE: 25,
    TaskType.TABLE_STYLE: 26,
    TaskType.TABLE_TOTAL_ROW: 27,
    TaskType.CALCULATED_COLUMN: 28,
    TaskType.NAMED_RANGE: 30,
    TaskType.DATA_VALIDATION: 32,
    TaskType.AUTOFILTER: 35,
    TaskType.SORT: 36,
    TaskType.ADVANCED_FILTER: 37,
    TaskType.SUBTOTAL: 38,
    TaskType.FREEZE_PANES: 40,
    TaskType.SPLIT_PANES: 41,
    TaskType.CHART_BAR: 50,
    TaskType.CHART_LINE: 50,
    TaskType.CHART_PIE: 50,
    TaskType.CHART_HISTOGRAM: 55,
    TaskType.SLICER: 60,
    TaskType.PIVOT_TABLE: 70,
    TaskType.PIVOT_CHART: 75,
    TaskType.PRINT_SETTINGS: 80,
    TaskType.SAVE: 99,
    TaskType.SAVE_AS: 100,
}


class TaskPlanner:
    """
    Section-based task planner implementing Engine Protocol v2.0.

    Steps:
      1. Group tasks into sections (by sheet/dependency)
      2. Order tasks within each section
      3. Determine section execution order via dependency graph
      4. Validate plan soundness
    """

    def __init__(self, config: Optional[EngineConfig] = None) -> None:
        self.config = config or EngineConfig()

    def plan(self, tasks: list[Task]) -> ExecutionPlan:
        """
        Create an execution plan from a list of tasks.
        Groups by sheet, orders by dependency, validates.
        """
        if not tasks:
            return ExecutionPlan(sections=[], total_tasks=0, estimated_time_seconds=0)

        # Step 1: Group into sections by sheet
        sections = self._group_by_sheet(tasks)

        # Step 2: Order tasks within each section
        for section in sections:
            section.tasks = self._order_tasks(section.tasks)

        # Step 3: Determine section execution order
        sections = self._order_sections(sections)

        # Step 4: Validate
        issues = self._validate_plan(sections)
        if issues:
            for issue in issues:
                logger.warning("Plan issue: %s", issue)

        # Estimate time
        total_time = sum(
            _TIME_ESTIMATES.get(t.task_type, 5.0)
            for sec in sections
            for t in sec.tasks
        )

        plan = ExecutionPlan(
            sections=sections,
            total_tasks=sum(s.task_count for s in sections),
            estimated_time_seconds=total_time,
        )

        logger.info(
            "Plan created: %d sections, %d tasks, ~%.0fs estimated",
            plan.section_count, plan.total_tasks, plan.estimated_time_seconds,
        )
        return plan

    def _group_by_sheet(self, tasks: list[Task]) -> list[Section]:
        """Group tasks into sections by worksheet."""
        groups: dict[Optional[str], list[Task]] = defaultdict(list)
        for task in tasks:
            groups[task.sheet].append(task)

        sections = []
        for i, (sheet, sheet_tasks) in enumerate(groups.items()):
            section_name = f"Sheet: {sheet}" if sheet else "General"
            sections.append(Section(
                id=f"sec_{i + 1:02d}",
                name=section_name,
                sheet=sheet,
                tasks=sheet_tasks,
            ))

        return sections

    def _order_tasks(self, tasks: list[Task]) -> list[Task]:
        """Order tasks within a section by type priority and dependencies."""
        graph = DependencyGraph[Task]()

        for task in tasks:
            graph.add_node(task.id, task)

        for task in tasks:
            for dep_id in task.depends_on:
                # Only add edge if dependency is in this section
                if any(t.id == dep_id for t in tasks):
                    graph.add_edge(dep_id, task.id)

        try:
            order = graph.topological_sort()
            ordered = [graph.get_node(nid) for nid in order]
        except ValueError:
            logger.warning("Cycle in task dependencies, falling back to priority sort")
            ordered = tasks

        # Stable sort by task type priority
        ordered.sort(key=lambda t: _TASK_ORDER.get(t.task_type, 50))
        return ordered

    def _order_sections(self, sections: list[Section]) -> list[Section]:
        """Order sections by inter-section dependencies."""
        # Build a mapping of task_id → section_id
        task_to_section: dict[str, str] = {}
        for sec in sections:
            for task in sec.tasks:
                task_to_section[task.id] = sec.id

        # Build section dependency graph
        graph = DependencyGraph[Section]()
        for sec in sections:
            graph.add_node(sec.id, sec)

        for sec in sections:
            for task in sec.tasks:
                for dep_id in task.depends_on:
                    dep_sec = task_to_section.get(dep_id)
                    if dep_sec and dep_sec != sec.id:
                        graph.add_edge(dep_sec, sec.id)
                        if dep_sec not in sec.depends_on_sections:
                            sec.depends_on_sections.append(dep_sec)

        try:
            order = graph.topological_sort()
            return [graph.get_node(sid) for sid in order]
        except ValueError:
            logger.warning("Cycle in section dependencies, using original order")
            return sections

    def _validate_plan(self, sections: list[Section]) -> list[str]:
        """Validate the plan for common issues."""
        issues: list[str] = []

        for sec in sections:
            for task in sec.tasks:
                # Check: table operations need table creation first
                if task.task_type in (
                    TaskType.TABLE_STYLE, TaskType.TABLE_TOTAL_ROW,
                    TaskType.CALCULATED_COLUMN,
                ):
                    has_table_create = any(
                        t.task_type == TaskType.TABLE_CREATE
                        for t in sec.tasks
                        if _TASK_ORDER.get(t.task_type, 50) < _TASK_ORDER.get(task.task_type, 50)
                    )
                    if not has_table_create and not task.depends_on:
                        issues.append(
                            f"Task {task.id} ({task.task_type.value}) may need a "
                            f"TABLE_CREATE prerequisite in section {sec.id}"
                        )

                # Check: structural references should not go through openpyxl
                if task.task_type == TaskType.CALCULATED_COLUMN:
                    layers = TASK_LAYER_MAP.get(task.task_type, [])
                    if Layer.OPENPYXL in layers and layers[0] == Layer.OPENPYXL:
                        issues.append(
                            f"Task {task.id}: Calculated column should use LIVE layer "
                            f"(xlwings/AppleScript), not openpyxl (will produce #REF!)"
                        )

        return issues
