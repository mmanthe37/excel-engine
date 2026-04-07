"""
Tests for the task planner and dependency graph.
"""

import pytest

from excel_engine.planner.dependency_graph import DependencyGraph
from excel_engine.planner.task_planner import TaskPlanner, Section
from excel_engine.parsers.task_extractor import Task
from excel_engine.config import TaskType


class TestDependencyGraph:
    def test_empty_graph(self):
        graph = DependencyGraph()
        assert graph.node_count == 0
        assert graph.edge_count == 0

    def test_add_nodes(self):
        graph = DependencyGraph()
        graph.add_node("a", "data_a")
        graph.add_node("b", "data_b")
        assert graph.node_count == 2

    def test_topological_sort_linear(self):
        graph = DependencyGraph()
        graph.add_node("a", "A")
        graph.add_node("b", "B")
        graph.add_node("c", "C")
        graph.add_edge("a", "b")  # a before b
        graph.add_edge("b", "c")  # b before c
        order = graph.topological_sort()
        assert order.index("a") < order.index("b")
        assert order.index("b") < order.index("c")

    def test_topological_sort_diamond(self):
        graph = DependencyGraph()
        graph.add_node("a", "A")
        graph.add_node("b", "B")
        graph.add_node("c", "C")
        graph.add_node("d", "D")
        graph.add_edge("a", "b")
        graph.add_edge("a", "c")
        graph.add_edge("b", "d")
        graph.add_edge("c", "d")
        order = graph.topological_sort()
        assert order.index("a") < order.index("b")
        assert order.index("a") < order.index("c")
        assert order.index("b") < order.index("d")
        assert order.index("c") < order.index("d")

    def test_cycle_detection(self):
        graph = DependencyGraph()
        graph.add_node("a", "A")
        graph.add_node("b", "B")
        graph.add_edge("a", "b")
        graph.add_edge("b", "a")
        assert graph.has_cycle() is True
        with pytest.raises(ValueError, match="cycle"):
            graph.topological_sort()

    def test_no_cycle(self):
        graph = DependencyGraph()
        graph.add_node("a", "A")
        graph.add_node("b", "B")
        graph.add_edge("a", "b")
        assert graph.has_cycle() is False

    def test_get_ready_nodes(self):
        graph = DependencyGraph()
        graph.add_node("a", "A")
        graph.add_node("b", "B")
        graph.add_node("c", "C")
        graph.add_edge("a", "b")
        graph.add_edge("a", "c")

        ready = graph.get_ready_nodes(completed=set())
        assert "a" in ready
        assert "b" not in ready

        ready = graph.get_ready_nodes(completed={"a"})
        assert "b" in ready
        assert "c" in ready

    def test_get_node(self):
        graph = DependencyGraph()
        graph.add_node("x", {"key": "value"})
        assert graph.get_node("x") == {"key": "value"}


class TestTaskPlanner:
    def setup_method(self):
        self.planner = TaskPlanner()

    def test_plan_empty(self):
        plan = self.planner.plan([])
        assert plan.section_count == 0
        assert plan.total_tasks == 0

    def test_plan_single_task(self):
        tasks = [
            Task(id="t1", task_type=TaskType.FORMULA, description="formula",
                 sheet="Sheet1", cell="A1", formula="=SUM(B1:B10)"),
        ]
        plan = self.planner.plan(tasks)
        assert plan.total_tasks == 1
        assert plan.section_count == 1

    def test_plan_groups_by_sheet(self):
        tasks = [
            Task(id="t1", task_type=TaskType.FORMULA, description="f1",
                 sheet="Sheet1"),
            Task(id="t2", task_type=TaskType.FORMULA, description="f2",
                 sheet="Sheet2"),
            Task(id="t3", task_type=TaskType.CHART_BAR, description="chart",
                 sheet="Sheet1"),
        ]
        plan = self.planner.plan(tasks)
        assert plan.section_count == 2  # Sheet1 and Sheet2
        assert plan.total_tasks == 3

    def test_plan_task_ordering(self):
        tasks = [
            Task(id="t1", task_type=TaskType.TABLE_STYLE, description="style",
                 sheet="S1"),
            Task(id="t2", task_type=TaskType.CELL_VALUE, description="value",
                 sheet="S1"),
            Task(id="t3", task_type=TaskType.TABLE_CREATE, description="create",
                 sheet="S1"),
        ]
        plan = self.planner.plan(tasks)
        section = plan.sections[0]
        types = [t.task_type for t in section.tasks]
        # CELL_VALUE should come before TABLE_CREATE which should come before TABLE_STYLE
        assert types.index(TaskType.CELL_VALUE) < types.index(TaskType.TABLE_CREATE)
        assert types.index(TaskType.TABLE_CREATE) < types.index(TaskType.TABLE_STYLE)

    def test_plan_estimated_time(self):
        tasks = [
            Task(id="t1", task_type=TaskType.FORMULA, description="f1"),
            Task(id="t2", task_type=TaskType.CHART_BAR, description="chart"),
        ]
        plan = self.planner.plan(tasks)
        assert plan.estimated_time_seconds > 0

    def test_plan_summary(self):
        tasks = [
            Task(id="t1", task_type=TaskType.FORMULA, description="f1",
                 sheet="Sheet1"),
        ]
        plan = self.planner.plan(tasks)
        summary = plan.summary()
        assert "Execution Plan" in summary
        assert "formula" in summary
