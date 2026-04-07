"""
Dependency Graph — DAG-based task dependency resolution.

Provides topological sorting and cycle detection for task execution ordering.
"""

from __future__ import annotations

import logging
from collections import defaultdict, deque
from typing import TypeVar, Generic

logger = logging.getLogger(__name__)

T = TypeVar("T")


class DependencyGraph(Generic[T]):
    """
    Directed Acyclic Graph for task dependency resolution.
    Supports topological sort with Kahn's algorithm.
    """

    def __init__(self) -> None:
        self._adj: dict[str, set[str]] = defaultdict(set)  # node → dependents
        self._in_degree: dict[str, int] = defaultdict(int)
        self._nodes: dict[str, T] = {}

    def add_node(self, node_id: str, data: T) -> None:
        """Add a node to the graph."""
        self._nodes[node_id] = data
        if node_id not in self._in_degree:
            self._in_degree[node_id] = 0

    def add_edge(self, from_id: str, to_id: str) -> None:
        """
        Add a dependency edge: to_id depends on from_id.
        (from_id must complete before to_id)
        """
        if to_id not in self._adj[from_id]:
            self._adj[from_id].add(to_id)
            self._in_degree[to_id] = self._in_degree.get(to_id, 0) + 1

    def topological_sort(self) -> list[str]:
        """
        Return nodes in topological order using Kahn's algorithm.
        Raises ValueError if a cycle is detected.
        """
        in_degree = dict(self._in_degree)
        queue: deque[str] = deque()

        for node_id in self._nodes:
            if in_degree.get(node_id, 0) == 0:
                queue.append(node_id)

        result: list[str] = []
        while queue:
            node_id = queue.popleft()
            result.append(node_id)
            for dependent in self._adj.get(node_id, set()):
                in_degree[dependent] -= 1
                if in_degree[dependent] == 0:
                    queue.append(dependent)

        if len(result) != len(self._nodes):
            missing = set(self._nodes.keys()) - set(result)
            raise ValueError(
                f"Dependency cycle detected involving: {missing}"
            )

        return result

    def get_node(self, node_id: str) -> T:
        """Get node data by ID."""
        return self._nodes[node_id]

    def get_ready_nodes(self, completed: set[str]) -> list[str]:
        """Get all nodes whose dependencies are fully completed."""
        ready = []
        for node_id in self._nodes:
            if node_id in completed:
                continue
            deps_met = True
            for dep_id in self._get_dependencies(node_id):
                if dep_id not in completed:
                    deps_met = False
                    break
            if deps_met:
                ready.append(node_id)
        return ready

    def _get_dependencies(self, node_id: str) -> list[str]:
        """Get all nodes that node_id depends on (reverse edges)."""
        deps = []
        for from_id, to_ids in self._adj.items():
            if node_id in to_ids:
                deps.append(from_id)
        return deps

    def has_cycle(self) -> bool:
        """Check if the graph has any cycles."""
        try:
            self.topological_sort()
            return False
        except ValueError:
            return True

    @property
    def node_count(self) -> int:
        return len(self._nodes)

    @property
    def edge_count(self) -> int:
        return sum(len(deps) for deps in self._adj.values())

    def __repr__(self) -> str:
        return f"DependencyGraph(nodes={self.node_count}, edges={self.edge_count})"
