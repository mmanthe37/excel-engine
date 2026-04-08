"""
Error Recovery — Classification, retry strategy, and backoff for task execution.

Provides:
  - TaskError: structured error record
  - ErrorClassifier: categorise exceptions as transient / permanent / layer_incompatible
  - RecoveryStrategy: decide whether to retry (with exponential back-off) or escalate
"""

from __future__ import annotations

import random
from dataclasses import dataclass, field

from excel_engine.config import Layer, TaskType


# ── Error record ──

@dataclass
class TaskError:
    """A single recorded error from task execution."""
    task_id: str
    task_type: TaskType
    layer: Layer
    error_type: str  # "transient", "permanent", "layer_incompatible"
    message: str
    timestamp: float


# ── Error classification ──

class ErrorClassifier:
    """Classify exceptions to decide retry vs. escalate."""

    TRANSIENT_PATTERNS: list[str] = [
        "timed out",
        "connection",
        "not responding",
        "busy",
        "AppleEvent timed out",
        "Application isn't running",
        "System Events got an error",
    ]

    PERMANENT_PATTERNS: list[str] = [
        "FileNotFoundError",
        "PermissionError",
        "No such sheet",
        "not supported",
        "invalid reference",
    ]

    @classmethod
    def classify(cls, error: Exception, layer: Layer) -> str:
        """Return ``"transient"``, ``"permanent"``, or ``"layer_incompatible"``."""
        msg = f"{type(error).__name__}: {error}"

        # Check transient patterns first
        for pattern in cls.TRANSIENT_PATTERNS:
            if pattern.lower() in msg.lower():
                return "transient"

        # Check permanent patterns
        for pattern in cls.PERMANENT_PATTERNS:
            if pattern.lower() in msg.lower():
                return "permanent"

        # NotImplementedError or ValueError("Unknown layer") ⇒ layer mismatch
        if isinstance(error, (NotImplementedError, AttributeError)):
            return "layer_incompatible"

        # Default: treat unknown errors as permanent so we escalate quickly
        return "permanent"


# ── Recovery / retry strategy ──

class RecoveryStrategy:
    """Decide whether to retry, escalate to the next layer, or abort."""

    def __init__(
        self,
        max_retries: int = 3,
        base_delay: float = 1.0,
        max_delay: float = 30.0,
    ) -> None:
        self.max_retries = max_retries
        self.base_delay = base_delay
        self.max_delay = max_delay

    def should_retry(self, error_type: str, attempt: int) -> bool:
        """Return *True* if the caller should retry the current layer.

        * transient → retry up to *max_retries*
        * permanent / layer_incompatible → never retry, escalate
        """
        if error_type != "transient":
            return False
        return attempt < self.max_retries

    def get_delay(self, attempt: int) -> float:
        """Exponential back-off with jitter, capped at *max_delay*.

        delay = min(base_delay * 2^attempt, max_delay) + jitter
        """
        delay = min(self.base_delay * (2 ** attempt), self.max_delay)
        jitter = random.uniform(0, delay * 0.1)
        return delay + jitter
