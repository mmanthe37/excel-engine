"""
Error Recovery — Classification, retry strategy, and backoff for task execution.

Provides:
  - TaskError: structured error record
  - ErrorClassifier: categorise exceptions as transient / permanent / layer_incompatible
  - RecoveryStrategy: decide whether to retry (with exponential back-off) or escalate
  - CircuitBreaker: prevent repeated calls to a consistently-failing layer
"""

from __future__ import annotations

import random
import time
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


# ── Circuit Breaker ──

class CircuitBreaker:
    """Prevent repeated calls to a consistently-failing layer.

    States per layer:
      - CLOSED (normal): requests pass through
      - OPEN: layer has hit failure_threshold; requests are blocked
      - HALF-OPEN: reset_timeout has elapsed since opening; one probe allowed

    After a successful probe the breaker resets to CLOSED.
    After a failed probe the breaker re-opens.
    """

    def __init__(
        self,
        failure_threshold: int = 5,
        reset_timeout: float = 300,
    ) -> None:
        self.failure_threshold = failure_threshold
        self.reset_timeout = reset_timeout
        self._failures: dict[str, int] = {}       # layer_name -> consecutive failures
        self._open_since: dict[str, float] = {}    # layer_name -> timestamp when opened
        self._half_open: dict[str, bool] = {}      # layer_name -> currently probing

    def record_failure(self, layer_name: str) -> None:
        """Record a consecutive failure for *layer_name*."""
        self._failures[layer_name] = self._failures.get(layer_name, 0) + 1
        if self._failures[layer_name] >= self.failure_threshold:
            self._open_since[layer_name] = time.time()
            self._half_open[layer_name] = False

    def record_success(self, layer_name: str) -> None:
        """Reset failure count after a success."""
        self._failures.pop(layer_name, None)
        self._open_since.pop(layer_name, None)
        self._half_open.pop(layer_name, None)

    def is_open(self, layer_name: str) -> bool:
        """Return True if the breaker is open (skip this layer).

        If reset_timeout has elapsed, transition to half-open and allow
        a single probe request (returns False once).
        """
        if layer_name not in self._open_since:
            return False

        elapsed = time.time() - self._open_since[layer_name]
        if elapsed >= self.reset_timeout:
            if not self._half_open.get(layer_name, False):
                # Transition to half-open: allow one probe
                self._half_open[layer_name] = True
                return False
        return True

    def reset(self, layer_name: str) -> None:
        """Manually reset a layer's circuit breaker."""
        self._failures.pop(layer_name, None)
        self._open_since.pop(layer_name, None)
        self._half_open.pop(layer_name, None)
