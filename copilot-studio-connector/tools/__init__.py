"""Tool registry — imports every tool module and exposes TOOL_REGISTRY / TOOL_DEFINITIONS."""

from __future__ import annotations

from .complete_assignment import DEFINITION as _D1, run as _run1
from .execute_live import DEFINITION as _D4, run as _run4
from .execute_openpyxl import DEFINITION as _D3, run as _run3
from .get_engine_status import DEFINITION as _D6, run as _run6
from .parse_instructions import DEFINITION as _D2, run as _run2
from .verify_workbook import DEFINITION as _D5, run as _run5

TOOL_REGISTRY: dict = {
    "complete_assignment": _run1,
    "parse_instructions": _run2,
    "execute_openpyxl": _run3,
    "execute_live": _run4,
    "verify_workbook": _run5,
    "get_engine_status": _run6,
}

TOOL_DEFINITIONS: list = [_D1, _D2, _D3, _D4, _D5, _D6]
