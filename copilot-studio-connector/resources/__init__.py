"""Resource registry for Excel Engine MCP server."""

from __future__ import annotations

from .workbook_report import (
    DEFINITION as _WORKBOOK_REPORT_DEF,
    URI as _WORKBOOK_REPORT_URI,
    read as _workbook_report_read,
)

RESOURCE_DEFINITIONS: list = [_WORKBOOK_REPORT_DEF]

# Map URI → async handler(params) → resource content dict
RESOURCE_HANDLERS: dict = {
    _WORKBOOK_REPORT_URI: _workbook_report_read,
}
