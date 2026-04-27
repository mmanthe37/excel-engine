"""MCP JSON-RPC 2.0 dispatcher for Excel Engine Copilot Studio connector."""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Any

# ── Ensure tools/ and resources/ are importable ──────────────────────
_CONNECTOR_ROOT = Path(__file__).resolve().parent.parent
if str(_CONNECTOR_ROOT) not in sys.path:
    sys.path.insert(0, str(_CONNECTOR_ROOT))

from tools import TOOL_DEFINITIONS, TOOL_REGISTRY  # noqa: E402
from resources import RESOURCE_DEFINITIONS, RESOURCE_HANDLERS  # noqa: E402

logger = logging.getLogger(__name__)

_SERVER_INFO = {"name": "excel-engine-mcp", "version": "1.0.0"}
_PROTOCOL_VERSION = "2024-11-05"
_CAPABILITIES = {
    "tools": {"listChanged": False},
    "resources": {"subscribe": False, "listChanged": False},
}


class MCPHandler:
    """Handles MCP JSON-RPC 2.0 request dispatch."""

    async def handle(self, request: dict, session_id: str) -> dict:
        if not isinstance(request, dict):
            return self._error(None, -32600, "Invalid Request: expected a JSON object")

        req_id = request.get("id")
        method = request.get("method", "")
        params = request.get("params") or {}

        if request.get("jsonrpc") != "2.0":
            return self._error(req_id, -32600, "Invalid Request: jsonrpc must be '2.0'")

        try:
            return await self._dispatch(req_id, method, params, session_id)
        except Exception as exc:
            logger.exception("Unhandled error for method %r", method)
            return self._error(req_id, -32603, f"Internal error: {type(exc).__name__}")

    # ── Dispatcher ────────────────────────────────────────────────────

    async def _dispatch(
        self,
        req_id: Any,
        method: str,
        params: dict,
        session_id: str,
    ) -> dict:
        if method == "initialize":
            return self._ok(req_id, {
                "protocolVersion": _PROTOCOL_VERSION,
                "serverInfo": _SERVER_INFO,
                "capabilities": _CAPABILITIES,
            })

        if method in {"notifications/initialized", "notifications/cancelled"}:
            # Fire-and-forget notifications — respond with empty result
            return self._ok(req_id, {})

        if method == "ping":
            return self._ok(req_id, {})

        if method == "tools/list":
            return self._ok(req_id, {"tools": TOOL_DEFINITIONS})

        if method == "tools/call":
            return await self._tools_call(req_id, params)

        if method == "resources/list":
            return self._ok(req_id, {"resources": RESOURCE_DEFINITIONS})

        if method == "resources/read":
            return await self._resources_read(req_id, params)

        return self._error(req_id, -32601, f"Method not found: {method}")

    # ── tools/call ────────────────────────────────────────────────────

    async def _tools_call(self, req_id: Any, params: dict) -> dict:
        tool_name = params.get("name", "")
        arguments = params.get("arguments") or {}

        if not isinstance(arguments, dict):
            return self._error(req_id, -32602, "Invalid params: arguments must be an object")

        handler = TOOL_REGISTRY.get(tool_name)
        if handler is None:
            return self._error(req_id, -32601, f"Tool not found: {tool_name!r}")

        try:
            text = await handler(arguments)
            return self._ok(req_id, {
                "content": [{"type": "text", "text": text}],
                "isError": False,
            })
        except ValueError as exc:
            return self._ok(req_id, {
                "content": [{"type": "text", "text": str(exc)}],
                "isError": True,
            })
        except Exception as exc:
            logger.exception("Tool %r raised an unexpected error", tool_name)
            return self._ok(req_id, {
                "content": [{"type": "text", "text": f"Tool error ({type(exc).__name__}): {exc}"}],
                "isError": True,
            })

    # ── resources/read ────────────────────────────────────────────────

    async def _resources_read(self, req_id: Any, params: dict) -> dict:
        uri = params.get("uri", "")
        handler = RESOURCE_HANDLERS.get(uri)
        if handler is None:
            return self._error(req_id, -32601, f"Resource not found: {uri!r}")

        try:
            contents = await handler(params)
            return self._ok(req_id, {"contents": [contents]})
        except Exception as exc:
            logger.exception("Resource %r raised an unexpected error", uri)
            return self._error(req_id, -32603, f"Resource error: {exc}")

    # ── Helpers ───────────────────────────────────────────────────────

    @staticmethod
    def _ok(req_id: Any, result: dict) -> dict:
        return {"jsonrpc": "2.0", "id": req_id, "result": result}

    @staticmethod
    def _error(req_id: Any, code: int, message: str) -> dict:
        return {"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": message}}
