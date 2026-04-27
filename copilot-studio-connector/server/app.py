"""Excel Engine MCP — Copilot Studio Streamable HTTP Server

POST /mcp   — MCP JSON-RPC 2.0 endpoint (supports SSE streaming)
GET  /health — Liveness probe (no auth)

Environment variables:
    EXCEL_ENGINE_API_KEY  — shared secret validated in X-API-Key header.
                            Leave unset to disable auth (development only).
    PORT                  — TCP port to listen on (default: 8080).
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import sys
import uuid
from pathlib import Path
from typing import Any, AsyncGenerator

from fastapi import Depends, FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.security import APIKeyHeader

# ── Bootstrap package roots ──────────────────────────────────────────
# server/ sits inside copilot-studio-connector/; add both so that
# `mcp_handler` (sibling) and `tools`/`resources` (parent) are importable.
_SERVER_DIR = Path(__file__).resolve().parent
_CONNECTOR_ROOT = _SERVER_DIR.parent
for _p in (_SERVER_DIR, _CONNECTOR_ROOT):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from mcp_handler import MCPHandler  # noqa: E402

# ── Logging ──────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("excel-engine-mcp-http")

# ── FastAPI app ───────────────────────────────────────────────────────
app = FastAPI(
    title="Excel Engine MCP",
    description="MCP streamable HTTP server for Copilot Studio integration.",
    version="1.0.0",
    docs_url="/docs",
    redoc_url=None,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
    expose_headers=["Mcp-Session-Id"],
)

# ── Authentication ────────────────────────────────────────────────────
_api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)


def _require_api_key(api_key: str | None = Depends(_api_key_header)) -> str:
    expected = os.getenv("EXCEL_ENGINE_API_KEY", "")
    if expected and api_key != expected:
        raise HTTPException(status_code=401, detail="Invalid or missing X-API-Key header")
    return api_key or ""


# ── MCP handler ───────────────────────────────────────────────────────
_handler = MCPHandler()


# ── SSE helpers ───────────────────────────────────────────────────────
async def _sse_events(payloads: list[dict]) -> AsyncGenerator[str, None]:
    for payload in payloads:
        yield f"data: {json.dumps(payload)}\n\n"
        await asyncio.sleep(0)
    yield "data: [DONE]\n\n"


def _sse_response(payloads: list[dict], session_id: str) -> StreamingResponse:
    return StreamingResponse(
        _sse_events(payloads),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Mcp-Session-Id": session_id,
        },
    )


# ── Routes ────────────────────────────────────────────────────────────

@app.get("/health")
async def health_check() -> dict[str, str]:
    """Liveness probe — no authentication required."""
    return {"status": "ok", "server": "excel-engine-mcp", "version": "1.0.0"}


@app.post("/mcp")
async def mcp_endpoint(
    request: Request,
    _: str = Depends(_require_api_key),
) -> Any:
    """MCP JSON-RPC 2.0 endpoint with streamable HTTP support.

    Accepts ``application/json``.  When the client sends
    ``Accept: text/event-stream`` the response is delivered as SSE.
    """
    # ── Parse body ───────────────────────────────────────────────────
    try:
        body = await request.json()
    except Exception:
        error_resp = {
            "jsonrpc": "2.0",
            "id": None,
            "error": {"code": -32700, "message": "Parse error: request body is not valid JSON"},
        }
        return JSONResponse(error_resp, status_code=400)

    accepts_sse = "text/event-stream" in request.headers.get("Accept", "")
    session_id = request.headers.get("Mcp-Session-Id") or str(uuid.uuid4())

    # ── Batch request ─────────────────────────────────────────────────
    if isinstance(body, list):
        results = [await _handler.handle(req, session_id) for req in body]
        if accepts_sse:
            return _sse_response(results, session_id)
        return JSONResponse(results, headers={"Mcp-Session-Id": session_id})

    # ── Single request ────────────────────────────────────────────────
    result = await _handler.handle(body, session_id)

    if accepts_sse:
        return _sse_response([result], session_id)

    status = 400 if _is_client_error(result) else 200
    return JSONResponse(result, status_code=status, headers={"Mcp-Session-Id": session_id})


def _is_client_error(result: dict) -> bool:
    """Return True for JSON-RPC errors that map to HTTP 400."""
    error = result.get("error")
    if not error:
        return False
    return error.get("code", 0) in {-32600, -32601, -32602, -32700}


# ── Entry point ───────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn

    port = int(os.getenv("PORT", "8080"))
    uvicorn.run("app:app", host="0.0.0.0", port=port, reload=False)
