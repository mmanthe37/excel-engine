"""get_engine_status — return engine version, layers, and current configuration."""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

_ENGINE_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ENGINE_ROOT) not in sys.path:
    sys.path.insert(0, str(_ENGINE_ROOT))

DEFINITION: dict = {
    "name": "get_engine_status",
    "description": (
        "Return the current Excel Engine version, available automation layers, "
        "supported task types, layer-to-task-type mapping, and active configuration. "
        "Use this to confirm the engine is reachable and to discover supported operations "
        "before calling complete_assignment or execute_openpyxl."
    ),
    "inputSchema": {
        "type": "object",
        "properties": {},
    },
}


def _run_sync() -> str:
    from excel_engine import __version__ as engine_version
    from excel_engine.config import EngineConfig, Layer, TaskType, TASK_LAYER_MAP

    config = EngineConfig()
    return json.dumps(
        {
            "engine_version": engine_version,
            "mcp_server_version": "1.0.0",
            "copilot_studio_connector_version": "1.0.0",
            "available_layers": [
                {"number": layer.value, "name": layer.name}
                for layer in Layer
            ],
            "supported_task_types": [tt.value for tt in TaskType],
            "task_layer_mapping": {
                tt.value: [la.name for la in layers]
                for tt, layers in TASK_LAYER_MAP.items()
            },
            "config": {
                "scan_timeout": config.scan_timeout,
                "section_timeout": config.section_timeout,
                "applescript_timeout": config.applescript_timeout,
                "max_retries": config.max_retries,
                "retry_delay": config.retry_delay,
                "layer_order": [la.name for la in config.layer_order],
                "verify_after_each_section": config.verify_after_each_section,
                "parallel_execution": getattr(config, "parallel_execution", False),
                "max_workers": getattr(config, "max_workers", 4),
                "circuit_breaker_threshold": getattr(config, "circuit_breaker_threshold", 5),
            },
        },
        indent=2,
    )


async def run(arguments: dict) -> str:  # noqa: ARG001
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _run_sync)
