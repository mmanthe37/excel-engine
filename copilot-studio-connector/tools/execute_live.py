"""execute_live — run Layers 2-6 live Excel operations."""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

_ENGINE_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ENGINE_ROOT) not in sys.path:
    sys.path.insert(0, str(_ENGINE_ROOT))

DEFINITION: dict = {
    "name": "execute_live",
    "description": (
        "Execute Excel tasks that require the live Microsoft Excel application (macOS only). "
        "Uses Layers 2-6: xlwings, AppleScript, System Events, VBA, and PyAutoGUI. "
        "Required for sorting, slicers, PivotTables, ribbon interactions, and any operation "
        "that needs a running Excel instance. "
        "Returns a JSON report of completed tasks and any errors."
    ),
    "inputSchema": {
        "type": "object",
        "required": ["workbook_path", "tasks"],
        "properties": {
            "workbook_path": {
                "type": "string",
                "description": "Absolute path to the .xlsx workbook file",
            },
            "tasks": {
                "type": "array",
                "description": (
                    "List of task descriptions for live Excel operations "
                    "(e.g. 'Create PivotTable on Sheet2', 'Sort column B ascending'). "
                    "Maximum 500 tasks per call."
                ),
                "items": {"type": "string"},
            },
        },
    },
}

_MAX_TASKS = 500
_MAX_TASK_LEN = 10_000


def _resolve(p: str) -> Path:
    resolved = Path(p).expanduser().resolve()
    if not resolved.is_relative_to(Path.home()):
        raise ValueError(f"Path must be under home directory: {resolved}")
    return resolved


def _run_sync(workbook_path: str, tasks: list[str]) -> str:
    from excel_engine import ExcelEngine, EngineConfig
    from excel_engine.config import Layer
    from excel_engine.parsers.task_extractor import TaskExtractor

    wb = _resolve(workbook_path)
    if not wb.exists():
        return json.dumps({"error": f"Workbook not found: {wb}"})

    extractor = TaskExtractor()
    task_text = "\n".join(f"- {t}" for t in tasks)
    task_objs = extractor.extract(task_text)

    config = EngineConfig()
    config.layer_order = [
        Layer.XLWINGS,
        Layer.APPLESCRIPT,
        Layer.SYSTEM_EVENTS,
        Layer.VBA,
        Layer.PYAUTOGUI,
    ]
    engine = ExcelEngine(config=config)
    result = engine.run(workbook=wb, tasks=task_objs)

    return json.dumps(
        {
            "success": result.success,
            "workbook": str(result.workbook_path),
            "tasks_completed": result.tasks_completed,
            "tasks_total": result.tasks_total,
            "elapsed_seconds": round(result.elapsed_seconds, 2),
            "errors": result.errors,
            "summary": result.summary(),
        },
        indent=2,
    )


async def run(arguments: dict) -> str:
    workbook_path = arguments.get("workbook_path", "").strip()
    tasks = arguments.get("tasks", [])

    if not workbook_path:
        raise ValueError("workbook_path is required")
    if not tasks:
        raise ValueError("tasks list must contain at least one item")
    if len(tasks) > _MAX_TASKS:
        raise ValueError(f"Too many tasks — maximum is {_MAX_TASKS}")
    if any(len(str(t)) > _MAX_TASK_LEN for t in tasks):
        raise ValueError(f"A task description exceeds the maximum length of {_MAX_TASK_LEN} characters")

    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _run_sync, workbook_path, list(tasks))
