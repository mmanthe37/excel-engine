"""complete_assignment — run the full 6-layer Excel automation pipeline."""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path
from typing import Any

# ── Resolve excel_engine package ────────────────────────────────────
_ENGINE_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ENGINE_ROOT) not in sys.path:
    sys.path.insert(0, str(_ENGINE_ROOT))

# ── Copilot Studio-compatible tool definition ────────────────────────
# Rules: no $ref types, single types only, no enum fields.
DEFINITION: dict = {
    "name": "complete_assignment",
    "description": (
        "Complete an Excel assignment end-to-end. "
        "Parses the instruction file, builds a task plan, and executes all tasks "
        "using the 6-layer automation architecture (openpyxl → xlwings → AppleScript "
        "→ System Events → VBA → PyAutoGUI). Returns a JSON report with success status, "
        "section/task counts, verification results, timing, and any errors."
    ),
    "inputSchema": {
        "type": "object",
        "required": ["workbook_path", "instruction_path"],
        "properties": {
            "workbook_path": {
                "type": "string",
                "description": "Absolute path to the .xlsx workbook file (e.g. /Users/me/hw1.xlsx)",
            },
            "instruction_path": {
                "type": "string",
                "description": (
                    "Absolute path to the instruction file. "
                    "Supported formats: .docx, .rtfd, .pdf, .txt"
                ),
            },
            "max_retries": {
                "type": "integer",
                "description": "Maximum retry attempts per task before escalating layers (default: 3)",
            },
            "verify_after_each_section": {
                "type": "boolean",
                "description": "Run the verifier after completing each instruction section (default: true)",
            },
            "recalculate_formulas": {
                "type": "boolean",
                "description": (
                    "Use LibreOffice to recalculate all formulas and scan for errors "
                    "after execution (default: false). Requires LibreOffice installed."
                ),
            },
        },
    },
}


# ── Path guard ───────────────────────────────────────────────────────

def _resolve(p: str) -> Path:
    resolved = Path(p).expanduser().resolve()
    if not resolved.is_relative_to(Path.home()):
        raise ValueError(f"Path must be under home directory: {resolved}")
    return resolved


# ── Synchronous implementation (runs in thread pool) ─────────────────

def _run_sync(workbook_path: str, instruction_path: str, options: dict[str, Any]) -> str:
    from excel_engine import ExcelEngine, EngineConfig

    wb = _resolve(workbook_path)
    inst = _resolve(instruction_path)

    if not wb.exists():
        return json.dumps({"error": f"Workbook not found: {wb}"})
    if not inst.exists() and not inst.is_dir():
        return json.dumps({"error": f"Instruction file not found: {inst}"})

    config = EngineConfig()
    _SAFE_KEYS = {"max_retries", "verify_after_each_section", "retina_display",
                  "scan_timeout", "recalculate_formulas"}
    for key, val in options.items():
        if key in _SAFE_KEYS and hasattr(config, key):
            setattr(config, key, val)

    engine = ExcelEngine(config=config)
    result = engine.run(workbook=wb, instructions=inst)

    verifications = []
    for v in result.verifications:
        verifications.append({
            "section_id": v.section_id,
            "all_passed": v.all_passed,
            "pass_count": v.pass_count,
            "fail_count": v.fail_count,
            "results": [
                {
                    "task_id": r.task_id,
                    "task_type": r.task_type.value,
                    "passed": r.passed,
                    "message": r.message,
                }
                for r in v.results
            ],
        })

    return json.dumps(
        {
            "success": result.success,
            "workbook": str(result.workbook_path),
            "sections_completed": result.sections_completed,
            "sections_total": result.sections_total,
            "tasks_completed": result.tasks_completed,
            "tasks_total": result.tasks_total,
            "elapsed_seconds": round(result.elapsed_seconds, 2),
            "errors": result.errors,
            "verifications": verifications,
            "summary": result.summary(),
        },
        indent=2,
    )


# ── Async entry point ─────────────────────────────────────────────────

async def run(arguments: dict) -> str:
    workbook_path = arguments.get("workbook_path", "").strip()
    instruction_path = arguments.get("instruction_path", "").strip()

    if not workbook_path:
        raise ValueError("workbook_path is required")
    if not instruction_path:
        raise ValueError("instruction_path is required")

    options = {
        k: v
        for k, v in arguments.items()
        if k not in ("workbook_path", "instruction_path")
    }

    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _run_sync, workbook_path, instruction_path, options)
