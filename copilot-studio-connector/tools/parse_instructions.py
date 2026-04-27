"""parse_instructions — parse an instruction file into structured tasks."""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

_ENGINE_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ENGINE_ROOT) not in sys.path:
    sys.path.insert(0, str(_ENGINE_ROOT))

DEFINITION: dict = {
    "name": "parse_instructions",
    "description": (
        "Parse an Excel assignment instruction file into structured task objects. "
        "Reads .docx, .rtfd, .pdf, or .txt files and extracts individual Excel operations "
        "such as formulas, formatting, charts, PivotTables, and named ranges. "
        "Returns a JSON object with a raw text preview and a list of typed task objects."
    ),
    "inputSchema": {
        "type": "object",
        "required": ["instruction_path"],
        "properties": {
            "instruction_path": {
                "type": "string",
                "description": (
                    "Absolute path to the instruction file. "
                    "Supported formats: .docx, .rtfd, .pdf, .txt"
                ),
            },
        },
    },
}


def _resolve(p: str) -> Path:
    resolved = Path(p).expanduser().resolve()
    if not resolved.is_relative_to(Path.home()):
        raise ValueError(f"Path must be under home directory: {resolved}")
    return resolved


def _run_sync(instruction_path: str) -> str:
    from excel_engine.parsers.instruction_parser import InstructionParser
    from excel_engine.parsers.task_extractor import TaskExtractor

    inst = _resolve(instruction_path)
    if not inst.exists() and not inst.is_dir():
        return json.dumps({"error": f"File not found: {inst}"})

    parser = InstructionParser()
    text = parser.parse(inst)

    extractor = TaskExtractor()
    tasks = extractor.extract(text)

    return json.dumps(
        {
            "instruction_file": str(inst),
            "raw_text_preview": text[:500] + ("..." if len(text) > 500 else ""),
            "raw_text_length": len(text),
            "task_count": len(tasks),
            "tasks": [
                {
                    "id": t.id,
                    "task_type": t.task_type.value,
                    "description": t.description,
                    "sheet": t.sheet,
                    "cell": t.cell,
                    "range": t.range,
                    "value": t.value,
                    "formula": t.formula,
                    "style": t.style,
                    "params": t.params,
                    "depends_on": t.depends_on,
                }
                for t in tasks
            ],
        },
        indent=2,
    )


async def run(arguments: dict) -> str:
    instruction_path = arguments.get("instruction_path", "").strip()
    if not instruction_path:
        raise ValueError("instruction_path is required")

    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _run_sync, instruction_path)
