"""
Interactive Mode — Guided step-by-step Excel Engine session.

Prompts the user for inputs, shows the plan, and executes with progress.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Optional

import excel_engine
from excel_engine.config import EngineConfig, TASK_LAYER_MAP
from excel_engine.parsers.instruction_parser import InstructionParser
from excel_engine.parsers.task_extractor import TaskExtractor
from excel_engine.planner.task_planner import TaskPlanner


def _prompt(message: str, default: str = "") -> str:
    """Prompt for input with an optional default."""
    suffix = f" [{default}]" if default else ""
    try:
        value = input(f"{message}{suffix}: ").strip()
    except EOFError:
        return default
    return value or default


def _confirm(message: str, default: bool = True) -> bool:
    """Ask a yes/no question."""
    hint = "Y/n" if default else "y/N"
    try:
        answer = input(f"{message} ({hint}): ").strip().lower()
    except EOFError:
        return default
    if not answer:
        return default
    return answer in ("y", "yes")


def _resolve_path(raw: str) -> Optional[Path]:
    """Resolve a user-provided path, expanding ~ and checking existence."""
    if not raw:
        return None
    p = Path(raw).expanduser().resolve()
    if p.exists():
        return p
    # Try as-is (maybe relative)
    p2 = Path.cwd() / raw
    if p2.exists():
        return p2.resolve()
    return None


def interactive_session() -> int:
    """Run an interactive guided session."""
    print()
    print("╔══════════════════════════════════════════════╗")
    print(f"║  Excel Engine v{excel_engine.__version__:<8} — Interactive Mode   ║")
    print("╚══════════════════════════════════════════════╝")
    print()

    # Step 1: Workbook path
    print("Step 1: Workbook")
    print("─" * 40)
    while True:
        raw = _prompt("  Path to Excel workbook (.xlsx)")
        if not raw:
            print("  Workbook path is required.")
            continue
        workbook = _resolve_path(raw)
        if workbook and workbook.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
            print(f"  ✓ Found: {workbook.name}")
            break
        if workbook:
            print(f"  ✗ Not a recognized Excel file: {workbook.suffix}")
        else:
            print(f"  ✗ File not found: {raw}")
    print()

    # Step 2: Instructions
    print("Step 2: Instructions")
    print("─" * 40)
    while True:
        raw = _prompt("  Path to instruction file (.docx, .rtfd, .pdf, .txt)")
        if not raw:
            print("  Instruction file is required.")
            continue
        instructions = _resolve_path(raw)
        if instructions:
            print(f"  ✓ Found: {instructions.name}")
            break
        print(f"  ✗ File not found: {raw}")
    print()

    # Step 3: Parse & Show Plan
    print("Step 3: Parsing Instructions")
    print("─" * 40)

    try:
        parser = InstructionParser()
        extractor = TaskExtractor()
        text = parser.parse(instructions)
        tasks = extractor.extract(text)
    except Exception as exc:
        print(f"  ✗ Failed to parse instructions: {exc}", file=sys.stderr)
        return 1

    print(f"  ✓ Extracted {len(tasks)} tasks")
    print()

    # Build the execution plan
    config = EngineConfig()
    planner = TaskPlanner(config=config)
    plan = planner.plan(tasks)

    print("Step 4: Execution Plan")
    print("─" * 40)
    print(f"  Sections:       {plan.section_count}")
    print(f"  Total tasks:    {plan.total_tasks}")
    print(f"  Estimated time: {plan.estimated_time_seconds:.0f}s")
    print()

    for section in plan.sections:
        print(f"  [{section.id}] {section.name} ({section.task_count} tasks)")
        for t in section.tasks:
            layers = TASK_LAYER_MAP.get(t.task_type, [])
            layer_name = layers[0].name if layers else "?"
            print(f"      • {t.task_type.value}: {t.description[:60]}")
            print(f"        → Layer: {layer_name}")
    print()

    # Step 5: Confirm
    if not _confirm("  Proceed with execution?"):
        print("  Aborted by user.")
        return 0
    print()

    # Step 6: Execute
    print("Step 5: Executing")
    print("─" * 40)

    from excel_engine.engine import ExcelEngine

    engine = ExcelEngine(config=config)
    result = engine.run(workbook=workbook, tasks=tasks)

    print()

    # Step 7: Results
    print("Step 6: Results")
    print("─" * 40)
    print(result.summary())
    print()

    # Verification details
    if result.verifications:
        print("Verification Details:")
        for sv in result.verifications:
            icon = "✓" if sv.all_passed else "✗"
            print(f"  {icon} {sv.summary()}")
    print()

    # Offer to save results
    if _confirm("  Save results to JSON?", default=False):
        default_out = workbook.stem + "_results.json"
        out_path = _prompt("  Output file", default=default_out)
        data = {
            "success": result.success,
            "workbook": str(result.workbook_path),
            "sections_completed": result.sections_completed,
            "sections_total": result.sections_total,
            "tasks_completed": result.tasks_completed,
            "tasks_total": result.tasks_total,
            "elapsed_seconds": result.elapsed_seconds,
            "errors": result.errors,
        }
        Path(out_path).write_text(json.dumps(data, indent=2, default=str) + "\n")
        print(f"  ✓ Results saved to {out_path}")

    print()
    status = "✓ Complete!" if result.success else "✗ Completed with errors"
    print(f"  {status}")
    return 0 if result.success else 1
