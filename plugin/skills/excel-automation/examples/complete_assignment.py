#!/usr/bin/env python3
"""
Example: Complete an Excel SAM assignment using the ExcelEngine.

Usage:
    python complete_assignment.py /path/to/workbook.xlsx /path/to/instructions.docx
"""

import sys
from pathlib import Path

from excel_engine import ExcelEngine
from excel_engine.config import EngineConfig


def main():
    if len(sys.argv) < 3:
        print("Usage: complete_assignment.py <workbook.xlsx> <instructions.docx>")
        sys.exit(1)

    workbook = Path(sys.argv[1]).resolve()
    instructions = Path(sys.argv[2]).resolve()

    if not workbook.exists():
        print(f"Error: Workbook not found: {workbook}")
        sys.exit(1)
    if not instructions.exists():
        print(f"Error: Instructions not found: {instructions}")
        sys.exit(1)

    # Configure the engine
    config = EngineConfig(
        verify_after_each_section=True,
        retina_display=True,
        sam_fingerprint_protected=True,
        max_retries=3,
    )

    engine = ExcelEngine(config=config)

    print(f"=== Excel Engine v1.0 ===")
    print(f"Workbook:     {workbook.name}")
    print(f"Instructions: {instructions.name}")
    print()

    # Run the engine
    result = engine.run(workbook=workbook, instructions=instructions)

    # Print results
    print()
    print(result.summary())

    if result.success:
        print(f"\n✅ Assignment completed successfully!")
        print(f"   File: {result.workbook_path}")
    else:
        print(f"\n⚠️  Assignment partially completed.")
        print(f"   {result.tasks_completed}/{result.tasks_total} tasks done.")
        if result.errors:
            print(f"   Errors:")
            for err in result.errors:
                print(f"     - {err}")

    sys.exit(0 if result.success else 1)


if __name__ == "__main__":
    main()
