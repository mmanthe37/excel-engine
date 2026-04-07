#!/usr/bin/env python3
"""
Example: Parse SAM instruction files into structured task lists.

Demonstrates reading instruction files in various formats (.docx, .rtfd,
.pdf, .txt) and extracting structured tasks for the Excel Engine.

Usage:
    python parse_sam_instructions.py /path/to/instructions.docx
"""

import sys
import subprocess
from pathlib import Path


def extract_text(path: Path) -> str:
    """Extract text from various instruction file formats."""
    suffix = path.suffix.lower()

    if suffix == ".docx":
        from docx import Document
        doc = Document(str(path))
        return "\n".join(p.text for p in doc.paragraphs)

    elif suffix in (".rtfd", ".rtf"):
        result = subprocess.run(
            ["textutil", "-convert", "txt", str(path), "-stdout"],
            capture_output=True, text=True,
        )
        if result.returncode == 0:
            return result.stdout
        # Fallback: read .rtf inside .rtfd package
        rtf_file = path / "TXT.rtf"
        if rtf_file.exists():
            result = subprocess.run(
                ["textutil", "-convert", "txt", str(rtf_file), "-stdout"],
                capture_output=True, text=True,
            )
            return result.stdout
        raise RuntimeError(f"Failed to extract text from {path}")

    elif suffix == ".pdf":
        try:
            import pdfplumber
            with pdfplumber.open(str(path)) as pdf:
                return "\n".join(
                    page.extract_text() or "" for page in pdf.pages
                )
        except ImportError:
            result = subprocess.run(
                ["pdftotext", str(path), "-"],
                capture_output=True, text=True,
            )
            return result.stdout

    elif suffix == ".txt":
        return path.read_text()

    else:
        raise ValueError(f"Unsupported format: {suffix}")


def parse_steps(text: str) -> list[dict]:
    """Parse instruction text into structured step dictionaries."""
    import re

    steps = []
    # Match numbered steps like "1.", "2.", "Step 1:", etc.
    pattern = r'(?:^|\n)\s*(?:Step\s+)?(\d+)[.):\s]+(.+?)(?=\n\s*(?:Step\s+)?\d+[.):\s]|\Z)'
    matches = re.findall(pattern, text, re.DOTALL | re.IGNORECASE)

    for num, body in matches:
        step = {
            "number": int(num),
            "text": body.strip(),
            "cells": [],
            "sheets": [],
            "task_type": "unknown",
        }

        # Extract cell references (A1, B2:B10, etc.)
        cell_refs = re.findall(r'\b([A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?)\b', body)
        step["cells"] = cell_refs

        # Detect task type from keywords
        body_lower = body.lower()
        if any(w in body_lower for w in ["formula", "function", "=sum", "=if", "=vlookup"]):
            step["task_type"] = "formula"
        elif "table" in body_lower and "create" in body_lower:
            step["task_type"] = "table_create"
        elif "table style" in body_lower:
            step["task_type"] = "table_style"
        elif "sort" in body_lower:
            step["task_type"] = "sort"
        elif "filter" in body_lower:
            step["task_type"] = "autofilter"
        elif "pivot" in body_lower:
            step["task_type"] = "pivot_table"
        elif "slicer" in body_lower:
            step["task_type"] = "slicer"
        elif "chart" in body_lower:
            step["task_type"] = "chart"
        elif any(w in body_lower for w in ["format", "currency", "percent", "number format"]):
            step["task_type"] = "number_format"
        elif "freeze" in body_lower:
            step["task_type"] = "freeze_panes"
        elif "conditional" in body_lower:
            step["task_type"] = "conditional_format"

        steps.append(step)

    return steps


def main():
    if len(sys.argv) < 2:
        print("Usage: parse_sam_instructions.py <instructions_file>")
        sys.exit(1)

    path = Path(sys.argv[1]).resolve()
    if not path.exists():
        print(f"Error: File not found: {path}")
        sys.exit(1)

    print(f"=== SAM Instruction Parser ===")
    print(f"File: {path.name}")
    print(f"Format: {path.suffix}")
    print()

    # Extract text
    text = extract_text(path)
    print(f"Extracted {len(text)} characters of text")
    print()

    # Parse steps
    steps = parse_steps(text)
    print(f"Found {len(steps)} steps:")
    print()

    for step in steps:
        cells = ", ".join(step["cells"]) if step["cells"] else "—"
        print(f"  Step {step['number']}: [{step['task_type']}]")
        print(f"    Cells: {cells}")
        print(f"    Text:  {step['text'][:100]}{'...' if len(step['text']) > 100 else ''}")
        print()

    # Try using the engine's parser if available
    try:
        from excel_engine.parsers.instruction_parser import InstructionParser
        from excel_engine.parsers.task_extractor import TaskExtractor

        parser = InstructionParser()
        engine_text = parser.parse(path)
        extractor = TaskExtractor()
        tasks = extractor.extract(engine_text)

        print(f"\n=== Engine Parser Results ===")
        print(f"Extracted {len(tasks)} tasks:")
        for task in tasks:
            print(f"  {task.id}: {task.task_type.value} → {task.sheet}!{task.cell or task.range}")
    except ImportError:
        print("\n(excel_engine not installed — using standalone parser only)")


if __name__ == "__main__":
    main()
