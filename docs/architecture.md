# Architecture

## Overview

Excel Engine is an autonomous desktop automation engine designed to complete Microsoft Excel assignments on macOS. It reads structured instructions (typically from `.rtfd` or `.docx` files), decomposes them into ordered checkpoints, and executes each checkpoint using a 6-layer cascading automation stack.

## Core Components

### Planner

The **Planner** parses assignment instruction documents and produces an execution plan:

1. **Document Parsing** — Extracts text and structure from `.rtfd` or `.docx` files using `python-docx`.
2. **Section Detection** — Identifies numbered sections and subsections.
3. **Checkpoint Extraction** — Breaks each section into atomic, verifiable checkpoints (e.g., "Set cell B4 to `=SUM(B5:B10)`").
4. **Dependency Resolution** — Orders checkpoints to respect data dependencies (e.g., a formula referencing a named range requires the named range to exist first).

### Executor

The **Executor** takes a plan and runs each checkpoint through the layer stack:

```
Checkpoint → Layer 1 → success? → done
                ↓ fail
             Layer 2 → success? → done
                ↓ fail
             Layer 3 → ...
                ↓ fail
             Layer 6 → success or raise LayerExhaustionError
```

Each layer returns a `LayerResult` indicating success, failure, or "not applicable" (the operation is outside the layer's capability).

### Verifier

The **Verifier** confirms that a checkpoint was completed correctly by inspecting the workbook state after execution. Verification strategies include:

- **Cell value comparison** — Read the cell and compare to expected value.
- **Formula verification** — Check that the cell contains the correct formula string.
- **Formatting inspection** — Verify font, fill, border, alignment properties.
- **Structural checks** — Confirm sheets, named ranges, tables, or charts exist.

## Execution Flow

```
Instructions (.rtfd)
        │
        ▼
   ┌─────────┐
   │ Planner  │ ── parse → section list → checkpoint list
   └────┬─────┘
        │
        ▼
   ┌──────────┐
   │ Executor  │ ── for each checkpoint:
   └────┬──────┘      try Layer 1..6 until success
        │
        ▼
   ┌──────────┐
   │ Verifier  │ ── confirm checkpoint result
   └────┬──────┘
        │
        ▼
   Completed Workbook (.xlsx)
```

## Layer Selection Strategy

The engine determines which layers are applicable for each checkpoint type:

| Checkpoint Type | Primary Layer | Fallback Layers |
|----------------|--------------|-----------------|
| Set cell value/formula | openpyxl | xlwings → AppleScript |
| Cell formatting | openpyxl | xlwings |
| Named ranges | openpyxl | xlwings |
| Sheet operations | openpyxl | xlwings → AppleScript |
| Page setup / print | AppleScript | System Events |
| Conditional formatting | xlwings | VBA |
| PivotTables | VBA | System Events → PyAutoGUI |
| Charts | xlwings | VBA → System Events |
| Ribbon commands | System Events | PyAutoGUI |
| Dialog interaction | System Events | PyAutoGUI |

## Error Handling

- **LayerExhaustionError** — All 6 layers failed. The checkpoint is logged as failed and the engine can optionally continue or abort.
- **WorkbookCorruptionError** — A layer corrupted the workbook. The engine restores from the last known-good state (automatic backup before each section).
- **ExcelNotRunningError** — Excel is not open or has crashed. The engine attempts to relaunch and reopen the workbook.

## State Management

The engine maintains state across checkpoints:

- **Workbook backups** — Created before each section for rollback.
- **Execution log** — Every layer attempt is logged with timing, result, and error details.
- **Checkpoint status** — Tracks which checkpoints are completed, failed, or pending.
