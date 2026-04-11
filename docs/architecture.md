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
- **WorkbookCorruptionError** — A layer corrupted the workbook. The engine logs the failure and can optionally continue or abort.
- **ExcelNotRunningError** — Excel is not open or has crashed. The engine attempts to relaunch and reopen the workbook.
- **Error Recovery** — Errors are classified by type and severity, with configurable retry counts and exponential backoff between attempts.

## Phase 4: Formula Recalculation (v1.1.0)

After all sections are executed and verified, the engine optionally runs a recalculation phase:

1. **LibreOffice Recalc** — Uses headless `soffice` with a StarBasic macro to force-recalculate all formulas and save computed values.
2. **Formula Error Scan** — Two-pass openpyxl scan detects #REF!, #DIV/0!, #VALUE!, #NAME?, #NULL!, #NUM!, #N/A across all sheets (works without LibreOffice).
3. **Result Integration** — Formula errors are stored in `EngineResult.formula_errors` and displayed in the GUI.

Enable with `EngineConfig(recalculate_formulas=True)`. Gracefully skips if LibreOffice is not installed.

## Presets Subsystem (v1.1.0)

The `excel_engine/presets/` module provides opt-in formatting presets:

- **Financial (IB)** — `financial.py` applies investment banking color standards (blue=inputs, black=formulas, green=cross-sheet links, red=external, yellow=assumptions) and number formats (currency, percentages, negatives in parentheses).
- Presets are never applied automatically — they must be explicitly invoked.

## Parallel Execution (v1.1.0)

When `EngineConfig(parallel_execution=True)`, the executor groups tasks by target sheet:

- **Same-sheet tasks** execute serially (preserve ordering guarantees)
- **Cross-sheet tasks** with no inter-dependencies execute in parallel via `ThreadPoolExecutor(max_workers)`
- Default: disabled. Enable for multi-sheet workbooks to achieve 30-50% speedup.

## Circuit Breaker (v1.1.0)

A `CircuitBreaker` in `recovery.py` tracks per-layer failure counts:

- **Closed** — Normal operation, all layers available.
- **Open** — Layer has exceeded `circuit_breaker_threshold` consecutive failures → auto-skipped.
- **Half-open** — After `circuit_breaker_reset_seconds`, one probe attempt is allowed. If it succeeds, the breaker resets to Closed.

This prevents the engine from wasting time retrying a layer that is consistently failing (e.g., Excel not installed).

## Progress Callbacks (v1.1.0)

`engine.execute()` and `engine.run()` accept an optional `progress_callback: Callable[[dict], None]` that fires events:

- `{"task": "T3", "status": "executing", "layer": "openpyxl"}` — task started
- `{"task": "T3", "status": "completed", "passed": true}` — task finished with verification result

The Streamlit GUI and MCP server both wire this for real-time status display.

## State Management

The engine maintains state across tasks:

- **Execution log** — Every layer attempt is logged with timing, result, and error details.
- **Task status** — Tracks which tasks are completed, failed, or pending.
- **Verification results** — Per-section pass/fail tracking with detailed messages.
