<div align="center">

# Excel Engine

### Autonomous Excel Automation — Cross-Platform GUI + macOS Desktop Engine

[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![macOS](https://img.shields.io/badge/platform-macOS-lightgrey.svg)](https://www.apple.com/macos/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![PyPI](https://img.shields.io/badge/PyPI-coming%20soon-orange)](https://github.com/mmanthe37/excel-engine)

**A 6-layer desktop automation engine that completes Microsoft Excel assignments autonomously. Cross-platform GUI for all platforms; full desktop automation on macOS.**

[Quick Start](#quick-start) · [Architecture](#architecture) · [Usage](#usage) · [Distribution Formats](#distribution-formats) · [Contributing](CONTRIBUTING.md)

</div>

---

## Features

- **6-Layer Automation Architecture** — Cascading strategy from lightweight file I/O to full desktop control
- **Section-Based Planning** — Reads assignment instructions, decomposes into ordered tasks, and executes them sequentially
- **114+ Task Patterns** — Battle-tested against real SAM training modules covering formulas, formatting, charts, PivotTables, and more
- **Formula Recalculation** — Optional LibreOffice-powered recalculation with error scanning (#REF!, #DIV/0!, etc.)
- **Financial Model Presets** — Investment banking color coding (blue=inputs, black=formulas, green=cross-sheet) and number formats
- **Parallel Execution** — Tasks on different sheets run concurrently via `ThreadPoolExecutor` for 30-50% speedup on multi-sheet workbooks
- **Circuit Breaker** — Layers that fail repeatedly are auto-skipped with half-open probe after configurable reset
- **Progress Callbacks** — Real-time per-task status events in GUI and MCP server
- **Self-Healing Execution** — Automatic retry with layer escalation when a step fails
- **Error Recovery** — Classified error handling with configurable retry and exponential backoff
- **7 Distribution Formats** — Python library, CLI, GUI App, MCP Server (Python), MCP Server (Go), Copilot Plugin, or Extension *(deprecated)*
- **Cross-Platform GUI** — Browser-based GUI works on macOS, Windows, and Linux — no terminal needed
- **macOS-Native Desktop Layers** — Layers 2–6 use AppleScript, System Events, and Accessibility APIs (macOS only); Layer 1 (openpyxl) works everywhere

## Quick Start

```bash
# Install from source (recommended)
git clone https://github.com/mmanthe37/excel-engine.git
cd excel-engine
python3 -m venv .venv && source .venv/bin/activate
pip install -e .

# Run against an assignment
excel-engine run assignment.xlsx instructions.docx
```

## Architecture

Excel Engine uses a 6-layer cascading architecture. Each layer is attempted in order; if a layer cannot complete a task, the engine escalates to the next.

```
┌─────────────────────────────────────────────────────┐
│                   Excel Engine                       │
│                                                      │
│  ┌───────────────────────────────────────────────┐  │
│  │  Layer 1: openpyxl         (offline file I/O) │  │
│  ├───────────────────────────────────────────────┤  │
│  │  Layer 2: xlwings        (live Excel bridge)  │  │
│  ├───────────────────────────────────────────────┤  │
│  │  Layer 3: AppleScript  (Excel-specific ops)   │  │
│  ├───────────────────────────────────────────────┤  │
│  │  Layer 4: System Events  (ribbon/dialog UI)   │  │
│  ├───────────────────────────────────────────────┤  │
│  │  Layer 5: VBA via VBE   (PivotTables, etc.)   │  │
│  ├───────────────────────────────────────────────┤  │
│  │  Layer 6: PyAutoGUI   (last-resort control)   │  │
│  └───────────────────────────────────────────────┘  │
│                                                      │
│  ┌──────────┐  ┌──────────┐  ┌───────────────────┐  │
│  │ Planner  │→ │ Executor │→ │     Verifier      │  │
│  └──────────┘  └──────────┘  └───────────────────┘  │
└─────────────────────────────────────────────────────┘
```

| Layer | Technology | Use Case | Speed |
|-------|-----------|----------|-------|
| 1 | openpyxl | Cell values, formulas, formatting, named ranges | ⚡ Fastest |
| 2 | xlwings | Live reads/writes, formula evaluation, sheet ops | ⚡ Fast |
| 3 | AppleScript | Save, print, page setup, Excel-specific commands | 🔶 Medium |
| 4 | System Events | Ribbon clicks, dialog interaction, menu navigation | 🔶 Medium |
| 5 | VBA via VBE | PivotTables, advanced charts, complex automation | 🐢 Slow |
| 6 | PyAutoGUI | Pixel-level clicks, drag-and-drop, visual search | 🐢 Slowest |

## Installation

> **Note:** This package is not yet published to PyPI. Install from source or directly from GitHub.

### Python Package (from source)

```bash
git clone https://github.com/mmanthe37/excel-engine.git
cd excel-engine
python3 -m venv .venv && source .venv/bin/activate
pip install -e .
```

### Python Package (from GitHub, no clone needed)

```bash
python3 -m venv ~/excel-env && source ~/excel-env/bin/activate
pip install git+https://github.com/mmanthe37/excel-engine.git
```

### CLI Application

```bash
# Installed automatically with the Python package
source .venv/bin/activate  # or whichever venv you used
excel-engine --help
```

### MCP Server

```bash
source .venv/bin/activate

# Run the MCP server via stdio transport
python mcp-server/server.py
```

### Copilot CLI Extension *(Deprecated — use MCP Server)*

```bash
# DEPRECATED — use the MCP Server instead.
# See mcp-server/README.md for setup instructions.
```

### GUI App (Easiest — No Terminal Needed)

The GUI app works on **macOS, Windows, and Linux**. No coding or terminal experience required.

1. Download the project: [Download ZIP](https://github.com/mmanthe37/excel-engine/archive/refs/heads/main.zip)
2. Unzip the folder
3. Double-click the launcher for your platform:
   - **macOS:** `gui/launch.command`
   - **Windows:** `gui/launch.bat`
   - **Linux:** `gui/launch.sh`

The app opens in your browser. Upload your Excel file + instructions, click Run, and download the result.

> See [`gui/README.md`](gui/README.md) for the full user guide.

### Copilot Plugin

See [`plugin/README.md`](plugin/README.md) for setup instructions.

## Distribution Formats

| # | Format | Description | Platform | How to Use |
|---|--------|-------------|----------|------------|
| 1 | Python Library | Import and use in your own scripts | All | `pip install -e .` → `from excel_engine import ExcelEngine` |
| 2 | CLI Tool | Command-line interface | All | `excel-engine run assignment.xlsx instructions.docx` |
| 3 | GUI App | Cross-platform browser GUI | macOS, Windows, Linux | Double-click launcher → browser UI |
| 4 | MCP Server (Python) | AI assistant integration (recommended) | All | `python mcp-server/server.py` |
| 5 | MCP Server (Go) | High-performance MCP bridge | All | `cd go-mcp-server && go build && ./go-mcp-server` |
| 6 | Copilot Plugin | GitHub Copilot Chat integration | All | See `plugin/README.md` |
| 7 | Copilot CLI Extension | *(Deprecated — removed in v1.1.0, use MCP Server)* | macOS | — |

> **Note:** Layers 2–6 of the automation engine (xlwings, AppleScript, System Events, VBA, PyAutoGUI) are macOS-only. The GUI and CLI use Layer 1 (openpyxl) on all platforms.

## Usage

### CLI

```bash
# Run a full assignment (positional: workbook then instructions)
excel-engine run assignment.xlsx instructions.docx

# Dry run — plan without executing
excel-engine run assignment.xlsx instructions.docx --dry-run

# Run only Phase 1 (openpyxl, offline)
excel-engine run assignment.xlsx instructions.docx --phase 1

# Write results to a JSON file
excel-engine run assignment.xlsx instructions.docx --output results.json

# Use a custom config file
excel-engine run assignment.xlsx instructions.docx --config engine.json

# Parse instructions into structured JSON tasks
excel-engine parse instructions.rtfd
excel-engine parse instructions.rtfd --output tasks.json

# Verify a completed workbook (basic structural check)
excel-engine verify completed.xlsx

# Verify against specific instructions (task-based)
excel-engine verify completed.xlsx --instructions instructions.docx

# Show engine version, layers, and configuration
excel-engine info

# Check macOS environment readiness
excel-engine check-env

# Enable verbose/debug output (works with any subcommand)
excel-engine --verbose run assignment.xlsx instructions.docx

# Real-time supervised run (task-by-task progress)
excel-engine run assignment.xlsx instructions.docx --watch
```

### Copilot CLI Supervised Run (repo main)

From a Copilot CLI session, use the repository wrapper so you always run the local
`main` checkout (not an older global/venv install):

```bash
cd ~/Dev/excel-engine
./copilot_excel_run.sh "/full/path/to/assignment.xlsx" "/full/path/to/instructions.docx"
```

You can pass any additional `excel-engine run` flags after the two paths:

```bash
./copilot_excel_run.sh "/path/assignment.xlsx" "/path/instructions.txt" --phase 1 --output run.json
```

By default, this wrapper also runs a final task-based verification pass. Add
`--skip-verify` only if you explicitly want execution without the final check.

### Copilot Autonomous Completion Protocol

Use this protocol for Copilot sessions that must autonomously complete workbooks from instruction files.

1. **Canonical repo/path**
   ```bash
   cd ~/excel-engine
   ```
2. **Activate environment**
   ```bash
   source .venv/bin/activate
   ```
3. **Preflight checks**
   ```bash
   python -m excel_engine.cli check-env
   python -m pytest -q tests/test_engine_fallback.py tests/test_execution_correctness.py
   ```
4. **Safe execution flow (instruction file + workbook)**
   ```bash
   ./copilot_excel_run.sh "/full/path/to/assignment.xlsx" "/full/path/to/instructions.docx" --output run.json
   ```
5. **False-positive completion guards**
   - Do **not** pass `--skip-verify` for autonomous runs.
   - Require wrapper exit code `0`.
   - If using `--output`, require `success: true`, empty `failed_tasks`, and no verification failures before declaring completion.
6. **Troubleshooting quick map**
   - `Worksheet '...' does not exist` → sheet target mismatch; fix extracted task/instruction target before rerun.
   - `Verification failed for task ...` → treat section/run as failed; rerun with `--watch`, inspect task target/value.
   - `Save validation failed ... missing [Content_Types].xml` → workbook save corruption risk; stop and rerun after checking file permissions/free space.
   - `xlwings ... Cannot find Excel` / Excel hung → open Excel manually, dismiss dialogs, rerun `check-env`.
   - Accessibility/System Events errors → grant terminal Accessibility permissions (see `docs/troubleshooting.md`).

### Python API

```python
from pathlib import Path
from excel_engine import ExcelEngine, EngineConfig

# Initialize with default config
engine = ExcelEngine(config=EngineConfig())

# Full pipeline: parse → plan → execute → verify
result = engine.run(
    workbook=Path("assignment.xlsx"),
    instructions=Path("instructions.docx"),
)
print(result.summary())
print(f"Success: {result.success}")
print(f"Tasks: {result.tasks_completed}/{result.tasks_total}")

# Or pass raw text instead of a file
result = engine.run(
    workbook=Path("assignment.xlsx"),
    instruction_text="Set cell B4 to =SUM(B5:B10)",
)

# Or pass pre-extracted tasks
tasks = engine.scan(instructions=Path("instructions.docx"))
plan = engine.plan(tasks)
print(plan.summary())

# Verify completion without re-executing
verification = engine.verify(
    workbook=Path("assignment.xlsx"),
    tasks=tasks,
)
print(f"Passed: {verification.pass_count}, Failed: {verification.fail_count}")
```

### MCP Server (Python)

The MCP server (`mcp-server/server.py`) exposes 6 tools for AI assistant integration:

| Tool | Description |
|------|-------------|
| `complete_assignment` | Full pipeline: parse instructions → plan → execute → verify |
| `parse_instructions` | Parse an instruction file into structured task JSON |
| `execute_openpyxl` | Run Phase 1 offline operations via openpyxl only |
| `execute_live` | Run Phase 2 live operations via xlwings/AppleScript/System Events/VBA/PyAutoGUI |
| `verify_workbook` | Verify assignment completion against expected tasks |
| `get_engine_status` | Get engine version, available layers, and configuration |

Run the server:

```bash
python mcp-server/server.py
```

### MCP Server (Go)

A high-performance Go implementation that bridges MCP protocol to the Python engine via subprocess. Same 6 tools as the Python server but compiled to a single binary with no Python runtime needed at the MCP layer.

```bash
cd go-mcp-server
go build -o go-mcp-server .
./go-mcp-server  # Communicates via stdio transport
```

Environment variables: `EXCEL_ENGINE_SERVER_NAME`, `EXCEL_ENGINE_VERSION`, `EXCEL_ENGINE_PYTHON_PATH`, `EXCEL_ENGINE_TIMEOUT`.

See [`go-mcp-server/README.md`](go-mcp-server/README.md) for details.

### Formula Recalculation

Excel Engine can optionally recalculate formulas using LibreOffice headless mode and scan for formula errors:

```python
from excel_engine import EngineConfig

config = EngineConfig(recalculate_formulas=True, recalc_timeout=60)
result = engine.run(workbook=Path("file.xlsx"), instructions=Path("instr.docx"))
if result.formula_errors:
    for err in result.formula_errors:
        print(f"  {err}")
```

If LibreOffice is not installed, recalculation is silently skipped. Formula error scanning (for #REF!, #DIV/0!, etc.) works without LibreOffice using openpyxl.

### Financial Presets

Opt-in formatting presets for investment banking and financial modeling standards:

```python
from excel_engine.presets.financial import apply_ib_colors, apply_number_formats

apply_ib_colors(workbook)       # Blue=inputs, black=formulas, green=cross-sheet
apply_number_formats(workbook)  # Currency, percentages, negative in parens
```

### Copilot CLI Extension *(Removed)*

> The Copilot CLI Extension was removed in v1.1.0. Use the MCP Server (Python or Go) for AI-powered interactions.

### Copilot Plugin

The Copilot plugin integrates directly into GitHub Copilot Chat, allowing natural language interaction:

> **You:** Complete section 3 of the SAM assignment  
> **Copilot:** Running Excel Engine against section 3...

## Requirements

### System

| Requirement | Version |
|------------|---------|
| macOS | 13.0+ (Ventura or later) |
| Microsoft Excel | Excel for Mac 365 |
| Python | 3.10+ |

### Python Packages

| Package | Purpose | Required |
|---------|---------|----------|
| `openpyxl` | Offline .xlsx file manipulation | ✅ Yes |
| `xlwings` | Live Excel application bridge | Optional (`pip install "excel-engine[live]"`) |
| `pyautogui` | Desktop GUI automation | Optional (`pip install "excel-engine[ui]"`) |
| `python-docx` | Reading .docx/.rtfd instructions | Optional (`pip install "excel-engine[parsers]"`) |
| `pdfplumber` | Reading .pdf instructions | Optional (`pip install "excel-engine[parsers]"`) |

Install everything: `pip install "excel-engine[all]"` (from inside your venv)

### macOS Permissions

Excel Engine requires the following macOS permissions (System Settings → Privacy & Security):

| Permission | Why |
|-----------|-----|
| **Accessibility** | Required for System Events UI automation and keyboard/mouse control |
| **Screen Recording** | Required for PyAutoGUI screenshot-based element detection |

> **Note:** You will be prompted to grant these permissions on first run. The terminal application or IDE running Excel Engine must be added to both permission lists.

## Project Structure

```
excel-engine/
├── excel_engine/               # Core Python package
│   ├── __init__.py
│   ├── engine.py               # Main orchestrator (ExcelEngine class)
│   ├── cli.py                  # CLI entry point (excel-engine command)
│   ├── config.py               # EngineConfig, Layer, TaskType enums
│   ├── interactive.py          # Interactive guided mode
│   ├── recovery.py             # Error recovery with retry and backoff
│   ├── recalc.py               # LibreOffice formula recalculation + error scanning
│   ├── layers/                 # 6-layer automation stack
│   │   ├── openpyxl_layer.py   # Layer 1: offline file I/O
│   │   ├── xlwings_layer.py    # Layer 2: live Excel bridge
│   │   ├── applescript_layer.py # Layer 3: Excel-specific commands
│   │   ├── system_events.py    # Layer 4: ribbon/dialog UI
│   │   ├── vba_layer.py        # Layer 5: VBA via VBE
│   │   └── pyautogui_layer.py  # Layer 6: last-resort desktop control
│   ├── parsers/                # Instruction parsing
│   │   ├── instruction_parser.py  # .docx/.rtfd/.pdf/.txt reader
│   │   └── task_extractor.py      # Natural language → Task objects
│   ├── planner/                # Execution planning
│   │   ├── dependency_graph.py # Task dependency resolution
│   │   └── task_planner.py     # Section-based ExecutionPlan builder
│   ├── presets/                # Formatting presets
│   │   └── financial.py        # IB color coding + number formats
│   ├── verifier/               # Completion verification
│   │   └── workbook_verifier.py # Cell/formula/format/structure checks
│   └── utils/                  # Shared utilities
│       ├── excel_constants.py  # Excel format constants
│       ├── mac_utils.py        # macOS-specific helpers
│       └── path_handler.py     # Path resolution and normalization
├── mcp-server/                 # MCP Server — Python (AI assistant interface)
│   ├── __main__.py
│   └── server.py               # FastMCP server with 6 tools
├── go-mcp-server/              # MCP Server — Go (high-performance bridge)
│   ├── main.go                 # Entry point with stdio transport
│   ├── go.mod
│   ├── config/config.go        # Environment-based configuration
│   └── tools/                  # Tool definitions + Python subprocess bridge
│       ├── registry.go         # 6 MCP tool registrations
│       ├── handlers.go         # Tool handler logic
│       └── bridge.go           # Python subprocess bridge
├── plugin/                     # Copilot Plugin
│   ├── install.sh
│   ├── agents/
│   │   └── excel-engine-agent.md
│   └── skills/
├── gui/                        # Cross-platform GUI app
│   ├── README.md               # Non-technical user guide
│   ├── app.py                  # Streamlit GUI application
│   ├── build_app.py            # Standalone app builder
│   ├── launch.command          # macOS launcher (double-click)
│   ├── launch.bat              # Windows launcher (double-click)
│   └── launch.sh               # Linux launcher (double-click)
├── tests/                      # Test suite (600+ tests)
├── docs/                       # Documentation
│   ├── architecture.md
│   ├── layers.md
│   └── troubleshooting.md
├── pyproject.toml
├── README.md
├── LICENSE
├── CONTRIBUTING.md
└── CHANGELOG.md
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, testing, and PR guidelines.

## License

[MIT](LICENSE) © 2026 Michael Manthe
