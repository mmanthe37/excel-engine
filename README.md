<div align="center">

# Excel Engine

### Autonomous Excel Automation for macOS

[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![macOS](https://img.shields.io/badge/platform-macOS-lightgrey.svg)](https://www.apple.com/macos/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![PyPI](https://img.shields.io/pypi/v/excel-engine.svg)](https://pypi.org/project/excel-engine/)

**A 6-layer desktop automation engine that completes Microsoft Excel assignments autonomously on macOS.**

[Quick Start](#quick-start) · [Architecture](#architecture) · [Usage](#usage) · [Distribution Formats](#distribution-formats) · [Contributing](CONTRIBUTING.md)

</div>

---

## Features

- **6-Layer Automation Architecture** — Cascading strategy from lightweight file I/O to full desktop control
- **Section-Based Planning** — Reads assignment instructions, decomposes into ordered tasks, and executes them sequentially
- **114+ Task Patterns** — Battle-tested against real SAM training modules covering formulas, formatting, charts, PivotTables, and more
- **Self-Healing Execution** — Automatic retry with layer escalation when a step fails
- **Error Recovery** — Classified error handling with configurable retry and exponential backoff
- **5 Distribution Formats** — Use as a Python library, CLI tool, MCP Server (recommended AI interface), Copilot CLI Extension *(deprecated — use MCP Server)*, or Copilot Plugin
- **macOS-Native** — Built specifically for Microsoft Excel for Mac 365 using AppleScript, System Events, and Accessibility APIs

## Quick Start

```bash
# Install from PyPI
pip3 install excel-engine

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

### Python Package

```bash
pip3 install excel-engine
```

### CLI Application

```bash
# Installed automatically with the Python package
pip3 install excel-engine
excel-engine --help
```

### MCP Server

```bash
pip3 install excel-engine

# Run the MCP server via stdio transport
python mcp-server/server.py
```

### Copilot CLI Extension *(Deprecated — use MCP Server)*

```bash
# DEPRECATED — use the MCP Server instead.
# See mcp-server/README.md for setup instructions.
```

### Copilot Plugin

See [`plugin/README.md`](plugin/README.md) for setup instructions.

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
```

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

### MCP Server

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

### Copilot CLI Extension *(Deprecated)*

> **Note:** The Copilot CLI Extension is deprecated. Use the MCP Server for
> AI-powered interactions. See [`mcp-server/README.md`](mcp-server/README.md).

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
| `xlwings` | Live Excel application bridge | Optional (`pip3 install excel-engine[live]`) |
| `pyautogui` | Desktop GUI automation | Optional (`pip3 install excel-engine[ui]`) |
| `python-docx` | Reading .docx/.rtfd instructions | Optional (`pip3 install excel-engine[parsers]`) |
| `pdfplumber` | Reading .pdf instructions | Optional (`pip3 install excel-engine[parsers]`) |

Install everything: `pip3 install excel-engine[all]`

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
│   ├── layers/                 # 6-layer automation stack
│   │   ├── __init__.py
│   │   ├── openpyxl_layer.py   # Layer 1: offline file I/O
│   │   ├── xlwings_layer.py    # Layer 2: live Excel bridge
│   │   ├── applescript_layer.py # Layer 3: Excel-specific commands
│   │   ├── system_events.py    # Layer 4: ribbon/dialog UI
│   │   ├── vba_layer.py        # Layer 5: VBA via VBE
│   │   └── pyautogui_layer.py  # Layer 6: last-resort desktop control
│   ├── parsers/                # Instruction parsing
│   │   ├── __init__.py
│   │   ├── instruction_parser.py  # .docx/.rtfd/.pdf/.txt reader
│   │   └── task_extractor.py      # Natural language → Task objects
│   ├── planner/                # Execution planning
│   │   ├── __init__.py
│   │   ├── dependency_graph.py # Task dependency resolution
│   │   └── task_planner.py     # Section-based ExecutionPlan builder
│   ├── verifier/               # Completion verification
│   │   ├── __init__.py
│   │   └── workbook_verifier.py # Cell/formula/format/structure checks
│   └── utils/                  # Shared utilities
│       ├── __init__.py
│       ├── excel_constants.py  # Excel format constants
│       ├── mac_utils.py        # macOS-specific helpers
│       └── path_handler.py     # Path resolution and normalization
├── mcp-server/                 # MCP Server (AI assistant interface)
│   ├── __main__.py
│   └── server.py               # FastMCP server with 6 tools
├── extensions/                 # Copilot CLI Extension (deprecated)
│   ├── excel-engine.py
│   └── README.md
├── plugin/                     # Copilot Plugin
│   ├── install.sh
│   ├── README.md
│   ├── agents/
│   │   └── excel-engine-agent.md
│   └── skills/
│       ├── excel-automation/
│       └── instruction-parsing/
├── tests/                      # Test suite
│   ├── conftest.py
│   ├── test_engine.py
│   ├── test_parsers.py
│   ├── test_planner.py
│   ├── test_layers.py
│   ├── test_integration.py
│   └── test_e2e_sam.py
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
