<div align="center">

# Excel Engine

### Autonomous Excel Automation for macOS

[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue.svg)](https://www.python.org/downloads/)
[![macOS](https://img.shields.io/badge/platform-macOS-lightgrey.svg)](https://www.apple.com/macos/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![PyPI](https://img.shields.io/pypi/v/excel-engine.svg)](https://pypi.org/project/excel-engine/)

**A 6-layer desktop automation engine that completes Microsoft Excel assignments autonomously on macOS.**

[Quick Start](#quick-start) · [Architecture](#architecture) · [Usage](#usage) · [Distribution Formats](#distribution-formats) · [Contributing](CONTRIBUTING.md)

</div>

---

## Features

- **6-Layer Automation Architecture** — Cascading strategy from lightweight file I/O to full desktop control
- **Section-Based Planning** — Reads assignment instructions, decomposes into ordered checkpoints, and executes them sequentially
- **114+ Checkpoint Trained** — Battle-tested against real SAM training modules covering formulas, formatting, charts, PivotTables, and more
- **Self-Healing Execution** — Automatic retry with layer escalation when a step fails
- **5 Distribution Formats** — Use as a Python library, CLI tool, MCP Server (recommended AI interface), Copilot CLI Extension *(deprecated — use MCP Server)*, or Copilot Plugin
- **macOS-Native** — Built specifically for Microsoft Excel for Mac 365 using AppleScript, System Events, and Accessibility APIs

## Quick Start

```bash
# Install from PyPI
pip3 install excel-engine

# Run against an assignment
excel-engine run --instructions assignment.rtfd --workbook workbook.xlsx
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
│  │ Planner  │→ │ Executor │→ │ Checkpoint Verify │  │
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
pip3 install excel-engine[mcp]

# Add to your MCP configuration
excel-engine mcp install
```

### Copilot CLI Extension *(Deprecated — use MCP Server)*

```bash
# DEPRECATED — use the MCP Server instead.
# See mcp-server/README.md for setup instructions.
```

### Copilot Plugin

```bash
# Register as a Copilot plugin
excel-engine plugin install
```

## Usage

### CLI

```bash
# Run a full assignment
excel-engine run \
  --instructions "Module 3 Instructions.rtfd" \
  --workbook "NP_EX_3-2.xlsx"

# Run a specific section
excel-engine run \
  --instructions "Module 3 Instructions.rtfd" \
  --workbook "NP_EX_3-2.xlsx" \
  --section 4

# Dry run — plan without executing
excel-engine plan \
  --instructions "Module 3 Instructions.rtfd"

# Verify a completed workbook
excel-engine verify \
  --workbook "NP_EX_3-2.xlsx" \
  --checkpoints checkpoints.json
```

### Python API

```python
from excel_engine import ExcelEngine

engine = ExcelEngine()

# Load assignment instructions
engine.load_instructions("Module 3 Instructions.rtfd")

# Open workbook
engine.open("NP_EX_3-2.xlsx")

# Plan and execute all sections
plan = engine.plan()
engine.execute(plan)

# Or execute a single checkpoint
engine.execute_checkpoint("Set cell B4 to =SUM(B5:B10)")
```

### MCP Server

When running as an MCP server, Excel Engine exposes tools for:

```json
{
  "tools": [
    "excel_engine_run",
    "excel_engine_plan",
    "excel_engine_verify",
    "excel_engine_execute_checkpoint"
  ]
}
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
| Python | 3.11+ |

### Python Packages

| Package | Purpose |
|---------|---------|
| `openpyxl` | Offline .xlsx file manipulation |
| `xlwings` | Live Excel application bridge |
| `pyautogui` | Desktop GUI automation |
| `python-docx` | Reading .docx/.rtfd instructions |
| `Pillow` | Screenshot capture and image matching |

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
├── excel_engine/           # Core package
│   ├── __init__.py
│   ├── engine.py           # Main orchestrator
│   ├── planner.py          # Instruction parser & planner
│   ├── layers/             # 6-layer automation stack
│   │   ├── openpyxl_layer.py
│   │   ├── xlwings_layer.py
│   │   ├── applescript_layer.py
│   │   ├── system_events_layer.py
│   │   ├── vba_layer.py
│   │   └── pyautogui_layer.py
│   ├── cli.py              # CLI entry point
│   ├── mcp_server.py       # MCP server
│   └── verify.py           # Checkpoint verification
├── docs/                   # Documentation
├── tests/                  # Test suite
├── pyproject.toml          # Project metadata
├── README.md
├── LICENSE
├── CONTRIBUTING.md
└── CHANGELOG.md
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, testing, and PR guidelines.

## License

[MIT](LICENSE) © 2026 Michael Manthe

openpyxl, xlwings, AppleScript, System Events, VBA via VBE, PyAutoGUI
