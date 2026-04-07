# Excel Engine — Copilot Plugin

Autonomous macOS desktop automation plugin for completing Microsoft Excel SAM (Skills Assessment Manager) assignments. Uses a 6-layer architecture built from 114+ checkpoints of real-world automation sessions.

## What It Does

Given an Excel workbook and an instruction file, this plugin autonomously:

1. **Parses** instruction files (.docx, .rtfd, .pdf, .txt)
2. **Extracts** structured tasks (formulas, formatting, charts, PivotTables, etc.)
3. **Plans** execution order grouped by worksheet and dependency
4. **Executes** each task using the most reliable automation layer
5. **Verifies** results after each section before proceeding

### The 6 Layers

| Layer | Technology | Use Case |
|-------|-----------|----------|
| 1 | **openpyxl** | Offline file manipulation — formulas, formatting, charts, conditional formatting |
| 2 | **xlwings** | Live Excel API — table creation, real-time verification, table styles |
| 3 | **AppleScript** | Excel-specific commands — sort, filter, freeze panes, hyperlinks, save |
| 4 | **System Events** | Ribbon/dialog UI — subtotals, slicer insertion, contextual tabs |
| 5 | **VBA via VBE** | PivotTables, slicers, PivotCharts via clipboard injection |
| 6 | **PyAutoGUI** | Last-resort desktop control with screenshot verification |

## Installation

### Quick Install

```bash
cd ~/Dev/excel-engine/plugin
chmod +x install.sh
./install.sh
```

### Manual Install

1. Copy plugin to Copilot plugins directory:
   ```bash
   cp -R ~/Dev/excel-engine/plugin ~/.copilot/installed-plugins/excel-engine
   ```

2. Install Python dependencies:
   ```bash
   pip install openpyxl xlwings pyautogui python-docx Pillow pdfplumber
   ```

3. Install the Excel Engine package:
   ```bash
   pip install -e ~/Dev/excel-engine
   ```

4. Grant macOS permissions:
   - **Accessibility**: System Preferences → Privacy & Security → Accessibility → Terminal ✓
   - **Screen Recording**: System Preferences → Privacy & Security → Screen Recording → Terminal ✓

## How to Trigger the Agent

Say any of these to Copilot:

- `"Complete this Excel assignment: workbook.xlsx using instructions.docx"`
- `"Do my SAM Module 3 for me"`
- `"Automate my Excel worksheet using the instruction file"`
- `"Complete NP_EX365_CS1-4B.xlsx following the instructions in the docx"`
- `"Continue completing my Excel Module 7"`
- `"Run the excel engine on my assignment"`

## Available Skills

### `excel-automation`
Core knowledge for all 6 automation layers including code patterns, anti-patterns, error recovery, and the complete task-to-layer decision framework.

### `instruction-parsing`
Patterns for extracting structured task lists from SAM instruction files in `.docx`, `.rtfd`, `.pdf`, and `.txt` formats. Includes step identification, task classification, and dependency detection.

## Requirements

| Requirement | Details |
|-------------|---------|
| **OS** | macOS (tested on Ventura, Sonoma, Sequoia) |
| **Excel** | Microsoft Excel for Mac 365 |
| **Python** | 3.10+ (3.12–3.13 recommended; 3.14 breaks uvx/typer) |
| **Packages** | openpyxl, xlwings, pyautogui, python-docx, Pillow, pdfplumber |
| **Permissions** | Accessibility + Screen Recording for Terminal |

## Project Structure

```
plugin/
├── README.md                    # This file
├── agents/
│   └── excel-engine-agent.md    # Comprehensive agent definition
├── skills/
│   ├── excel-automation/
│   │   ├── skill.md             # Excel automation skill
│   │   └── examples/
│   │       ├── complete_assignment.py
│   │       └── verify_workbook.py
│   └── instruction-parsing/
│       ├── skill.md             # Instruction parsing skill
│       └── examples/
│           └── parse_sam_instructions.py
└── install.sh                   # Automated installer
```

## Built From

This plugin encapsulates knowledge from 114+ checkpoints of real-world SAM assignment automation sessions, covering Excel modules for data tools, charts, PivotTables, slicers, 3-D references, VBA macro injection, and more.
