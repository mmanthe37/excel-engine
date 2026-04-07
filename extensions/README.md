# Excel Engine — Copilot CLI Extension

A GitHub Copilot CLI extension that exposes the **Excel Engine** autonomous
Excel-automation system as callable tools. The Copilot CLI agent can parse
assignment instructions, execute Excel tasks across 6 automation layers, and
verify workbook completion — all from your terminal.

## Requirements

| Requirement          | Version       |
| -------------------- | ------------- |
| Python               | 3.11+         |
| macOS                | 13 Ventura+   |
| Microsoft Excel      | for Mac 16.x+ |
| GitHub Copilot CLI   | latest        |

The Excel Engine package and its dependencies (`openpyxl`, `xlwings`, `pyautogui`)
must be installed. From the repo root:

```bash
pip install -e .
```

## Installation

### Option A — Add to any repository (project-level)

Copy the extension into a repo's `.github/extensions/` directory:

```bash
mkdir -p <your-repo>/.github/extensions
cp excel-engine.py <your-repo>/.github/extensions/
```

### Option B — User-level (available everywhere)

Copy to your Copilot CLI config directory:

```bash
mkdir -p ~/.copilot/extensions
cp excel-engine.py ~/.copilot/extensions/
```

### Option C — MCP server integration

Register the extension as an MCP stdio server in your Copilot CLI MCP config
(`~/.copilot/mcp-config.json` or `.vscode/mcp.json`):

```json
{
  "mcpServers": {
    "excel-engine": {
      "command": "python3",
      "args": ["/absolute/path/to/excel-engine/extensions/excel-engine.py"],
      "env": {
        "PYTHONPATH": "/absolute/path/to/excel-engine"
      }
    }
  }
}
```

Then reload MCP servers in the CLI with `/mcp`.

## Available Tools

### `complete_excel_assignment`

Run the full automation pipeline on an Excel assignment.

| Parameter          | Type   | Required | Description                                  |
| ------------------ | ------ | -------- | -------------------------------------------- |
| `workbook_path`    | string | yes      | Path to the `.xlsx` workbook                 |
| `instruction_path` | string | yes      | Path to the instruction file                 |

**Returns:** success/failure, task counts, section verifications, timing, and a
human-readable summary.

### `parse_excel_instructions`

Parse an instruction document and extract structured tasks without executing.

| Parameter          | Type   | Required | Description                                  |
| ------------------ | ------ | -------- | -------------------------------------------- |
| `instruction_path` | string | yes      | Path to `.docx`, `.rtfd`, `.pdf`, or `.txt`  |

**Returns:** list of tasks with types, target cells, formulas, dependencies,
and a task-type summary.

### `verify_excel_workbook`

Verify whether a workbook meets the assignment requirements.

| Parameter          | Type   | Required | Description                                  |
| ------------------ | ------ | -------- | -------------------------------------------- |
| `workbook_path`    | string | yes      | Path to the `.xlsx` workbook                 |
| `instruction_path` | string | no       | Path to instructions for task-level checks   |

**Returns:** per-task pass/fail results (with instructions) or a structural
sheet report (without instructions).

### `get_excel_engine_info`

Return engine version, automation layers, and supported task types.

**No parameters.**

**Returns:** version, layer descriptions, supported task types, instruction
formats, and platform requirements.

### `list_tools`

Discovery endpoint — returns the full tool manifest.

**No parameters.**

## Protocol

The extension uses a **JSON stdin/stdout** protocol:

**Request:**
```json
{
  "tool": "complete_excel_assignment",
  "params": {
    "workbook_path": "~/Desktop/NP_EX_3-2.xlsx",
    "instruction_path": "~/Documents/instructions.rtfd"
  }
}
```

**Response:**
```json
{
  "status": "success",
  "result": {
    "success": true,
    "tasks_completed": 24,
    "tasks_total": 24,
    "elapsed_seconds": 45.3,
    "summary": "✓ SUCCESS: NP_EX_3-2.xlsx\n  Sections: 5/5\n  Tasks: 24/24\n  Time: 45.3s"
  }
}
```

## Usage Examples

### From the command line

```bash
# Discover available tools
echo '' | python3 extensions/excel-engine.py

# Run a full assignment
echo '{"tool":"complete_excel_assignment","params":{"workbook_path":"assignment.xlsx","instruction_path":"instructions.rtfd"}}' \
  | python3 extensions/excel-engine.py

# Parse instructions only
echo '{"tool":"parse_excel_instructions","params":{"instruction_path":"instructions.docx"}}' \
  | python3 extensions/excel-engine.py

# Verify a workbook
echo '{"tool":"verify_excel_workbook","params":{"workbook_path":"assignment.xlsx"}}' \
  | python3 extensions/excel-engine.py

# Get engine info
echo '{"tool":"get_excel_engine_info","params":{}}' \
  | python3 extensions/excel-engine.py
```

### From the Copilot CLI

Once configured as an MCP server, ask the agent naturally:

> "Complete my Excel assignment in NP_EX_3-2.xlsx using the instructions in the SAM Module 3 folder"

> "Parse the instructions in instructions.rtfd and show me what tasks were found"

> "Verify whether NP_EX_3-2.xlsx is complete"

## Architecture

The extension wraps the **Excel Engine** which implements **Engine Protocol v2.0**:

```
SCAN → GROUP → EXECUTE → VERIFY
```

Six automation layers handle different task types with automatic fallback:

| Layer | Technology     | Strength                         |
| ----- | -------------- | -------------------------------- |
| 1     | openpyxl       | Offline file manipulation        |
| 2     | xlwings        | Live Excel API bridge            |
| 3     | AppleScript    | Excel-specific commands          |
| 4     | System Events  | Ribbon/dialog UI automation      |
| 5     | VBA            | Macro injection via VBE          |
| 6     | PyAutoGUI      | Last-resort desktop control      |

## License

Same license as the parent Excel Engine repository.
