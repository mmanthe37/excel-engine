# Excel Engine MCP Server

Model Context Protocol server for the **Excel Engine** — autonomous Excel automation on macOS using a 6-layer architecture.

## Tools

| Tool | Description |
|---|---|
| `complete_assignment` | Full autonomous assignment completion (parse → plan → execute → verify) |
| `parse_instructions` | Parse a .docx/.rtfd/.pdf/.txt instruction file into structured tasks |
| `execute_openpyxl` | Run Phase 1 offline operations (openpyxl only, Excel not required) |
| `execute_live` | Run Phase 2 live operations (xlwings, AppleScript, VBA, etc.) |
| `verify_workbook` | Verify assignment completion with pass/fail per task |
| `get_engine_status` | Get engine version, available layers, and config |

## Quick Start

```bash
# 1. Install the excel-engine package (editable)
cd ~/Dev/excel-engine
pip install -e .

# 2. Install MCP server dependencies
cd mcp-server
pip install mcp openpyxl

# 3. Run the server
python server.py
```

## Add to Copilot CLI

Copy the contents of `config.json` into your Copilot CLI MCP configuration at `~/.copilot/mcp-config.json`:

```json
{
  "mcpServers": {
    "excel-engine": {
      "command": "/Users/michaelmanthejr/Dev/excel-engine/.venv/bin/python",
      "args": ["/Users/michaelmanthejr/Dev/excel-engine/mcp-server/server.py"],
      "env": {}
    }
  }
}
```

Or for Claude Desktop, add to `~/Library/Application Support/Claude/claude_desktop_config.json`.

## Example Usage

Once connected, an AI assistant can call:

```
complete_assignment(
  workbook_path="~/Desktop/NP_EX_3-2.xlsx",
  instruction_path="~/Desktop/instructions.rtfd"
)
```

Or step-by-step:

```
# 1. Parse instructions
parse_instructions(instruction_path="~/Desktop/instructions.rtfd")

# 2. Run offline tasks
execute_openpyxl(workbook_path="~/Desktop/NP_EX_3-2.xlsx", tasks=["Enter =SUM(B2:B10) in cell B11"])

# 3. Run live tasks (Excel must be open)
execute_live(workbook_path="~/Desktop/NP_EX_3-2.xlsx", tasks=["Sort column A ascending"])

# 4. Verify
verify_workbook(workbook_path="~/Desktop/NP_EX_3-2.xlsx")
```

## Architecture

```
Copilot CLI / Claude ──stdio──▶ MCP Server ──▶ ExcelEngine
                                                  ├── Layer 1: openpyxl (offline)
                                                  ├── Layer 2: xlwings (live API)
                                                  ├── Layer 3: AppleScript
                                                  ├── Layer 4: System Events (UI)
                                                  ├── Layer 5: VBA injection
                                                  └── Layer 6: PyAutoGUI (fallback)
```
