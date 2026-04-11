# Excel Engine — Go MCP Server

A Go server that implements the [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) and bridges MCP tool calls to the Python **excel-engine** automation library via subprocess.

Go handles the MCP protocol, JSON-RPC transport, and tool schema validation. Python handles the actual Excel automation (openpyxl, xlwings, AppleScript, etc.).

## Architecture

```
Claude / Copilot / MCP Client
        │  (stdio JSON-RPC)
        ▼
   Go MCP Server
        │  (subprocess: python3 -c "...")
        ▼
  Python excel-engine
        │
        ▼
  Excel workbook (.xlsx)
```

1. The MCP client sends a `tools/call` request over stdio.
2. The Go server matches the tool name, validates input, and builds a Python script.
3. `pythonBridge()` runs the script as a `python3 -c` subprocess with a 120-second timeout.
4. The Python process imports `excel-engine`, performs the operation, and prints JSON to stdout.
5. The Go server returns the JSON result to the MCP client.

## Prerequisites

| Dependency | Version |
|---|---|
| Go | 1.23+ |
| Python | 3.10+ |
| excel-engine | Installed and importable (`pip install -e .` from repo root) |

## Build

```bash
cd go-mcp-server
go build -o go-mcp-server .
```

## Run

The server communicates over **stdio** (stdin/stdout JSON-RPC), which is the standard MCP transport for local tools.

```bash
./go-mcp-server
```

### Environment Variables

| Variable | Default | Description |
|---|---|---|
| `SERVER_NAME` | `excel-engine-go` | Server name reported in MCP handshake |
| `VERSION` | `v1.1.0` | Server version |
| `LOG_LEVEL` | `info` | Log verbosity |
| `PYTHON_PATH` | `python3` | Path to the Python interpreter |
| `ENGINE_ROOT` | *(empty)* | Path to the excel-engine repo root (added to `sys.path`) |

Example with custom Python path:

```bash
PYTHON_PATH=/opt/homebrew/bin/python3.12 \
ENGINE_ROOT=$HOME/Dev/excel-engine \
./go-mcp-server
```

### MCP Client Configuration

Add to your MCP client config (e.g. Claude Desktop, Copilot CLI):

```json
{
  "mcpServers": {
    "excel-engine": {
      "command": "/path/to/go-mcp-server",
      "env": {
        "ENGINE_ROOT": "/path/to/excel-engine"
      }
    }
  }
}
```

## Tools

| Tool | Description |
|---|---|
| `complete_assignment` | End-to-end: parse instructions, execute all tasks across 6 automation layers, and verify results. |
| `parse_instructions` | Parse a `.docx`, `.pdf`, or `.txt` instruction file into structured Excel tasks. |
| `execute_openpyxl` | Run offline operations via the openpyxl layer (Phase 1). Excel does not need to be open. |
| `execute_live` | Run live Excel operations via xlwings, AppleScript, System Events, VBA, or PyAutoGUI (Phase 2). Requires Excel to be open on macOS. |
| `verify_workbook` | Verify task completion against expected results, or audit workbook structure. |
| `get_engine_status` | Return engine version, available layers, supported task types, and current configuration. |

## Tests

```bash
go test ./... -v
```

Tests cover:
- **`tools/bridge_test.go`** — Python bridge subprocess (success, failure, invalid JSON, empty output, context cancellation, timeout), path resolution, and helper functions.
- **`main_test.go`** — Server initialization, tool registration (all 6 tools), and config loading from environment variables.

## Project Structure

```
go-mcp-server/
├── main.go              # Entry point — creates server, registers tools, runs stdio transport
├── main_test.go         # Server and config integration tests
├── config/
│   └── config.go        # Environment-based configuration
├── tools/
│   ├── bridge.go        # pythonBridge() — subprocess execution and JSON validation
│   ├── bridge_test.go   # Bridge unit tests
│   ├── handlers.go      # Tool handler functions and input types
│   └── registry.go      # RegisterTools() — wires handlers to MCP server
├── resources/           # MCP resources (reserved)
├── go.mod
└── go.sum
```

## License

See the repository root for license information.
