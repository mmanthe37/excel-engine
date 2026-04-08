#!/usr/bin/env python3
"""
Excel Engine — GitHub Copilot CLI Extension (DEPRECATED)

This extension has been replaced by the MCP server at mcp-server/server.py.
It remains as a thin wrapper that directs users to the MCP server.
"""

import json
import sys

_DEPRECATION_MSG = (
    "DEPRECATED: The Copilot CLI Extension (extensions/excel-engine.py) has been "
    "replaced by the MCP server. Configure the MCP server instead:\n"
    "\n"
    "  1. See mcp-server/README.md for setup instructions.\n"
    "  2. Add to ~/.copilot/mcp-config.json or .vscode/mcp.json:\n"
    '     { "mcpServers": { "excel-engine": { "command": "python3",\n'
    '       "args": ["<repo>/mcp-server/server.py"] } } }\n'
)

print(_DEPRECATION_MSG, file=sys.stderr)

raw = sys.stdin.read().strip()
response = {
    "status": "error",
    "error": (
        "This Copilot CLI Extension is deprecated. "
        "Please use the MCP server at mcp-server/server.py instead. "
        "See mcp-server/README.md for configuration instructions."
    ),
}
print(json.dumps(response, indent=2))
