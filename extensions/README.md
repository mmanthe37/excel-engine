# Excel Engine — Copilot CLI Extension

> **⚠️ DEPRECATED:** This Copilot CLI Extension has been replaced by the
> **MCP Server** at [`mcp-server/`](../mcp-server/). The MCP server is strictly
> more capable — it includes all the same tools plus `execute_openpyxl`,
> `execute_live`, and full MCP protocol support.
>
> **→ See [`mcp-server/README.md`](../mcp-server/README.md) for setup and usage.**

---

## Migration

The `excel-engine.py` extension now returns a deprecation error for all
requests. To migrate:

1. **Remove** any copies of `excel-engine.py` from `.github/extensions/` or
   `~/.copilot/extensions/`.
2. **Configure the MCP server** in `~/.copilot/mcp-config.json` or
   `.vscode/mcp.json`:

```json
{
  "mcpServers": {
    "excel-engine": {
      "command": "python3",
      "args": ["/absolute/path/to/excel-engine/mcp-server/server.py"],
      "env": {
        "PYTHONPATH": "/absolute/path/to/excel-engine"
      }
    }
  }
}
```

3. Reload MCP servers in the CLI with `/mcp`.

## License

Same license as the parent Excel Engine repository.
