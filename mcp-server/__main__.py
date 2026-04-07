"""Allow running with: python -m mcp_server"""
from server import mcp

mcp.run(transport="stdio")
