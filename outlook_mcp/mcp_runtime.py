"""Shared FastMCP runtime state."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

DEFAULT_HOST = "0.0.0.0"
DEFAULT_PORT = 8000

# Single FastMCP instance reused across the process.
mcp = FastMCP("outlook-assistant", host=DEFAULT_HOST, port=DEFAULT_PORT)

__all__ = ["mcp", "DEFAULT_HOST", "DEFAULT_PORT"]
