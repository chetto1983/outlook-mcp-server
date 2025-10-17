"""System/utility MCP tools that delegate to server implementations."""

from __future__ import annotations

from typing import Any, Dict, Optional

from ..features import feature_gate
from ..features import is_tool_enabled, get_tool_group

# Import the FastMCP instance and the underlying implementations from the
# main server module to avoid duplicating logic.
from outlook_mcp_server import mcp, params as _params_impl, get_current_datetime as _get_current_datetime_impl


@mcp.tool()
@feature_gate(group="system")
def params(
    protocolVersion: Optional[str] = None,
    capabilities: Optional[Dict[str, Any]] = None,
    clientInfo: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Ritorna metadati MCP/HTTP per handshake con client esterni."""
    # Delegate to the underlying implementation (which already filters
    # tool listing by enabled gates).
    return _params_impl(protocolVersion=protocolVersion, capabilities=capabilities, clientInfo=clientInfo)


@mcp.tool()
@feature_gate(group="general")
def get_current_datetime(include_utc: bool = True) -> str:
    """Ritorna data/ora correnti; opzionalmente include i riferimenti UTC."""
    return _get_current_datetime_impl(include_utc=include_utc)
