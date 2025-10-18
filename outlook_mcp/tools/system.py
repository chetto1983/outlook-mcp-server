"""System/utility MCP tools that delegate to server implementations."""

from __future__ import annotations

from typing import Any, Dict, Optional

from ..features import feature_gate
from ..features import is_tool_enabled, get_tool_group

# Import the FastMCP instance and the underlying implementations from the
# main server module to avoid duplicating logic.
from outlook_mcp_server import mcp
from outlook_mcp.services.system import build_params_payload, get_current_datetime as _service_get_current_datetime


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
    return build_params_payload(mcp_instance=mcp, protocol_version=protocolVersion, capabilities=capabilities, client_info=clientInfo)


@mcp.tool()
@feature_gate(group="general")
def get_current_datetime(include_utc: bool = True) -> str:
    """Ritorna data/ora correnti; opzionalmente include i riferimenti UTC."""
    return _service_get_current_datetime(include_utc=include_utc)
