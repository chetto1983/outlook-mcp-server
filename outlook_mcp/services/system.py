"""System/utility helpers shared between MCP server and tools."""

from __future__ import annotations

import datetime
from typing import Any, Dict, Optional

from outlook_mcp import logger
from outlook_mcp.features import get_tool_group, is_tool_enabled
from outlook_mcp.utils import coerce_bool

__all__ = ["build_params_payload", "get_current_datetime"]


def build_params_payload(
    mcp_instance: Any,
    protocol_version: Optional[str] = None,
    capabilities: Optional[Dict[str, Any]] = None,
    client_info: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Produce the payload returned by the legacy `params` MCP tool."""
    requested_version = protocol_version or "2025-03-26"
    logger.info(
        "params tool invocato (protocolVersion=%s, clientInfo=%s)",
        requested_version,
        client_info,
    )

    tool_summaries: Dict[str, Dict[str, Any]] = {}
    for tool in mcp_instance._tool_manager.list_tools():  # type: ignore[attr-defined]
        tool_group = get_tool_group(getattr(tool, "name", ""))
        if not is_tool_enabled(getattr(tool, "name", ""), tool_group):
            continue
        tool_summaries[tool.name] = {
            "description": getattr(tool, "description", None),
            "inputSchema": getattr(tool, "input_schema", None),
            "outputSchema": getattr(tool, "output_schema", None),
            "annotations": getattr(tool, "annotations", None),
        }

    default_capabilities = {"tools": {"list": True, "call": True}}
    response_capabilities: Dict[str, Any] = default_capabilities
    if capabilities:
        response_capabilities = {**capabilities}
        tools_caps = dict(default_capabilities.get("tools", {}))
        tools_caps.update(capabilities.get("tools", {}))
        response_capabilities["tools"] = tools_caps

    return {
        "protocolVersion": requested_version,
        "serverInfo": {
            "name": "outlook-assistant",
            "version": "1.0.0-http",
            "description": (
                "Bridge MCP per Outlook. Gli strumenti abilitano ricerche email, "
                "risposte rapide e consultazione calendario tramite HTTP."
            ),
        },
        "capabilities": response_capabilities,
        "tools": tool_summaries,
        "httpBridge": {
            "health": "GET /health",
            "listTools": "GET /tools",
            "invokeTool": "POST /tools/{tool_name}",
            "invokeToolRoot": "POST /",
        },
    }


def get_current_datetime(include_utc: bool = True) -> str:
    """Return formatted information about the current local/UTC time."""
    include_utc_bool = coerce_bool(include_utc)
    logger.info("get_current_datetime chiamato con include_utc=%s", include_utc_bool)
    try:
        local_dt = datetime.datetime.now()
        lines = [
            "Data e ora correnti:",
            f"- Locale: {local_dt.strftime('%Y-%m-%d %H:%M:%S')}",
            f"- Locale ISO: {local_dt.isoformat()}",
        ]
        if include_utc_bool:
            utc_dt = datetime.datetime.now(datetime.UTC).replace(tzinfo=datetime.timezone.utc)
            lines.append(f"- UTC: {utc_dt.strftime('%Y-%m-%d %H:%M:%S')}")
            lines.append(f"- UTC ISO: {utc_dt.isoformat()}")
        return "\n".join(lines)
    except Exception as exc:
        logger.exception("Errore durante get_current_datetime.")
        return f"Errore durante il calcolo della data/ora corrente: {exc}"

