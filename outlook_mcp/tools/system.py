"""System/utility MCP tools that delegate to server implementations."""

from __future__ import annotations

from typing import Any, Dict, Optional

from ..features import feature_gate, feature_metrics, reload_features

# Tools register via outlook_mcp.toolkit and access the shared MCP instance lazily.
from outlook_mcp.toolkit import get_current_mcp, mcp_tool, register_all_tools
from outlook_mcp.services.system import (
    build_params_payload,
    get_current_datetime as _service_get_current_datetime,
    get_profile_identity as _service_get_profile_identity,
)


@mcp_tool()
@feature_gate(group="system")
def params(
    protocolVersion: Optional[str] = None,
    capabilities: Optional[Dict[str, Any]] = None,
    clientInfo: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Ritorna metadati MCP/HTTP per handshake con client esterni."""
    # Delegate to the underlying implementation (which already filters
    # tool listing by enabled gates).
    return build_params_payload(
        mcp_instance=get_current_mcp(),
        protocol_version=protocolVersion,
        capabilities=capabilities,
        client_info=clientInfo,
    )


@mcp_tool()
@feature_gate(group="general")
def get_current_datetime(include_utc: bool = True) -> str:
    """Ritorna data/ora correnti; opzionalmente include i riferimenti UTC."""
    return _service_get_current_datetime(include_utc=include_utc)


@mcp_tool()
@feature_gate(group="system")
def get_profile_identity() -> str:
    """Restituisce nome visualizzato e indirizzi associati al profilo Outlook corrente."""
    payload = _service_get_profile_identity()
    if "error" in payload:
        return str(payload["error"])

    lines = ["Profilo Outlook corrente:"]
    display_name = payload.get("display_name") or "-"
    primary_address = payload.get("primary_address") or "-"
    lines.append(f"- Nome visualizzato: {display_name}")
    lines.append(f"- Indirizzo primario: {primary_address}")

    addresses = payload.get("addresses") or []
    if addresses:
        lines.append("- Alias riconosciuti:")
        for addr in addresses:
            lines.append(f"  - {addr}")
    else:
        lines.append("- Alias riconosciuti: Nessuno")

    accounts = payload.get("accounts") or []
    if accounts:
        lines.append("- Account configurati:")
        for account in accounts:
            display = account.get("display_name") or "-"
            smtp = account.get("smtp_address") or "-"
            lines.append(f"  - {display} ({smtp})")
    return "\n".join(lines)


@mcp_tool()
@feature_gate(group="system")
def reload_configuration() -> str:
    """Ricarica la configurazione runtime (features.json e variabili ambiente)."""
    reload_features()
    register_all_tools(get_current_mcp(), force=True)
    metrics = feature_metrics()
    enabled_groups = metrics["enabled_groups"] or ["-"]
    active_tools = metrics["active_tools"]
    lines = [
        "Configurazione ricaricata con successo.",
        f"Gruppi abilitati: {', '.join(enabled_groups)}",
        f"Tool attivi: {len(active_tools)}/{metrics['registered_tools']}",
    ]
    return "\n".join(lines)


@mcp_tool()
@feature_gate(group="system")
def feature_status() -> str:
    """Mostra un riepilogo delle configurazioni feature correnti."""
    metrics = feature_metrics()
    lines = [
        "Stato features:",
        f"- Gruppi abilitati: {', '.join(metrics['enabled_groups']) or '-'}",
        f"- Gruppi disabilitati: {', '.join(metrics['disabled_groups']) or '-'}",
        f"- Tool abilitati esplicitamente: {', '.join(metrics['enabled_tools']) or '-'}",
        f"- Tool disabilitati esplicitamente: {', '.join(metrics['disabled_tools']) or '-'}",
        f"- Tool attivi: {len(metrics['active_tools'])}/{metrics['registered_tools']}",
    ]
    return "\n".join(lines)
