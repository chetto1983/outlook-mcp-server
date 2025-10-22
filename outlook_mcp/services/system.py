"""System/utility helpers shared between MCP server and tools."""

from __future__ import annotations

import datetime
from typing import Any, Dict, Optional

from outlook_mcp import logger
from outlook_mcp.features import get_tool_group, is_tool_enabled
from outlook_mcp.utils import coerce_bool

from outlook_mcp import connect_to_outlook
from outlook_mcp.services.email import collect_user_addresses, normalize_email_address

__all__ = ["build_params_payload", "get_current_datetime", "get_profile_identity"]


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


def _safe_collect_accounts(namespace) -> list[dict[str, str]]:
    """Return display name and SMTP address for each Outlook account."""
    accounts_info: list[dict[str, str]] = []
    try:
        session = namespace.Application.Session
        accounts = getattr(session, "Accounts", None)
        if not accounts:
            return accounts_info
        for idx in range(1, accounts.Count + 1):
            account = accounts.Item(idx)
            display_name = getattr(account, "DisplayName", None)
            smtp_address = getattr(account, "SmtpAddress", None)
            info: dict[str, str] = {}
            if display_name:
                info["display_name"] = str(display_name)
            if smtp_address:
                info["smtp_address"] = str(smtp_address)
            if info:
                accounts_info.append(info)
    except Exception:
        logger.debug("Impossibile enumerare gli account Outlook.", exc_info=True)
    return accounts_info


def get_profile_identity() -> dict[str, object]:
    """Return display name, primary address and aliases for the active Outlook profile."""
    try:
        _, namespace = connect_to_outlook()
    except Exception as exc:  # pragma: no cover - Outlook COM guarded
        logger.exception("Connessione Outlook fallita durante get_profile_identity.")
        return {"error": f"Impossibile connettersi a Outlook: {exc}"}

    display_name: str | None = None
    primary_address: str | None = None
    try:
        current_user = getattr(namespace, "CurrentUser", None)
        if current_user:
            display_name = str(getattr(current_user, "Name", "") or "") or None
            address_entry = getattr(current_user, "AddressEntry", None)
            if address_entry:
                primary_candidate = getattr(address_entry, "Address", None)
                if primary_candidate:
                    primary_address = normalize_email_address(str(primary_candidate))
                try:
                    exchange_user = address_entry.GetExchangeUser()
                    if exchange_user:
                        primary_smtp = getattr(exchange_user, "PrimarySmtpAddress", None)
                        if primary_smtp:
                            primary_address = normalize_email_address(str(primary_smtp)) or primary_address
                except Exception:
                    logger.debug("Impossibile ottenere ExchangeUser per CurrentUser.", exc_info=True)
    except Exception:
        logger.debug("Impossibile leggere CurrentUser da Outlook.", exc_info=True)

    addresses = sorted(collect_user_addresses(namespace))
    normalized_addresses = sorted(
        {addr for addr in (normalize_email_address(address) for address in addresses) if addr}
    )

    if not primary_address and normalized_addresses:
        primary_address = normalized_addresses[0]

    if not display_name:
        try:
            session = namespace.Application.Session
            if session and getattr(session, "AccountName", None):
                display_name = str(session.AccountName)
        except Exception:
            logger.debug("Impossibile recuperare AccountName dalla sessione.", exc_info=True)
        if not display_name and normalized_addresses:
            display_name = normalized_addresses[0]

    account_entries = _safe_collect_accounts(namespace)
    return {
        "display_name": display_name,
        "primary_address": primary_address,
        "addresses": normalized_addresses or addresses,
        "accounts": account_entries,
    }
