import argparse
import datetime
import os
from typing import Any, Dict, List, Optional, Set, Tuple

from mcp.server.fastmcp import FastMCP, Context
from mcp.server.fastmcp.exceptions import ToolError

from outlook_mcp import (
    ATTACHMENT_NAME_PREVIEW_MAX,
    BODY_PREVIEW_MAX_CHARS,
    CONVERSATION_ID_PREVIEW_MAX,
    DEFAULT_CONVERSATION_SAMPLE_LIMIT,
    DEFAULT_DOMAIN_ROOT_NAME,
    DEFAULT_DOMAIN_SUBFOLDERS,
    DEFAULT_MAX_RESULTS,
    LAST_VERB_REPLY_CODES,
    MAX_CONVERSATION_LOOKBACK_DAYS,
    MAX_DAYS,
    MAX_EMAIL_SCAN_PER_FOLDER,
    MAX_EVENT_LOOKAHEAD_DAYS,
    PENDING_SCAN_MULTIPLIER,
    PR_LAST_VERB_EXECUTED,
    PR_LAST_VERB_EXECUTION_TIME,
    calendar_cache,
    clear_calendar_cache,
    clear_email_cache,
    connect_to_outlook,
    email_cache,
    logger,
)
from outlook_mcp.utils import coerce_bool
from outlook_mcp import folders as folder_service
from outlook_mcp.features import feature_gate, is_tool_enabled, get_tool_group
from outlook_mcp.services.email import (
    resolve_mail_item,
    update_cached_email,
    normalize_email_address,
    extract_email_domain,
    derive_sender_email,
    ensure_domain_folder_structure,
    collect_user_addresses,
    mail_item_marked_replied,
    format_email,
    get_emails_from_folder,
    resolve_additional_folders,
    get_all_mail_folders,
    collect_emails_across_folders,
    get_related_conversation_emails,
    email_has_user_reply,
    email_has_user_reply_with_context,
    build_conversation_outline,
    present_email_listing,
    apply_categories_to_item,
    get_email_context as _service_get_email_context,
)
from outlook_mcp.services.calendar import (
    get_all_calendar_folders,
    get_calendar_folder_by_name,
    format_calendar_event,
    get_events_from_folder,
    collect_events_across_calendars,
    present_event_listing,
)
from outlook_mcp.services.system import (
    build_params_payload,
    get_current_datetime as _service_get_current_datetime,
)
from outlook_mcp.services.common import parse_datetime_string

try:
    from fastapi import Body, FastAPI, HTTPException
    from pydantic import BaseModel, Field
    import uvicorn
except ImportError:  # Optional dependencies loaded only for HTTP mode
    Body = None  # type: ignore[assignment]
    FastAPI = None  # type: ignore[assignment]
    HTTPException = None  # type: ignore[assignment]
    BaseModel = None  # type: ignore[assignment]
    Field = None  # type: ignore[assignment]
    uvicorn = None  # type: ignore[assignment]


# Initialize FastMCP server
mcp = FastMCP("outlook-assistant", host="0.0.0.0", port=8000)

# Ensure a stable module alias so that tool modules importing
# 'outlook_mcp_server' reference this same module even when executed as a script
import sys as _sys
_sys.modules.setdefault("outlook_mcp_server", _sys.modules[__name__])

# Backwards-compatible aliases for shared helpers
get_folder_by_name = folder_service.get_folder_by_name
get_folder_by_path = folder_service.get_folder_by_path
resolve_folder = folder_service.resolve_folder


def _register_tools() -> None:
    """Import tool modules so their @mcp.tool() decorators register tools.

    Note: modules import helpers from this file, so run after those helpers exist.
    """
    try:
        import importlib

        modules = [
            "outlook_mcp.tools.system",
            "outlook_mcp.tools.folders",
            "outlook_mcp.tools.email_list",
            "outlook_mcp.tools.email_detail",
            "outlook_mcp.tools.email_actions",
            "outlook_mcp.tools.attachments",
            "outlook_mcp.tools.contacts",
            "outlook_mcp.tools.calendar_read",
            "outlook_mcp.tools.calendar_write",
            "outlook_mcp.tools.domain_rules",
        ]
        for name in modules:
            try:
                logger.info("Registrazione tool MCP: import %s", name)
                importlib.import_module(name)
            except Exception:
                logger.exception("Errore durante l'import del modulo tool %s.", name)
        try:
            count = len(mcp._tool_manager.list_tools())  # type: ignore[attr-defined]
            logger.info("Registrazione tool MCP completata: %s tool", count)
        except Exception:
            pass
    except Exception:
        logger.exception("Registrazione dei tool MCP fallita.")

# Email and calendar helper implementations sono ora forniti da outlook_mcp.services
_resolve_mail_item = resolve_mail_item
_update_cached_email = update_cached_email
_normalize_email_address = normalize_email_address
_extract_email_domain = extract_email_domain
_derive_sender_email = derive_sender_email
_ensure_domain_folder_structure = ensure_domain_folder_structure
_collect_user_addresses = collect_user_addresses
_mail_item_marked_replied = mail_item_marked_replied
_present_email_listing = present_email_listing
_present_event_listing = present_event_listing
_email_has_user_reply = email_has_user_reply
_email_has_user_reply_with_context = email_has_user_reply_with_context
_build_conversation_outline = build_conversation_outline
_apply_categories_to_item = apply_categories_to_item
_parse_datetime_string = parse_datetime_string

def params(
    protocolVersion: Optional[str] = None,  # type: ignore[non-literal-used]
    capabilities: Optional[Dict[str, Any]] = None,
    clientInfo: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Provide handshake metadata compatible with MCP-aware HTTP clients (e.g. n8n).
    """
    return build_params_payload(
        mcp_instance=mcp,
        protocol_version=protocolVersion,
        capabilities=capabilities,
        client_info=clientInfo,
    )


def get_current_datetime(include_utc: bool = True) -> str:
    """Compatibilita' retro: delega al servizio di sistema."""
    return _service_get_current_datetime(include_utc=include_utc)


# MCP Tools
def list_folders(
    root_folder_id: Optional[str] = None,
    root_folder_path: Optional[str] = None,
    root_folder_name: Optional[str] = None,
    max_depth: int = 2,
    include_counts: bool = True,
    include_ids: bool = False,
    include_store: bool = False,
    include_paths: bool = True,
) -> str:
    from outlook_mcp.tools.folders import list_folders as _impl
    return _impl(
        root_folder_id=root_folder_id,
        root_folder_path=root_folder_path,
        root_folder_name=root_folder_name,
        max_depth=max_depth,
        include_counts=include_counts,
        include_ids=include_ids,
        include_store=include_store,
        include_paths=include_paths,
    )

def get_folder_metadata(
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
    include_children: bool = False,
    max_children: int = 20,
    include_counts: bool = True,
) -> str:
    from outlook_mcp.tools.folders import get_folder_metadata as _impl
    return _impl(
        folder_id=folder_id,
        folder_path=folder_path,
        folder_name=folder_name,
        include_children=include_children,
        max_children=max_children,
        include_counts=include_counts,
    )

def create_folder(
    new_folder_name: str,
    parent_folder_id: Optional[str] = None,
    parent_folder_path: Optional[str] = None,
    parent_folder_name: Optional[str] = None,
    item_type: Optional[Any] = None,
    allow_existing: bool = False,
) -> str:
    from outlook_mcp.tools.folders import create_folder as _impl
    return _impl(
        new_folder_name=new_folder_name,
        parent_folder_id=parent_folder_id,
        parent_folder_path=parent_folder_path,
        parent_folder_name=parent_folder_name,
        item_type=item_type,
        allow_existing=allow_existing,
    )

def rename_folder(
    folder_id: Optional[str] = None,
    new_name: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
) -> str:
    from outlook_mcp.tools.folders import rename_folder as _impl
    return _impl(
        folder_id=folder_id,
        new_name=new_name,
        folder_path=folder_path,
        folder_name=folder_name,
    )

def delete_folder(
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
    confirm: bool = False,
) -> str:
    from outlook_mcp.tools.folders import delete_folder as _impl
    return _impl(
        folder_id=folder_id,
        folder_path=folder_path,
        folder_name=folder_name,
        confirm=confirm,
    )


def list_recent_emails(
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
    folder_ids: Optional[Any] = None,
    folder_paths: Optional[Any] = None,
    offset: int = 0,
    unread_only: bool = False,
) -> str:
    from outlook_mcp.tools.email_list import list_recent_emails as _impl
    return _impl(
        days=days,
        folder_name=folder_name,
        max_results=max_results,
        include_preview=include_preview,
        include_all_folders=include_all_folders,
        folder_ids=folder_ids,
        folder_paths=folder_paths,
        offset=offset,
        unread_only=unread_only,
    )

def list_sent_emails(
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    offset: int = 0,
) -> str:
    from outlook_mcp.tools.email_list import list_sent_emails as _impl
    return _impl(days=days, folder_name=folder_name, max_results=max_results, include_preview=include_preview, offset=offset)


def search_emails(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
    folder_ids: Optional[Any] = None,
    folder_paths: Optional[Any] = None,
    offset: int = 0,
    unread_only: bool = False,
) -> str:
    from outlook_mcp.tools.email_list import search_emails as _impl
    return _impl(
        search_term=search_term,
        days=days,
        folder_name=folder_name,
        max_results=max_results,
        include_preview=include_preview,
        include_all_folders=include_all_folders,
        folder_ids=folder_ids,
        folder_paths=folder_paths,
        offset=offset,
        unread_only=unread_only,
    )

def list_pending_replies(
    days: int = 14,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
    include_unread_only: bool = False,
    conversation_lookback_days: Optional[int] = None,
) -> str:
    from outlook_mcp.tools.email_list import list_pending_replies as _impl
    return _impl(days=days, folder_name=folder_name, max_results=max_results, include_preview=include_preview, include_all_folders=include_all_folders, include_unread_only=include_unread_only, conversation_lookback_days=conversation_lookback_days)
def list_upcoming_events(
    days: int = 7,
    calendar_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_description: bool = False,
    include_all_calendars: bool = False,
) -> str:
    from outlook_mcp.tools.calendar_read import list_upcoming_events as _impl
    return _impl(days=days, calendar_name=calendar_name, max_results=max_results, include_description=include_description, include_all_calendars=include_all_calendars)
def search_calendar_events(
    search_term: str,
    days: int = 30,
    calendar_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_description: bool = False,
    include_all_calendars: bool = False,
) -> str:
    from outlook_mcp.tools.calendar_read import search_calendar_events as _impl
    return _impl(search_term=search_term, days=days, calendar_name=calendar_name, max_results=max_results, include_description=include_description, include_all_calendars=include_all_calendars)
def get_event_by_number(event_number: int) -> str:
    from outlook_mcp.tools.calendar_read import get_event_by_number as _impl
    return _impl(event_number)



def create_calendar_event(
    subject: str,
    start_time: str,
    duration_minutes: Optional[int] = 60,
    location: Optional[str] = None,
    body: Optional[str] = None,
    attendees: Optional[Any] = None,
    reminder_minutes: Optional[int] = 15,
    calendar_name: Optional[str] = None,
    all_day: bool = False,
    send_invitations: bool = True,
) -> str:
    from outlook_mcp.tools.calendar_write import create_calendar_event as _impl
    return _impl(
        subject=subject,
        start_time=start_time,
        duration_minutes=duration_minutes,
        location=location,
        body=body,
        attendees=attendees,
        reminder_minutes=reminder_minutes,
        calendar_name=calendar_name,
        all_day=all_day,
        send_invitations=send_invitations,
    )
def ensure_domain_folder(
    email_number: Optional[int] = None,
    sender_email: Optional[str] = None,
    root_folder_name: Optional[str] = None,
    subfolders: Optional[str] = None,
) -> str:
    from outlook_mcp.tools.domain_rules import ensure_domain_folder as _impl
    return _impl(email_number=email_number, sender_email=sender_email, root_folder_name=root_folder_name, subfolders=subfolders)


def move_email_to_domain_folder(
    email_number: int,
    root_folder_name: Optional[str] = None,
    create_if_missing: bool = True,
    subfolders: Optional[str] = None,
) -> str:
    from outlook_mcp.tools.domain_rules import move_email_to_domain_folder as _impl
    return _impl(
        email_number=email_number,
        root_folder_name=root_folder_name,
        create_if_missing=create_if_missing,
        subfolders=subfolders,
    )


def set_email_category(
    email_number: int,
    category: str,
    overwrite: bool = False,
) -> str:
    """Compatibilita retro: delega a apply_category."""
    overwrite_bool = coerce_bool(overwrite)
    append_flag = not overwrite_bool
    return apply_category(
        categories=[category],
        email_number=email_number,
        overwrite=overwrite_bool,
        append=append_flag,
    )

def get_email_context(
    email_number: int,
    include_thread: bool = True,
    thread_limit: int = 5,
    lookback_days: int = 30,
    include_sent: bool = True,
    additional_folders: Optional[List[str]] = None,
) -> str:
    """Compatibilita' retro: delega al servizio email."""
    return _service_get_email_context(
        email_number=email_number,
        include_thread=include_thread,
        thread_limit=thread_limit,
        lookback_days=lookback_days,
        include_sent=include_sent,
        additional_folders=additional_folders,
    )


def move_email_to_folder(
    target_folder_id: Optional[str] = None,
    target_folder_path: Optional[str] = None,
    target_folder_name: Optional[str] = None,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    create_if_missing: bool = False,
) -> str:
    from outlook_mcp.tools.email_actions import move_email_to_folder as _impl
    return _impl(target_folder_id=target_folder_id, target_folder_path=target_folder_path, target_folder_name=target_folder_name, email_number=email_number, message_id=message_id, create_if_missing=create_if_missing)
def mark_email_read_unread(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    unread: Optional[bool] = None,
    flag: Optional[str] = None,
) -> str:
    from outlook_mcp.tools.email_actions import mark_email_read_unread as _impl
    return _impl(email_number=email_number, message_id=message_id, unread=unread, flag=flag)
def apply_category(
    categories: Optional[Any] = None,
    category: Optional[str] = None,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    overwrite: bool = False,
    append: bool = False,
) -> str:
    from outlook_mcp.tools.email_actions import apply_category as _impl
    return _impl(categories=categories, category=category, email_number=email_number, message_id=message_id, overwrite=overwrite, append=append)
def reply_to_email_by_number(
    email_number: Optional[int] = None,
    reply_text: str = "",
    message_id: Optional[str] = None,
    reply_all: bool = False,
    send: bool = True,
    attachments: Optional[Any] = None,
    use_html: bool = False,
) -> str:
    from outlook_mcp.tools.email_actions import reply_to_email_by_number as _impl
    return _impl(email_number=email_number, reply_text=reply_text, message_id=message_id, reply_all=reply_all, send=send, attachments=attachments, use_html=use_html)
def compose_email(
    recipient_email: str,
    subject: str,
    body: str,
    cc_email: Optional[str] = None,
    bcc_email: Optional[str] = None,
    attachments: Optional[Any] = None,
    send: bool = True,
    use_html: bool = False,
) -> str:
    from outlook_mcp.tools.email_actions import compose_email as _impl
    return _impl(recipient_email=recipient_email, subject=subject, body=body, cc_email=cc_email, bcc_email=bcc_email, attachments=attachments, send=send, use_html=use_html)
def get_attachments(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    save_to: Optional[str] = None,
    download: bool = False,
    limit: Optional[int] = None,
) -> str:
    from outlook_mcp.tools.attachments import get_attachments as _impl
    return _impl(email_number=email_number, message_id=message_id, save_to=save_to, download=download, limit=limit)


def search_contacts(
    search_term: Optional[str] = None,
    max_results: int = 50,
) -> str:
    from outlook_mcp.tools.contacts import search_contacts as _impl
    return _impl(search_term=search_term, max_results=max_results)

def attach_to_email(
    attachments: Any,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    send: bool = False,
) -> str:
    from outlook_mcp.tools.attachments import attach_to_email as _impl
    return _impl(attachments=attachments, email_number=email_number, message_id=message_id, send=send)

def batch_manage_emails(
    email_numbers: Optional[Any] = None,
    message_ids: Optional[Any] = None,
    move_to_folder_id: Optional[str] = None,
    move_to_folder_path: Optional[str] = None,
    move_to_folder_name: Optional[str] = None,
    mark_as: Optional[str] = None,
    delete: bool = False,
) -> str:
    from outlook_mcp.tools.email_actions import batch_manage_emails as _impl
    return _impl(email_numbers=email_numbers, message_ids=message_ids, move_to_folder_id=move_to_folder_id, move_to_folder_path=move_to_folder_path, move_to_folder_name=move_to_folder_name, mark_as=mark_as, delete=delete)
def _serialize_tool_metadata(tool: Any) -> Dict[str, Any]:
    """Convert FastMCP tool metadata into plain dicts for JSON responses."""
    return {
        "name": getattr(tool, "name", None),
        "description": getattr(tool, "description", None),
        "input_schema": getattr(tool, "inputSchema", None),
        "output_schema": getattr(tool, "outputSchema", None),
        "annotations": getattr(tool, "annotations", None),
    }


def _serialize_contents(contents: List[Any]) -> List[Dict[str, Any]]:
    """Convert FastMCP content payloads into JSON serializable dictionaries."""
    serialized: List[Dict[str, Any]] = []
    for item in contents:
        if hasattr(item, "__dict__"):
            # Copy to avoid leaking internal references
            serialized.append(dict(item.__dict__))
        else:
            serialized.append({"type": "text", "text": str(item)})
    return serialized


def _verify_outlook_connection() -> None:
    """Validate Outlook availability before starting any server mode."""
    print("Connessione a Outlook...")
    logger.info("Verifica della connessione a Outlook in corso.")
    outlook, namespace = connect_to_outlook()
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
    inbox_items = getattr(getattr(inbox, "Items", None), "Count", "sconosciuto")
    print(f"Connessione a Outlook riuscita. La Posta in arrivo contiene {inbox_items} elementi.")
    logger.info(
        "Connessione a Outlook verificata. La Posta in arrivo contiene %s elementi.",
        inbox_items,
    )
    # Release COM references to avoid locking Outlook instances unnecessarily
    del inbox
    del namespace
    del outlook


# Also register tools when the module is merely imported (defensive)
try:
    _register_tools()
except Exception:  # pragma: no cover - best effort
    logger.debug("Registrazione dei tool MCP non effettuata all'import; verra' ripetuta.", exc_info=True)


def _create_http_app() -> Any:
    """Instantiate the optional FastAPI bridge for HTTP integrations."""
    if (
        FastAPI is None
        or BaseModel is None
        or Field is None
        or HTTPException is None
        or Body is None
        or uvicorn is None
    ):
        raise RuntimeError(
            "La modalita HTTP richiede fastapi, uvicorn e pydantic. "
            "Installa le dipendenze: pip install fastapi uvicorn"
        )

    # Defensive: make sure tools are registered when the HTTP app is created
    try:
        _register_tools()
    except Exception:
        logger.debug("Registrazione tool MCP non eseguita in _create_http_app.", exc_info=True)

    class ToolCallRequest(BaseModel):
        """Pydantic model describing an HTTP tool invocation."""

        arguments: Dict[str, Any] = Field(default_factory=dict)

    app = FastAPI(
        title="Outlook MCP HTTP Bridge",
        description="Bridge HTTP leggero per richiamare gli strumenti MCP di Outlook da automazioni esterne (es. n8n).",
        version="1.0.0",
    )

    async def _handle_tool_invocation(tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """Shared execution helper for HTTP-triggered tool calls."""
        logger.info("Invocazione HTTP del tool %s con argomenti=%s", tool_name, arguments)
        try:
            # Feature-gate check before actual invocation
            group = get_tool_group(tool_name)
            if not is_tool_enabled(tool_name, group):
                raise HTTPException(status_code=403, detail=f"Tool '{tool_name}' disabilitato dal server")  # type: ignore[misc]
            contents, output = await mcp.call_tool(tool_name, arguments)
            return {
                "tool": tool_name,
                "content": _serialize_contents(contents),
                "result": output,
            }
        except ToolError as exc:
            logger.warning("Tool %s non trovato per invocazione HTTP: %s", tool_name, exc)
            raise HTTPException(status_code=404, detail=str(exc))  # type: ignore[misc]
        except Exception as exc:  # pylint: disable=broad-except
            logger.exception("Errore durante l'esecuzione HTTP del tool %s.", tool_name)
            raise HTTPException(status_code=500, detail=str(exc))  # type: ignore[misc]

    def _resolve_root_payload(payload: Dict[str, Any]) -> Dict[str, Any]:
        """Normalize POST / payloads into {'tool': str, 'arguments': dict}."""
        preferred_keys = ("tool", "tool_name", "toolName", "name")
        tool_name: Optional[str] = None
        arguments: Any = payload.get("arguments", {})

        for key in preferred_keys:
            value = payload.get(key)
            if isinstance(value, str) and value.strip():
                tool_name = value.strip()
                break

        if tool_name is None:
            candidate_pairs = [
                (key, value)
                for key, value in payload.items()
                if key not in {"arguments"}
                and isinstance(key, str)
                and isinstance(value, dict)
            ]
            if len(candidate_pairs) == 1:
                tool_name, arguments = candidate_pairs[0]

        if tool_name is None:
            raise HTTPException(
                status_code=400,
                detail=(
                    "Specifica il tool da eseguire usando il campo 'tool' (stringa) "
                    "oppure struttura il payload come {\"nome_tool\": {...}}."
                ),
            )

        if not isinstance(arguments, dict):
            raise HTTPException(
                status_code=400,
                detail="Il campo 'arguments' deve essere un oggetto JSON (dizionario).",
            )

        return {"tool": tool_name, "arguments": arguments}

    @app.get("/")
    async def root() -> Dict[str, Any]:
        """Provide a quick-start payload when browsing the root endpoint."""
        # Ensure tools are registered before listing
        try:
            _register_tools()
        except Exception:
            logger.debug("Registrazione tool MCP non eseguita in GET /.", exc_info=True)
        tools = await mcp.list_tools()
        visible = [t for t in tools if is_tool_enabled(t.name, get_tool_group(t.name))]
        return {
            "message": "Outlook MCP HTTP Bridge attivo. Usa POST /tools/{tool_name} oppure POST / con {\"tool\": \"nome\", \"arguments\": {...}}.",
            "available_tools": [tool.name for tool in visible],
        }

    @app.post("/")
    async def invoke_tool_root(payload: Dict[str, Any] = Body(...)) -> Dict[str, Any]:
        """Allow POST / to execute a tool with flexible payload aliases."""
        normalized = _resolve_root_payload(payload)
        return await _handle_tool_invocation(
            normalized["tool"], normalized.get("arguments", {})
        )

    @app.get("/health")
    async def health_check() -> Dict[str, str]:
        """Simple readiness probe for container orchestrators."""
        return {"status": "ok"}

    @app.get("/tools")
    async def list_tools() -> Dict[str, Any]:
        """Return metadata for the registered MCP tools."""
        try:
            _register_tools()
        except Exception:
            logger.debug("Registrazione tool MCP non eseguita in GET /tools.", exc_info=True)
        tools = await mcp.list_tools()
        visible = [t for t in tools if is_tool_enabled(t.name, get_tool_group(t.name))]
        return {"tools": [_serialize_tool_metadata(tool) for tool in visible]}

    @app.post("/tools/{tool_name}")
    async def invoke_tool(tool_name: str, request: ToolCallRequest) -> Dict[str, Any]:
        """Execute an MCP tool and return the serialized content/result."""
        arguments = request.arguments or {}
        return await _handle_tool_invocation(tool_name, arguments)

    return app


def _start_http_bridge(host: str, port: int) -> None:
    """Run the FastAPI HTTP server for MCP bridge mode."""
    app = _create_http_app()
    logger.info("Avvio del bridge HTTP su http://%s:%s", host, port)
    uvicorn.run(app, host=host, port=port, log_level="info")  # type: ignore[arg-type]


def _build_arg_parser() -> argparse.ArgumentParser:
    """Create CLI parser supporting both MCP and HTTP bridge modes."""
    parser = argparse.ArgumentParser(
        description="Outlook MCP Server - accesso diretto o bridge HTTP."
    )
    parser.add_argument(
        "--mode",
        choices=("mcp", "http"),
        default="mcp",
        help="Modalita di esecuzione: 'mcp' (default) per FastMCP oppure 'http' per il bridge REST.",
    )
    parser.add_argument(
        "--host",
        default="0.0.0.0",
        help="Indirizzo di bind per i server di rete (streamable-http, sse o bridge REST).",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Porta di ascolto per i server di rete (streamable-http, sse o bridge REST).",
    )
    parser.add_argument(
        "--transport",
        choices=("stdio", "sse", "streamable-http"),
        default="stdio",
        help="Trasporto FastMCP quando --mode=mcp. Usa 'streamable-http' per n8n o altri client MCP HTTP.",
    )
    parser.add_argument(
        "--stream-path",
        default="/mcp",
        help="Percorso base per il trasporto streamable-http (solo quando --transport=streamable-http).",
    )
    parser.add_argument(
        "--mount-path",
        default="/",
        help="Percorso di montaggio SSE (solo quando --transport=sse).",
    )
    parser.add_argument(
        "--sse-path",
        default="/sse",
        help="Endpoint SSE relativo (solo quando --transport=sse).",
    )
    parser.add_argument(
        "--skip-outlook-check",
        action="store_true",
        help="Salta il controllo iniziale della connessione a Outlook (sconsigliato).",
    )
    return parser


def main() -> None:
    """Entrypoint principale per il server MCP/bridge HTTP."""
    parser = _build_arg_parser()
    args = parser.parse_args()

    print("Avvio di Outlook MCP Server...")
    logger.info("Outlook MCP Server avviato in modalita %s.", args.mode)

    # Ensure MCP tools are registered before serving/listing
    _register_tools()

    if not args.skip_outlook_check:
        _verify_outlook_connection()
    else:
        logger.warning("Controllo iniziale di Outlook disabilitato per richiesta esplicita.")

    if args.mode == "mcp":
        transport = args.transport
        if transport in {"sse", "streamable-http"}:
            mcp.settings.host = args.host
            mcp.settings.port = args.port
        if transport == "sse":
            mcp.settings.mount_path = args.mount_path
            mcp.settings.sse_path = args.sse_path
            print(
                f"Avvio del server MCP (SSE) su http://{args.host}:{args.port}{args.sse_path} "
                "(Ctrl+C per interrompere)."
            )
            logger.info(
                "Server MCP avviato in modalita SSE su http://%s:%s%s (mount=%s).",
                args.host,
                args.port,
                args.sse_path,
                args.mount_path,
            )
            mcp.run(transport="sse", mount_path=args.mount_path)
        elif transport == "streamable-http":
            mcp.settings.streamable_http_path = args.stream_path
            print(
                f"Avvio del server MCP (streamable-http) su "
                f"http://{args.host}:{args.port}{args.stream_path} (Ctrl+C per interrompere)."
            )
            logger.info(
                "Server MCP avviato in modalita streamable-http su http://%s:%s%s.",
                args.host,
                args.port,
                args.stream_path,
            )
            mcp.run(transport="streamable-http")
        else:
            print("Avvio del server MCP (stdio). Premi Ctrl+C per interrompere.")
            logger.info("Server MCP avviato su stdio.")
            mcp.run()
    else:
        print(f"Avvio del bridge HTTP su http://{args.host}:{args.port} (Ctrl+C per interrompere).")
        _start_http_bridge(args.host, args.port)


# Run the server
if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # pylint: disable=broad-except
        print(f"Errore durante l'avvio del server: {str(exc)}")
        logger.exception("Errore durante l'avvio di Outlook MCP Server.")