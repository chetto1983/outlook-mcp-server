import argparse
import datetime
import os
import signal
from typing import Any, Dict, List, Optional, Set, Tuple

from outlook_mcp.mcp_runtime import mcp
from outlook_mcp.toolkit import register_all_tools
from mcp.server.fastmcp.exceptions import ToolError

from outlook_mcp import clear_calendar_cache, clear_email_cache, connect_to_outlook, logger
from outlook_mcp import folders as folder_service
from outlook_mcp.features import is_tool_enabled, get_tool_group

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


# Backwards-compatible aliases for shared helpers
get_folder_by_name = folder_service.get_folder_by_name
get_folder_by_path = folder_service.get_folder_by_path
resolve_folder = folder_service.resolve_folder


def _register_tools() -> None:
    """Import tool modules so their @mcp.tool() decorators register tools.

    Note: tool modules register lazily via outlook_mcp.toolkit, so this simply ensures their definitions are loaded.
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
        register_all_tools(mcp)
        try:
            count = len(mcp._tool_manager.list_tools())  # type: ignore[attr-defined]
            logger.info("Registrazione tool MCP completata: %s tool", count)
        except Exception:
            pass
    except Exception:
        logger.exception("Registrazione dei tool MCP fallita.")

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


def _configure_signal_handlers() -> None:
    """Register signal handlers to perform a graceful shutdown."""

    def _handle_shutdown(signum, frame) -> None:  # type: ignore[unused-argument]
        try:
            signal_name = signal.Signals(signum).name  # type: ignore[attr-defined]
        except Exception:
            signal_name = str(signum)
        logger.info("Segnale %s ricevuto; arresto graduale in corso.", signal_name)
        try:
            clear_email_cache()
            clear_calendar_cache()
        except Exception:
            logger.debug("Errore durante la pulizia delle cache in fase di arresto.", exc_info=True)
        raise SystemExit(0)

    for candidate in (getattr(signal, "SIGINT", None), getattr(signal, "SIGTERM", None)):
        if candidate is None:
            continue
        try:
            signal.signal(candidate, _handle_shutdown)
        except (ValueError, AttributeError, OSError):
            continue


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

    _configure_signal_handlers()

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
