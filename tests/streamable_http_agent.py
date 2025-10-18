"""
Utility script for exercising the Outlook MCP server through the Streamable HTTP
transport (the same transport used by n8n and other HTTP-first MCP clients).

The agent performs the MCP handshake, lists the available tools and executes a
small, configurable battery of tool invocations so you can validate the full
round-trip without needing a dedicated UI.

Usage (assuming the Outlook MCP server is already running in streamable-http mode):

    python outlook_mcp_server.py --mode mcp --transport streamable-http --port 8000 --stream-path /mcp
    python tests/streamable_http_agent.py --url http://127.0.0.1:8000/mcp --demo

You can also target a different host/path and add extra tool calls:

    python tests/streamable_http_agent.py \\
        --url http://192.168.1.10:9000/custom \\
        --call \"list_recent_emails:max_results=5,include_preview=False\" \\
        --call \"list_upcoming_events:days=14,include_description=True\"
"""

from __future__ import annotations

import argparse
import ast
import json
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Sequence

import anyio

from mcp.client.session import ClientSession
from mcp.client.streamable_http import streamablehttp_client
from mcp.shared.session import RequestResponder
from mcp.shared.message import SessionMessage
from mcp.types import CallToolResult, ErrorData, ServerNotification, ServerRequest, TextContent


# ---------------------------------------------------------------------------
# Dataclasses & helpers
# ---------------------------------------------------------------------------


@dataclass(slots=True)
class ToolCall:
    """Represents a queued tool invocation."""

    name: str
    arguments: Dict[str, Any]
    label: str


def _parse_value(raw: str) -> Any:
    """Best-effort conversion of CLI argument values."""
    try:
        return ast.literal_eval(raw)
    except (ValueError, SyntaxError):
        lowered = raw.lower()
        if lowered in {"true", "false"}:
            return lowered == "true"
        return raw


def parse_call_spec(spec: str) -> ToolCall:
    """
    Parse a CLI `--call` specification (e.g. `tool:param=value,param2=42`).
    """
    tool_name, *rest = spec.split(":", 1)
    if not tool_name:
        raise ValueError("Tool name cannot be empty in --call specification.")

    args: Dict[str, Any] = {}
    if rest:
        for token in rest[0].split(","):
            if not token:
                continue
            if "=" not in token:
                raise ValueError(f"Invalid argument token '{token}' (expected key=value).")
            key, value_raw = token.split("=", 1)
            args[key] = _parse_value(value_raw)
    label = f"{tool_name}({json.dumps(args, ensure_ascii=False)})" if args else tool_name
    return ToolCall(name=tool_name, arguments=args, label=label)


def _format_call_result(result: CallToolResult) -> str:
    """Convert a CallToolResult into a printable string."""
    if result.structuredContent is not None:
        return json.dumps(result.structuredContent, indent=2, ensure_ascii=False)

    if not result.content:
        return "(nessun contenuto)"

    fragments: List[str] = []
    for chunk in result.content:
        if isinstance(chunk, TextContent):
            fragments.append(chunk.text)
        else:
            fragments.append(f"[{chunk.type}] {chunk.model_dump(mode='json')}")
    return "\n".join(fragments)


async def _default_message_handler(
    message: RequestResponder[ServerRequest, Any] | ServerNotification | Exception,
) -> None:
    """Minimal handler for unexpected server-initiated traffic."""
    if isinstance(message, Exception):
        print("[notifica] eccezione ricevuta dal server:", message)
        return

    if isinstance(message, RequestResponder):
        # The Outlook MCP server currently does not issue server->client requests.
        root_type = type(message.request.root).__name__
        print(f"[notifica] richiesta non gestita dal server: {root_type}")
        # Make sure we respond with an error to keep the session healthy.
        with message:
            await message.respond(ErrorData(code=0, message="Unsupported request", data=None))
        return

    root_type = type(message.root).__name__
    print(f"[notifica] evento dal server: {root_type}")


def _build_demo_calls() -> Sequence[ToolCall]:
    """Safe default walkthrough that exercises a handful of tools."""
    return [
        ToolCall(
            name="params",
            arguments={},
            label="params()",
        ),
        ToolCall(
            name="get_current_datetime",
            arguments={"include_utc": True},
            label="get_current_datetime(include_utc=True)",
        ),
        ToolCall(
            name="list_folders",
            arguments={"max_depth": 1, "include_counts": False, "include_paths": True, "include_ids": False},
            label="list_folders(max_depth=1)",
        ),
        ToolCall(
            name="list_recent_emails",
            arguments={"days": 3, "max_results": 3, "include_preview": False},
            label="list_recent_emails(days=3, max_results=3)",
        ),
        ToolCall(
            name="list_sent_emails",
            arguments={"days": 7, "max_results": 3, "include_preview": False},
            label="list_sent_emails(days=7, max_results=3)",
        ),
        ToolCall(
            name="list_pending_replies",
            arguments={"days": 14, "max_results": 3, "include_preview": False},
            label="list_pending_replies(days=14, max_results=3)",
        ),
        ToolCall(
            name="list_upcoming_events",
            arguments={"days": 14, "max_results": 5, "include_description": False},
            label="list_upcoming_events(days=14, max_results=5)",
        ),
        ToolCall(
            name="search_calendar_events",
            arguments={"search_term": "riunione", "days": 30, "max_results": 3},
            label="search_calendar_events(search_term='meeting')",
        ),
    ]


# ---------------------------------------------------------------------------
# Core runner
# ---------------------------------------------------------------------------


async def run_agent(url: str, calls: Sequence[ToolCall], headers: Dict[str, str] | None) -> None:
    """Main async entrypoint executing the requested tool calls."""
    async with streamablehttp_client(url, headers=headers) as (read_stream, write_stream, get_session_id):
        async with ClientSession(
            read_stream=read_stream,
            write_stream=write_stream,
            message_handler=_default_message_handler,
        ) as session:
            init_result = await session.initialize()
            print(f"[ok] Sessione inizializzata con protocolVersion={init_result.protocolVersion}")
            session_id = get_session_id()
            if session_id:
                print(f"[info] Session ID: {session_id}")

            tool_list = await session.list_tools()
            print(f"[ok] Strumenti disponibili ({len(tool_list.tools)}):")
            for tool in tool_list.tools:
                print(f"  - {tool.name}: {tool.description or '(nessuna descrizione)'}")

            if not calls:
                return

            print("\n=== Avvio chiamate tool ===\n")

            for call in calls:
                print(f"[call] {call.label}")
                try:
                    result = await session.call_tool(call.name, call.arguments)
                except Exception as exc:
                    print(f"[errore] Invocazione '{call.name}' fallita: {exc}")
                    continue

                if result.isError:
                    content_str = _format_call_result(result)
                    print(f"[errore] Il tool ha restituito un errore:\n{content_str}\n")
                else:
                    content_str = _format_call_result(result)
                    print(f"[risposta]\n{content_str}\n")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_cli() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Client di test per il trasporto MCP Streamable HTTP.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--url",
        default="http://127.0.0.1:8000/mcp",
        help="Endpoint streamable-http esposto dal server MCP.",
    )
    parser.add_argument(
        "--demo",
        action="store_true",
        help="Esegue una batteria di chiamate standard ai tool principali.",
    )
    parser.add_argument(
        "--call",
        action="append",
        default=[],
        metavar="SPEC",
        help="Specifica una chiamata addizionale (es. tool:param=value,param2=42). "
        "Può essere ripetuto più volte.",
    )
    parser.add_argument(
        "--header",
        action="append",
        default=[],
        metavar="NAME=VALUE",
        help="Header HTTP addizionali da inviare (es. Authorization=BearerToken).",
    )
    return parser


def parse_headers(raw_headers: Iterable[str]) -> Dict[str, str]:
    headers: Dict[str, str] = {}
    for item in raw_headers:
        if "=" not in item:
            raise ValueError(f"Intestazione non valida '{item}' (usa NAME=VALUE).")
        key, value = item.split("=", 1)
        headers[key.strip()] = value.strip()
    return headers


def gather_calls(args: argparse.Namespace) -> Sequence[ToolCall]:
    calls: List[ToolCall] = []

    if args.demo:
        calls.extend(_build_demo_calls())

    for spec in args.call:
        calls.append(parse_call_spec(spec))

    # Remove duplicates while preserving order (later specs override earlier ones)
    unique: Dict[str, ToolCall] = {}
    for call in calls:
        unique[call.label] = call
    return list(unique.values())


def main() -> None:
    parser = build_cli()
    args = parser.parse_args()

    try:
        headers = parse_headers(args.header)
        calls = gather_calls(args)
        anyio.run(run_agent, args.url, calls, headers)
    except KeyboardInterrupt:
        print("Interrotto dall'utente.")
    except Exception as exc:  # pylint: disable=broad-except
        print(f"Errore durante l'esecuzione del client: {exc}")


if __name__ == "__main__":
    main()
