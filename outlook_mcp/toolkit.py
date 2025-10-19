"""Registration helpers for Outlook MCP tools."""

from __future__ import annotations

from typing import Callable, Iterable, List, Optional

from mcp.server.fastmcp import FastMCP

ToolBinder = Callable[[FastMCP], None]

_tool_binders: List[ToolBinder] = []
_registered = False
_current_mcp: Optional[FastMCP] = None


def mcp_tool(*decorator_args, **decorator_kwargs):
    """Deferred variant of FastMCP.tool that registers once the MCP instance is ready."""

    def decorator(func):
        def binder(mcp: FastMCP) -> None:
            mcp.tool(*decorator_args, **decorator_kwargs)(func)

        _tool_binders.append(binder)
        return func

    return decorator


def register_all_tools(mcp: FastMCP, *, force: bool = False) -> None:
    """Register all deferred tool definitions against the supplied FastMCP instance."""
    global _registered, _current_mcp
    if _registered and not force:
        return
    _current_mcp = mcp
    for binder in _tool_binders:
        binder(mcp)
    _registered = True


def iter_registered_tool_binders() -> Iterable[ToolBinder]:
    """Expose raw binders for diagnostics/testing."""
    return tuple(_tool_binders)


def reset_tool_registry() -> None:
    """Allow forcing a re-registration (used by configuration reloads)."""
    global _registered, _current_mcp
    _registered = False
    _current_mcp = None


def get_current_mcp() -> FastMCP:
    """Return the active FastMCP instance (after registration)."""
    if _current_mcp is None:
        raise RuntimeError("FastMCP instance not yet registered. Call register_all_tools first.")
    return _current_mcp


__all__ = ["mcp_tool", "register_all_tools", "iter_registered_tool_binders", "reset_tool_registry", "get_current_mcp"]
