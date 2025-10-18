"""Feature flag support for enabling/disabling MCP tools.

This module provides a lightweight gating system so administrators can
enable/disable groups of tools or individual tools without changing code.

Precedence: explicit disables override enables.

Configuration sources (highest to lowest priority):
- Environment variables
  - OUTLOOK_MCP_FEATURES_FILE: path to a JSON file
  - OUTLOOK_MCP_ENABLED_GROUPS: comma/semicolon-separated list
  - OUTLOOK_MCP_DISABLED_GROUPS: comma/semicolon-separated list
  - OUTLOOK_MCP_ENABLED_TOOLS: comma/semicolon-separated list
  - OUTLOOK_MCP_DISABLED_TOOLS: comma/semicolon-separated list
- features.json in the project root (next to README.md)

If no config is found, all tools are enabled by default.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from functools import wraps
import inspect
from typing import Dict, Optional, Set, Any

from .logger import logger


# Runtime registry: tool_name -> group
_TOOL_GROUPS: Dict[str, str] = {}


def _project_root() -> Path:
    """Return repository root (one level above the package dir)."""
    return Path(__file__).resolve().parent.parent


def _read_json_file(path: Path) -> Optional[dict]:
    try:
        if path.is_file():
            with path.open("r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as exc:  # pragma: no cover - defensive
        logger.warning("Impossibile leggere la configurazione feature da %s: %s", path, exc)
    return None


def _split_values(value: Optional[str]) -> Set[str]:
    if not value:
        return set()
    parts = [seg.strip() for seg in value.replace(";", ",").split(",")]
    return {seg for seg in parts if seg}


class _FeatureState:
    def __init__(self) -> None:
        self.enabled_groups: Set[str] = set()
        self.disabled_groups: Set[str] = set()
        self.enabled_tools: Set[str] = set()
        self.disabled_tools: Set[str] = set()

    def __repr__(self) -> str:  # pragma: no cover - debug helper
        return (
            f"FeatureState(enabled_groups={self.enabled_groups}, "
            f"disabled_groups={self.disabled_groups}, "
            f"enabled_tools={self.enabled_tools}, "
            f"disabled_tools={self.disabled_tools})"
        )


_FEATURES = _FeatureState()


def _load_from_file() -> None:
    cfg_path_env = os.environ.get("OUTLOOK_MCP_FEATURES_FILE")
    path = Path(cfg_path_env) if cfg_path_env else _project_root() / "features.json"
    data = _read_json_file(path) or {}
    eg = data.get("enabled_groups", []) or []
    dg = data.get("disabled_groups", []) or []
    et = data.get("enabled_tools", []) or []
    dt = data.get("disabled_tools", []) or []
    _FEATURES.enabled_groups.update(str(x) for x in eg)
    _FEATURES.disabled_groups.update(str(x) for x in dg)
    _FEATURES.enabled_tools.update(str(x) for x in et)
    _FEATURES.disabled_tools.update(str(x) for x in dt)
    if data:
        logger.info(
            "Configurazione features caricata da %s (groups on=%s off=%s, tools on=%s off=%s)",
            path,
            sorted(_FEATURES.enabled_groups) or "*",
            sorted(_FEATURES.disabled_groups) or "-",
            sorted(_FEATURES.enabled_tools) or "-",
            sorted(_FEATURES.disabled_tools) or "-",
        )


def _load_from_env() -> None:
    _FEATURES.enabled_groups.update(_split_values(os.environ.get("OUTLOOK_MCP_ENABLED_GROUPS")))
    _FEATURES.disabled_groups.update(_split_values(os.environ.get("OUTLOOK_MCP_DISABLED_GROUPS")))
    _FEATURES.enabled_tools.update(_split_values(os.environ.get("OUTLOOK_MCP_ENABLED_TOOLS")))
    _FEATURES.disabled_tools.update(_split_values(os.environ.get("OUTLOOK_MCP_DISABLED_TOOLS")))


def _normalize_group(group: Optional[str]) -> Optional[str]:
    return group.lower() if isinstance(group, str) else None


def reload_features() -> None:
    """(Re)load configuration from disk and environment."""
    _FEATURES.enabled_groups.clear()
    _FEATURES.disabled_groups.clear()
    _FEATURES.enabled_tools.clear()
    _FEATURES.disabled_tools.clear()
    _load_from_file()
    _load_from_env()


def is_tool_enabled(tool_name: str, group: Optional[str] = None) -> bool:
    """Return True if a tool is enabled according to current configuration."""
    g = _normalize_group(group)

    # Explicit tool disable overrides everything
    if tool_name in _FEATURES.disabled_tools:
        return False

    # Group-level disable
    if g and g in (s.lower() for s in _FEATURES.disabled_groups):
        return False

    # If any allow-lists are present, require membership
    any_enables = bool(_FEATURES.enabled_groups or _FEATURES.enabled_tools)
    if any_enables:
        if tool_name in _FEATURES.enabled_tools:
            return True
        if g and g in (s.lower() for s in _FEATURES.enabled_groups):
            return True
        # Not explicitly enabled
        return False

    # Default: enabled
    return True


def feature_gate(group: Optional[str] = None, name: Optional[str] = None):
    """Decorator to gate a tool by group and record its mapping.

    Use alongside @mcp.tool():

        @mcp.tool()
        @feature_gate(group="email.list")
        def list_recent_emails(...):
            ...
    """

    def decorator(func):
        tool_name = name or getattr(func, "__name__", "tool")
        # Record mapping for discovery and filtering
        if group:
            _TOOL_GROUPS[tool_name] = str(group)

        @wraps(func)
        def wrapper(*args, **kwargs):
            grp = _TOOL_GROUPS.get(tool_name)
            if not is_tool_enabled(tool_name, grp):
                from mcp.server.fastmcp.exceptions import ToolError  # local import

                raise ToolError(
                    f"Lo strumento '{tool_name}' e' disabilitato. "
                    f"Aggiorna features.json o le variabili ambiente per abilitarlo."
                )
            return func(*args, **kwargs)

        # Preserve attributes/signature so FastMCP can derive JSON schema
        wrapper.__name__ = func.__name__
        wrapper.__doc__ = func.__doc__
        wrapper.__dict__.update(func.__dict__)
        try:
            wrapper.__signature__ = inspect.signature(func)
        except (ValueError, TypeError):
            pass
        return wrapper

    return decorator


def get_tool_group(tool_name: str) -> Optional[str]:
    return _TOOL_GROUPS.get(tool_name)


# Load configuration on import
reload_features()

__all__ = [
    "feature_gate",
    "is_tool_enabled",
    "reload_features",
    "get_tool_group",
]


