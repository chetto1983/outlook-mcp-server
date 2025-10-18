"""Common helper utilities shared across Outlook MCP service modules."""

from __future__ import annotations

import datetime
from typing import Any, Optional

__all__ = [
    "parse_datetime_string",
    "describe_importance",
    "format_yes_no",
    "format_read_status",
]


def parse_datetime_string(value: Optional[str]) -> Optional[datetime.datetime]:
    """Convert various Outlook string formats into ``datetime`` objects."""
    if not value:
        return None
    text = value.strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1]

    try:
        return datetime.datetime.fromisoformat(text)
    except ValueError:
        pass

    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d",
    ):
        try:
            return datetime.datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def describe_importance(value: Any) -> str:
    """Map Outlook importance values to readable Italian labels."""
    importance_map = {0: "Bassa", 1: "Normale", 2: "Alta"}
    if isinstance(value, int) and value in importance_map:
        return importance_map[value]
    return str(value) if value is not None else "Sconosciuta"


def format_yes_no(value: Any) -> str:
    """Return ``Si`` or ``No`` depending on truthiness."""
    return "Si" if bool(value) else "No"


def format_read_status(unread: bool) -> str:
    """Return localized read status labels."""
    return "Non letta" if bool(unread) else "Letta"

