"""Common helper utilities shared across Outlook MCP service modules."""

from __future__ import annotations

import datetime
from typing import Any, Optional

__all__ = [
    "parse_datetime_string",
    "describe_importance",
    "describe_sensitivity",
    "describe_flag_status",
    "parse_importance",
    "parse_sensitivity",
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
    tz_from_suffix: Optional[datetime.tzinfo] = None
    if text.endswith("Z"):
        tz_from_suffix = datetime.timezone.utc
        text = text[:-1]

    try:
        parsed = datetime.datetime.fromisoformat(text)
        if parsed.tzinfo:
            return parsed
        if tz_from_suffix:
            return parsed.replace(tzinfo=tz_from_suffix)
        return parsed
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
            parsed = datetime.datetime.strptime(text, fmt)
            if tz_from_suffix:
                return parsed.replace(tzinfo=tz_from_suffix)
            return parsed
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


def describe_sensitivity(value: Any) -> str:
    """Map Outlook sensitivity values to readable Italian labels."""
    sensitivity_map = {
        0: "Normale",
        1: "Personale",
        2: "Privato",
        3: "Confidenziale",
    }
    if isinstance(value, int) and value in sensitivity_map:
        return sensitivity_map[value]
    return str(value) if value is not None else "Sconosciuta"


def describe_flag_status(value: Any) -> str:
    """Map Outlook flag status values to readable Italian labels."""
    flag_status_map = {
        0: "Nessuno",
        1: "Completato",
        2: "Contrassegnato",
    }
    if isinstance(value, int) and value in flag_status_map:
        return flag_status_map[value]
    return str(value) if value is not None else "Sconosciuto"


def parse_importance(importance_input: Optional[str]) -> Optional[int]:
    """Parse an importance string to Outlook importance code."""
    if importance_input is None:
        return None
    normalized = importance_input.strip().lower()
    importance_reverse_map = {
        "bassa": 0,
        "low": 0,
        "normale": 1,
        "normal": 1,
        "alta": 2,
        "high": 2,
    }
    return importance_reverse_map.get(normalized)


def parse_sensitivity(sensitivity_input: Optional[str]) -> Optional[int]:
    """Parse a sensitivity string to Outlook sensitivity code."""
    if sensitivity_input is None:
        return None
    normalized = sensitivity_input.strip().lower()
    sensitivity_reverse_map = {
        "normale": 0,
        "normal": 0,
        "personale": 1,
        "personal": 1,
        "privato": 2,
        "private": 2,
        "confidenziale": 3,
        "confidential": 3,
    }
    return sensitivity_reverse_map.get(normalized)
