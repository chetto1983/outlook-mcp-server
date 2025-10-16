"""In-memory caches shared across Outlook MCP server components."""

from typing import Any, Dict

from .logger import logger

email_cache: Dict[int, Dict[str, Any]] = {}
calendar_cache: Dict[int, Dict[str, Any]] = {}


def clear_email_cache() -> None:
    """Clear the email cache."""
    email_cache.clear()
    logger.debug("Cache dei messaggi svuotata.")


def clear_calendar_cache() -> None:
    """Clear the calendar event cache."""
    calendar_cache.clear()
    logger.debug("Cache degli eventi svuotata.")


__all__ = ["email_cache", "calendar_cache", "clear_email_cache", "clear_calendar_cache"]
