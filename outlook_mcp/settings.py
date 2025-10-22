"""Runtime configuration helpers for the Outlook MCP server."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Iterable, Set

from outlook_mcp.logger import logger

_CONFIG_CACHE: dict | None = None
_PROMOTIONAL_CACHE: Set[str] | None = None

DEFAULT_PROMOTIONAL_KEYWORDS: Iterable[str] = (
    "newsletter",
    "iscriviti",
    "webinar",
    "evento online",
    "evento digitale",
    "promo ",
    "promozione",
    "the future of",
    "unsubscribe",
    "webversion",
    "marketing",
)

CONFIG_FILE = Path(__file__).resolve().parent.parent / "config.json"


def _load_raw_config() -> dict:
    """Read the JSON config from disk, returning an empty dict on failure."""
    global _CONFIG_CACHE

    if _CONFIG_CACHE is not None:
        return _CONFIG_CACHE

    try:
        with CONFIG_FILE.open("r", encoding="utf-8") as handle:
            _CONFIG_CACHE = json.load(handle)
    except FileNotFoundError:
        logger.warning("File di configurazione %s non trovato. Uso valori predefiniti.", CONFIG_FILE)
        _CONFIG_CACHE = {}
    except Exception:
        logger.exception("Errore durante il caricamento della configurazione da %s.", CONFIG_FILE)
        _CONFIG_CACHE = {}
    return _CONFIG_CACHE


def get_promotional_keywords() -> Set[str]:
    """Return a normalized set of promotional keywords used to filter marketing emails."""
    global _PROMOTIONAL_CACHE
    if _PROMOTIONAL_CACHE is not None:
        return _PROMOTIONAL_CACHE

    config = _load_raw_config()
    filters_section = config.get("filters", {})
    raw_keywords = []

    if isinstance(filters_section, dict):
        candidate = filters_section.get("promotional_keywords")
        if isinstance(candidate, list):
            raw_keywords = candidate
    if not raw_keywords:
        candidate = config.get("promotional_keywords")
        if isinstance(candidate, list):
            raw_keywords = candidate

    normalized = {
        str(keyword).strip().lower()
        for keyword in raw_keywords
        if str(keyword).strip()
    }
    if not normalized:
        normalized = {str(keyword).strip().lower() for keyword in DEFAULT_PROMOTIONAL_KEYWORDS if str(keyword).strip()}

    _PROMOTIONAL_CACHE = normalized
    return _PROMOTIONAL_CACHE


def reload_settings() -> None:
    """Clear cached configuration so the next access reloads data from disk."""
    global _CONFIG_CACHE, _PROMOTIONAL_CACHE
    _CONFIG_CACHE = None
    _PROMOTIONAL_CACHE = None


__all__ = ["get_promotional_keywords", "reload_settings", "DEFAULT_PROMOTIONAL_KEYWORDS"]
