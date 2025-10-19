"""In-memory caches shared across Outlook MCP server components."""

from __future__ import annotations

import time
from collections import OrderedDict
from collections.abc import MutableMapping
from typing import Any, Iterator, Optional, Tuple

from .logger import logger


class TimedLRUCache(MutableMapping[int, Any]):
    """LRU cache with optional TTL eviction."""

    def __init__(self, *, max_entries: int, ttl_seconds: Optional[float]) -> None:
        self.max_entries = max_entries
        self.ttl_seconds = ttl_seconds
        self._store: "OrderedDict[int, Any]" = OrderedDict()
        self._timestamps: dict[int, float] = {}

    def _now(self) -> float:
        return time.monotonic()

    def _is_expired(self, key: int) -> bool:
        if self.ttl_seconds is None:
            return False
        ts = self._timestamps.get(key)
        if ts is None:
            return False
        return (self._now() - ts) > self.ttl_seconds

    def _evict_key(self, key: int) -> None:
        self._store.pop(key, None)
        self._timestamps.pop(key, None)

    def _purge_expired(self) -> None:
        if self.ttl_seconds is None:
            return
        cutoff = self._now() - self.ttl_seconds
        expired = [key for key, ts in self._timestamps.items() if ts < cutoff]
        for key in expired:
            self._evict_key(key)

    def _ensure_capacity(self) -> None:
        while len(self._store) > self.max_entries:
            oldest_key, _ = self._store.popitem(last=False)
            self._timestamps.pop(oldest_key, None)
            logger.debug("Cache LRU: rimossa voce obsoleta con indice %s", oldest_key)

    def __getitem__(self, key: int) -> Any:
        if key not in self._store:
            raise KeyError(key)
        if self._is_expired(key):
            self._evict_key(key)
            raise KeyError(key)
        self._store.move_to_end(key)
        return self._store[key]

    def __setitem__(self, key: int, value: Any) -> None:
        self._store[key] = value
        self._timestamps[key] = self._now()
        self._store.move_to_end(key)
        self._purge_expired()
        self._ensure_capacity()

    def __delitem__(self, key: int) -> None:
        if key in self._store:
            self._store.pop(key, None)
        self._timestamps.pop(key, None)

    def __iter__(self) -> Iterator[int]:
        self._purge_expired()
        return iter(self._store.copy())

    def __len__(self) -> int:
        self._purge_expired()
        return len(self._store)

    def __contains__(self, key: object) -> bool:  # type: ignore[override]
        if not isinstance(key, int):
            return False
        if key not in self._store:
            return False
        if self._is_expired(key):
            self._evict_key(key)
            return False
        return True

    def clear(self) -> None:  # type: ignore[override]
        self._store.clear()
        self._timestamps.clear()

    def get(self, key: int, default: Any = None) -> Any:  # type: ignore[override]
        try:
            return self[key]
        except KeyError:
            return default

    def pop(self, key: int, default: Any = None) -> Any:  # type: ignore[override]
        if key in self and not self._is_expired(key):
            value = self._store.pop(key)
            self._timestamps.pop(key, None)
            return value
        if default is not None:
            return default
        raise KeyError(key)

    def items(self) -> Iterator[Tuple[int, Any]]:  # type: ignore[override]
        self._purge_expired()
        snapshot = list(self._store.items())
        for key, value in snapshot:
            if key in self:
                yield key, value


email_cache: TimedLRUCache = TimedLRUCache(max_entries=500, ttl_seconds=1800.0)
calendar_cache: TimedLRUCache = TimedLRUCache(max_entries=200, ttl_seconds=1200.0)


def clear_email_cache() -> None:
    """Clear the email cache."""
    email_cache.clear()
    logger.debug("Cache dei messaggi svuotata.")


def clear_calendar_cache() -> None:
    """Clear the calendar event cache."""
    calendar_cache.clear()
    logger.debug("Cache degli eventi svuotata.")


__all__ = ["email_cache", "calendar_cache", "clear_email_cache", "clear_calendar_cache", "TimedLRUCache"]
