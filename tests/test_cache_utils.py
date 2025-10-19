import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from outlook_mcp.cache import TimedLRUCache
from outlook_mcp.utils import obfuscate_identifier


def test_timed_lru_cache_evicts_oldest_when_full():
    cache = TimedLRUCache(max_entries=2, ttl_seconds=None)

    cache[1] = "one"
    cache[2] = "two"
    cache[3] = "three"  # evicts key 1

    assert 1 not in cache
    assert cache[2] == "two"
    assert cache[3] == "three"


def test_timed_lru_cache_expires_by_ttl(monkeypatch):
    now = [0.0]

    def fake_monotonic():
        return now[0]

    monkeypatch.setattr("outlook_mcp.cache.time.monotonic", fake_monotonic)

    cache = TimedLRUCache(max_entries=3, ttl_seconds=10.0)
    cache[1] = "uno"
    assert cache[1] == "uno"

    now[0] = 11.0  # advance beyond TTL
    assert 1 not in cache
    with pytest.raises(KeyError):
        _ = cache[1]


def test_obfuscate_identifier_masks_value():
    masked = obfuscate_identifier("ABCDEF123456")
    assert masked.startswith("ABCD…")
    assert len(masked) > 8
    assert "123456" not in masked
