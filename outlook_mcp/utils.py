"""General-purpose helper functions for the Outlook MCP server."""

from __future__ import annotations

import datetime
from typing import Any, Dict, Iterable, List, Optional

from .constants import (
    ATTACHMENT_NAME_PREVIEW_MAX,
    BODY_PREVIEW_MAX_CHARS,
    CONVERSATION_ID_PREVIEW_MAX,
    DEFAULT_ITEM_TYPE_LABELS,
    ITEM_TYPE_NAME_MAP,
)


def coerce_bool(value: Any) -> bool:
    """Best-effort conversion of user-provided values into booleans."""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "y", "yes", "on"}
    return bool(value)


def normalize_whitespace(text: Optional[str]) -> str:
    """Collapse whitespace so previews stay compact."""
    if not text:
        return ""
    return " ".join(text.split())


def build_body_preview(body: Optional[str], max_chars: int = BODY_PREVIEW_MAX_CHARS) -> str:
    """Create a trimmed preview of the email body for quick inspection."""
    normalized = normalize_whitespace(body)
    if not normalized:
        return ""
    if len(normalized) <= max_chars:
        return normalized
    return normalized[: max_chars - 3].rstrip() + "..."


def trim_conversation_id(conversation_id: Optional[str], max_chars: int = CONVERSATION_ID_PREVIEW_MAX) -> Optional[str]:
    """Shorten long conversation identifiers so they stay readable."""
    if not conversation_id:
        return None
    if len(conversation_id) <= max_chars:
        return conversation_id
    return conversation_id[: max_chars - 3] + "..."


def extract_recipients(mail_item) -> Dict[str, List[str]]:
    """Return recipients grouped by address type."""
    recipients_by_type = {"to": [], "cc": [], "bcc": []}
    if not hasattr(mail_item, "Recipients") or not mail_item.Recipients:
        return recipients_by_type

    type_mapping = {1: "to", 2: "cc", 3: "bcc"}  # Outlook constants
    for i in range(1, mail_item.Recipients.Count + 1):
        recipient = mail_item.Recipients(i)
        display_name = recipient.Name or "Sconosciuto"
        address = getattr(recipient, "Address", "") or ""
        formatted = f"{display_name} <{address}>" if address else display_name
        address_type = type_mapping.get(getattr(recipient, "Type", 1), "to")
        recipients_by_type[address_type].append(formatted)
    return recipients_by_type


def safe_folder_path(mail_item) -> str:
    """Return a readable folder path if available."""
    try:
        parent = getattr(mail_item, "Parent", None)
        return parent.FolderPath if parent else ""
    except Exception:
        return ""


def extract_attachment_names(mail_item, max_names: int = ATTACHMENT_NAME_PREVIEW_MAX) -> List[str]:
    """Return a small list of attachment names without downloading them."""
    names: List[str] = []
    if not hasattr(mail_item, "Attachments"):
        return names
    try:
        attachment_count = mail_item.Attachments.Count
    except Exception:
        attachment_count = 0
    if not attachment_count:
        return names

    for index in range(1, min(attachment_count, max_names) + 1):
        try:
            names.append(mail_item.Attachments(index).FileName)
        except Exception:
            continue
    if attachment_count > max_names:
        names.append(f"... (+{attachment_count - max_names} more)")
    return names


def safe_entry_id(obj: Any) -> Optional[str]:
    """Return the EntryID of an Outlook object when available."""
    try:
        entry_id = getattr(obj, "EntryID", None)
    except Exception:
        entry_id = None
    if not entry_id:
        return None
    entry_str = str(entry_id).strip()
    return entry_str or None


def safe_store_id(obj: Any) -> Optional[str]:
    """Return the StoreID of an Outlook folder when accessible."""
    try:
        store_id = getattr(obj, "StoreID", None)
    except Exception:
        store_id = None
    if not store_id:
        return None
    store_str = str(store_id).strip()
    return store_str or None


def shorten_identifier(value: Optional[str], max_chars: int = 24) -> Optional[str]:
    """Return a shortened identifier suitable for human-friendly output."""
    if not value:
        return None
    if len(value) <= max_chars:
        return value
    return value[: max_chars - 3] + "..."


def describe_item_type(item_type: Optional[int]) -> str:
    """Translate Outlook default item type constants into readable labels."""
    if item_type is None:
        return "Sconosciuto"
    return DEFAULT_ITEM_TYPE_LABELS.get(item_type, f"Sconosciuto ({item_type})")


def parse_item_type_hint(value: Optional[Any]) -> Optional[int]:
    """Convert a user hint into an Outlook default item type constant."""
    if value is None:
        return None
    if isinstance(value, int):
        return value
    try:
        text = str(value).strip().lower()
    except Exception:
        return None
    if not text:
        return None
    return ITEM_TYPE_NAME_MAP.get(text, None)


def safe_child_count(folder) -> Optional[int]:
    """Return the number of immediate subfolders when available."""
    try:
        folders = getattr(folder, "Folders", None)
        if folders is None:
            return None
        return folders.Count
    except Exception:
        return None


def safe_unread_count(folder) -> Optional[int]:
    """Return the number of unread items in a folder."""
    try:
        unread = getattr(folder, "UnreadItemCount", None)
        if unread is None:
            return None
        return int(unread)
    except Exception:
        return None


def safe_total_count(folder) -> Optional[int]:
    """Return the total number of items contained in a folder."""
    try:
        items = getattr(folder, "Items", None)
        if not items:
            return None
        return items.Count
    except Exception:
        return None


def safe_folder_size(folder) -> Optional[str]:
    """Return the textual size representation reported by Outlook."""
    try:
        size = getattr(folder, "FolderSize", None)
    except Exception:
        size = None
    if not size:
        return None
    size_str = str(size).strip()
    return size_str or None


def normalize_folder_path(path: Optional[str]) -> Optional[str]:
    """Standardise Outlook folder paths for comparison."""
    if not path:
        return None
    text = path.replace("/", "\\").strip()
    if not text:
        return None
    if text.startswith("\\\\"):
        text = text[2:]
    return text


def ensure_string_list(value: Optional[Any]) -> List[str]:
    """Normalize a user-supplied value into a list of non-empty strings."""
    if value is None:
        return []
    if isinstance(value, str):
        separators = [";", "|", ","]
        segments = [value]
        for separator in separators:
            if separator in value:
                segments = [part for part in value.split(separator)]
                break
        return [segment.strip() for segment in segments if segment.strip()]
    try:
        return [str(item).strip() for item in value if str(item).strip()]
    except TypeError:
        return []


def ensure_int_list(value: Optional[Any]) -> List[int]:
    """Normalize a value into a list of integers (ignoring invalid entries)."""
    if value is None:
        return []
    if isinstance(value, int):
        return [value]
    candidates: Iterable[str]
    if isinstance(value, str):
        candidates = ensure_string_list(value)
    else:
        try:
            candidates = [str(item) for item in value]
        except TypeError:
            candidates = []
    ints: List[int] = []
    for candidate in candidates:
        try:
            ints.append(int(candidate))
        except (TypeError, ValueError):
            continue
    return ints


def safe_filename(name: Optional[str], fallback: str = "allegato") -> str:
    """Return a filesystem-safe filename."""
    if not name:
        base = fallback
    else:
        invalid_chars = '<>:"/\\|?*'
        base = "".join("_" if ch in invalid_chars else ch for ch in name).strip()
        if not base:
            base = fallback
    return base


def to_python_datetime(raw_value: Any) -> Optional[datetime.datetime]:
    """Convert a COM datetime value into a Python datetime, if possible."""
    if not raw_value:
        return None
    if isinstance(raw_value, datetime.datetime):
        if raw_value.tzinfo:
            return raw_value.astimezone().replace(tzinfo=None)
        return raw_value
    try:
        timestamp = raw_value.timestamp()
        return datetime.datetime.fromtimestamp(timestamp)
    except Exception:
        try:
            parsed = datetime.datetime.fromisoformat(str(raw_value))
            if parsed.tzinfo:
                return parsed.astimezone().replace(tzinfo=None)
            return parsed
        except Exception:
            return None


__all__ = [
    "build_body_preview",
    "coerce_bool",
    "describe_item_type",
    "ensure_int_list",
    "ensure_string_list",
    "extract_attachment_names",
    "extract_recipients",
    "normalize_folder_path",
    "normalize_whitespace",
    "parse_item_type_hint",
    "safe_child_count",
    "safe_entry_id",
    "safe_folder_path",
    "safe_folder_size",
    "safe_store_id",
    "safe_total_count",
    "safe_unread_count",
    "safe_filename",
    "shorten_identifier",
    "trim_conversation_id",
    "to_python_datetime",
]
