"""Folder-related helpers and operations for the Outlook MCP server."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Set, Tuple

from .logger import logger
from .utils import (
    coerce_bool,
    describe_item_type,
    ensure_string_list,
    normalize_folder_path,
    parse_item_type_hint,
    safe_child_count,
    safe_entry_id,
    safe_folder_path,
    safe_folder_size,
    safe_store_id,
    safe_total_count,
    safe_unread_count,
    shorten_identifier,
)


def get_folder_by_name(namespace, folder_name: str):
    """Get a specific Outlook folder by name."""
    try:
        inbox = namespace.GetDefaultFolder(6)  # Inbox
        for folder in inbox.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder

        for folder in namespace.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder

            for subfolder in folder.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    return subfolder

        return None
    except Exception as exc:
        logger.exception("Impossibile accedere alla cartella '%s'.", folder_name)
        raise Exception(f"Impossibile accedere alla cartella {folder_name}: {exc}") from exc


def _get_folder_by_path(namespace, folder_path: str):
    """Resolve a folder using its Outlook FolderPath representation."""
    normalized = normalize_folder_path(folder_path)
    if not normalized:
        return None

    segments = [segment.strip() for segment in normalized.split("\\") if segment.strip()]
    if not segments:
        return None

    root_candidates = []
    try:
        for root in namespace.Folders:
            name = getattr(root, "Name", "")
            if name.strip().lower() == segments[0].lower():
                root_candidates.append(root)
    except Exception:
        root_candidates = []

    if not root_candidates:
        try:
            default_candidate = get_folder_by_name(namespace, segments[0])
        except Exception:
            default_candidate = None
        if default_candidate:
            root_candidates = [default_candidate]

    for root in root_candidates:
        current = root
        try:
            current_path = normalize_folder_path(current.FolderPath)
        except Exception:
            current_path = None
        if current_path and current_path.lower() == normalized.lower():
            return current

        success = True
        for segment in segments[1:]:
            next_folder = None
            try:
                for sub in current.Folders:
                    name = getattr(sub, "Name", "")
                    if name.strip().lower() == segment.lower():
                        next_folder = sub
                        break
            except Exception:
                next_folder = None
            if not next_folder:
                success = False
                break
            current = next_folder
        if success:
            return current

    return None


def get_folder_by_path(namespace, folder_path: str):
    """Public helper to resolve a folder by its Outlook path."""
    return _get_folder_by_path(namespace, folder_path)


def resolve_folder(
    namespace,
    *,
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
) -> Tuple[Optional[Any], List[str]]:
    """Attempt to resolve an Outlook folder using multiple strategies."""
    attempts: List[str] = []

    if folder_id:
        try:
            folder = namespace.GetFolderFromID(folder_id)
            if folder:
                return folder, attempts
        except Exception as exc:
            attempts.append(f"ID '{folder_id}': {exc}")

    if folder_path:
        try:
            folder = _get_folder_by_path(namespace, folder_path)
            if folder:
                return folder, attempts
        except Exception as exc:
            attempts.append(f"Percorso '{folder_path}': {exc}")
        else:
            attempts.append(f"Percorso '{folder_path}' non trovato")

    if folder_name:
        try:
            folder = get_folder_by_name(namespace, folder_name)
            if folder:
                return folder, attempts
        except Exception as exc:
            attempts.append(f"Nome '{folder_name}': {exc}")
        else:
            attempts.append(f"Nome '{folder_name}' non trovato")

    return None, attempts


def list_folders(
    namespace,
    *,
    root_folder_id: Optional[str] = None,
    root_folder_path: Optional[str] = None,
    root_folder_name: Optional[str] = None,
    max_depth: int = 2,
    include_counts: bool = True,
    include_ids: bool = False,
    include_store: bool = False,
    include_paths: bool = True,
) -> str:
    if root_folder_id or root_folder_path or root_folder_name:
        root, attempts = resolve_folder(
            namespace,
            folder_id=root_folder_id,
            folder_path=root_folder_path,
            folder_name=root_folder_name,
        )
        if not root:
            detail = "; ".join(attempts) if attempts else "cartella non trovata."
            return f"Errore: impossibile individuare la cartella specificata ({detail})."
        roots = [root]
        header = f"Cartelle a partire da '{getattr(root, 'Name', 'sconosciuta')}'"
    else:
        try:
            roots = list(namespace.Folders)
        except Exception:
            logger.exception("Impossibile enumerare le cartelle radice dell'account.")
            return "Errore durante l'accesso alle cartelle principali dell'account Outlook."
        header = "Cartelle di posta disponibili"

    include_counts_bool = coerce_bool(include_counts)
    include_ids_bool = coerce_bool(include_ids)
    include_store_bool = coerce_bool(include_store)
    include_paths_bool = coerce_bool(include_paths)

    seen_keys: Set[str] = set()
    lines: List[str] = [f"{header}:\n"]
    stack: List[Tuple[Any, int]] = [(folder, 0) for folder in roots]

    while stack:
        folder, depth = stack.pop(0)
        if depth > max_depth:
            continue

        entry_id = safe_entry_id(folder)
        folder_path = safe_folder_path(folder)
        dedupe_key = (entry_id or folder_path or str(id(folder))).lower()
        if dedupe_key in seen_keys:
            continue
        seen_keys.add(dedupe_key)

        name = getattr(folder, "Name", "(Senza nome)")
        indent = "  " * depth

        meta_parts: List[str] = []
        if include_counts_bool:
            unread_count = safe_unread_count(folder)
            total_count = safe_total_count(folder)
            count_parts: List[str] = []
            if unread_count is not None:
                count_parts.append(f"non letti={unread_count}")
            if total_count is not None:
                count_parts.append(f"totali={total_count}")
            if count_parts:
                meta_parts.append(", ".join(count_parts))

        if include_ids_bool and entry_id:
            meta_parts.append(f"id={shorten_identifier(entry_id, 28)}")
        if include_store_bool:
            store_id = safe_store_id(folder)
            if store_id:
                meta_parts.append(f"store={shorten_identifier(store_id, 28)}")
        if include_paths_bool and folder_path:
            meta_parts.append(f"path={folder_path}")

        child_count = safe_child_count(folder)
        if child_count:
            meta_parts.append(f"sottocartelle={child_count}")

        meta_segment = f" ({'; '.join(meta_parts)})" if meta_parts else ""
        lines.append(f"{indent}- {name}{meta_segment}")

        if depth < max_depth:
            try:
                children = list(folder.Folders)
            except Exception:
                children = []
            for child in children:
                stack.append((child, depth + 1))

    if len(lines) == 1:
        lines.append("Nessuna cartella trovata.")
    return "\n".join(lines).rstrip()


def folder_metadata(
    namespace,
    *,
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
    include_children: bool = False,
    max_children: int = 20,
    include_counts: bool = True,
) -> str:
    if not isinstance(max_children, int) or max_children < 0:
        return "Errore: 'max_children' deve essere un intero non negativo."

    include_children_bool = coerce_bool(include_children)
    include_counts_bool = coerce_bool(include_counts)
    folder, attempts = resolve_folder(
        namespace,
        folder_id=folder_id,
        folder_path=folder_path,
        folder_name=folder_name,
    )
    if not folder:
        detail = "; ".join(attempts) if attempts else "cartella non trovata."
        return f"Errore: impossibile individuare la cartella specificata ({detail})."

    entry_id = safe_entry_id(folder) or "N/D"
    store_id = safe_store_id(folder) or "N/D"
    folder_path_value = safe_folder_path(folder) or "N/D"
    item_type_value = getattr(folder, "DefaultItemType", None)
    item_type_label = describe_item_type(item_type_value)
    unread_count = safe_unread_count(folder) if include_counts_bool else None
    total_count = safe_total_count(folder) if include_counts_bool else None
    folder_size = safe_folder_size(folder)
    child_count = safe_child_count(folder)

    parent_path = None
    try:
        parent = getattr(folder, "Parent", None)
        if parent:
            parent_path = safe_folder_path(parent)
    except Exception:
        parent_path = None

    lines = [
        f"Metadati per la cartella '{getattr(folder, 'Name', '(sconosciuta)')}':",
        f"- Percorso: {folder_path_value}",
        f"- EntryID: {entry_id}",
        f"- StoreID: {store_id}",
        f"- Tipo elementi: {item_type_label}",
    ]

    if parent_path:
        lines.append(f"- Cartella padre: {parent_path}")
    if unread_count is not None:
        lines.append(f"- Messaggi non letti: {unread_count}")
    if total_count is not None:
        lines.append(f"- Elementi totali: {total_count}")
    if folder_size:
        lines.append(f"- Dimensione stimata: {folder_size}")
    if child_count is not None:
        lines.append(f"- Sottocartelle immediate: {child_count}")

    if include_children_bool and child_count:
        lines.append("")
        lines.append("Sottocartelle immediate:")
        listed = 0
        try:
            for child in folder.Folders:
                if listed >= max_children:
                    lines.append(f"- ... ({child_count - listed} ulteriori non mostrati)")
                    break
                child_name = getattr(child, "Name", "(senza nome)")
                meta_parts: List[str] = []
                if include_counts_bool:
                    unread = safe_unread_count(child)
                    total = safe_total_count(child)
                    if unread is not None:
                        meta_parts.append(f"non letti={unread}")
                    if total is not None:
                        meta_parts.append(f"totali={total}")
                child_path = safe_folder_path(child)
                if child_path:
                    meta_parts.append(f"path={child_path}")
                if meta_parts:
                    lines.append(f"- {child_name} ({'; '.join(meta_parts)})")
                else:
                    lines.append(f"- {child_name}")
                listed += 1
        except Exception:
            lines.append("- Impossibile enumerare le sottocartelle (accesso negato).")

    return "\n".join(lines)


def create_folder(
    parent_folder,
    *,
    new_folder_name: str,
    item_type: Optional[Any] = None,
    allow_existing: bool = False,
) -> Tuple[Any, str]:
    normalized_target = new_folder_name.strip()
    try:
        for child in parent_folder.Folders:
            if getattr(child, "Name", "").strip().lower() == normalized_target.lower():
                if allow_existing:
                    path = safe_folder_path(child) or getattr(child, "Name", "")
                    return child, f"La cartella '{normalized_target}' esiste gia: {path}"
                raise ValueError("Esiste gia una cartella con questo nome.")
    except Exception:
        logger.warning("Impossibile controllare le cartelle figlie prima della creazione.")

    item_type_value = parse_item_type_hint(item_type)

    try:
        if item_type_value is None:
            new_folder = parent_folder.Folders.Add(normalized_target)
        else:
            new_folder = parent_folder.Folders.Add(normalized_target, item_type_value)
    except Exception as exc:
        logger.exception("Errore durante la creazione della cartella figlia.")
        raise RuntimeError(f"Outlook ha rifiutato la creazione della cartella ({exc}).") from exc

    path = safe_folder_path(new_folder) or normalized_target
    entry_id = safe_entry_id(new_folder) or "N/D"
    item_label = describe_item_type(getattr(new_folder, "DefaultItemType", None))
    message = (
        f"Cartella '{normalized_target}' creata con successo "
        f"(percorso={path}, id={entry_id}, tipo={item_label})."
    )
    return new_folder, message


def rename_folder(target, new_name: str) -> None:
    parent = getattr(target, "Parent", None)
    if parent:
        try:
            for sibling in parent.Folders:
                if sibling is target:
                    continue
                if getattr(sibling, "Name", "").strip().lower() == new_name.strip().lower():
                    raise ValueError("Esiste gia una cartella con il nuovo nome desiderato.")
        except Exception:
            logger.debug("Impossibile controllare le cartelle sorelle durante la rinomina.")

    try:
        target.Name = new_name.strip()
        if hasattr(target, "Save"):
            target.Save()
    except Exception as exc:
        logger.exception("Outlook ha rifiutato la rinomina della cartella.")
        raise RuntimeError(f"Impossibile rinominare la cartella ({exc}).") from exc


def delete_folder(target) -> None:
    try:
        target.Delete()
    except Exception as exc:
        logger.exception("Outlook ha rifiutato l'eliminazione della cartella.")
        raise RuntimeError(f"Impossibile eliminare la cartella ({exc}).") from exc


__all__ = [
    "get_folder_by_name",
    "get_folder_by_path",
    "resolve_folder",
    "list_folders",
    "folder_metadata",
    "create_folder",
    "rename_folder",
    "delete_folder",
]
