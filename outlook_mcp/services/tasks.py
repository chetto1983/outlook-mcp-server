"""Task-oriented reusable helpers for Outlook MCP tools."""

from __future__ import annotations

import datetime
from typing import Any, Dict, Iterable, List, Optional, Sequence

from outlook_mcp import logger
from outlook_mcp.com import OutlookComError, run_com_call, wrap_com_exception
from outlook_mcp.utils import ensure_naive_datetime, build_body_preview, safe_folder_path, to_python_datetime
from outlook_mcp.constants import TASK_STATUS_MAP, TASK_PRIORITY_MAP, TASK_STATUS_REVERSE_MAP, TASK_PRIORITY_REVERSE_MAP

from .common import format_yes_no

__all__ = [
    "get_all_task_folders",
    "get_task_folder_by_name",
    "format_task_item",
    "get_tasks_from_folder",
    "collect_tasks_across_folders",
    "present_task_listing",
    "parse_task_status",
    "parse_task_priority",
]


def get_all_task_folders(namespace) -> List:
    """Return every Outlook folder that stores tasks."""
    task_folders: List = []
    visited_paths = set()

    def visit(folder) -> None:
        try:
            path = folder.FolderPath
        except Exception:
            path = str(folder)
        if path in visited_paths:
            return
        visited_paths.add(path)

        try:
            default_item_type = folder.DefaultItemType
        except Exception:
            default_item_type = None

        if default_item_type == 3:  # olTaskItem
            task_folders.append(folder)

        try:
            for sub in folder.Folders:
                visit(sub)
        except Exception:
            return

    try:
        default_tasks = namespace.GetDefaultFolder(13)  # olFolderTasks
        visit(default_tasks)
    except Exception:
        logger.warning("Impossibile ottenere la cartella Attività predefinita.")

    try:
        for root in namespace.Folders:
            visit(root)
    except Exception:
        logger.warning("Impossibile enumerare le radici per le attività.")

    logger.debug("Rilevate %s cartelle attività totali.", len(task_folders))
    return task_folders


def get_task_folder_by_name(namespace, folder_name: str):
    """Find a task folder by its display name."""
    if not folder_name:
        return namespace.GetDefaultFolder(13)
    target = folder_name.lower()
    for folder in get_all_task_folders(namespace):
        try:
            if folder.Name.lower() == target:
                return folder
        except Exception:
            continue
    return None


def format_task_item(task) -> Dict[str, Any]:
    """Generate a structured representation of an Outlook task."""
    try:
        subject = getattr(task, "Subject", "") or "(Senza oggetto)"
        body = getattr(task, "Body", "") or ""
        preview = build_body_preview(body, max_chars=220)

        # Date handling
        due_date_raw = getattr(task, "DueDate", None)
        start_date_raw = getattr(task, "StartDate", None)
        completed_date_raw = getattr(task, "DateCompleted", None)
        created_date_raw = getattr(task, "CreationTime", None)

        def fmt_date(dt_raw) -> Optional[str]:
            if not dt_raw:
                return None
            dt = to_python_datetime(dt_raw)
            if not dt:
                return None
            # Check if it's a valid date (Outlook sometimes uses 1/1/4501 for "no date")
            if dt.year > 4000:
                return None
            return dt.strftime("%Y-%m-%d %H:%M")

        due_date = fmt_date(due_date_raw)
        start_date = fmt_date(start_date_raw)
        completed_date = fmt_date(completed_date_raw)
        created_date = fmt_date(created_date_raw)

        # Status and priority
        status_code = getattr(task, "Status", 0)
        priority_code = getattr(task, "Importance", 1)
        percent_complete = getattr(task, "PercentComplete", 0)

        status_label = TASK_STATUS_MAP.get(status_code, f"Sconosciuto ({status_code})")
        priority_label = TASK_PRIORITY_MAP.get(priority_code, f"Sconosciuto ({priority_code})")

        # Other properties
        complete = getattr(task, "Complete", False)
        owner = getattr(task, "Owner", "") or ""
        categories = getattr(task, "Categories", "") or ""
        reminder_set = getattr(task, "ReminderSet", False)
        reminder_time = fmt_date(getattr(task, "ReminderTime", None)) if reminder_set else None

        # Get folder path
        try:
            parent_folder = task.Parent
            folder_path = safe_folder_path(parent_folder)
        except Exception:
            folder_path = ""

        # Entry ID
        try:
            entry_id = task.EntryID
        except Exception:
            entry_id = ""

        return {
            "id": entry_id,
            "subject": subject,
            "body": body,
            "preview": preview,
            "due_date": due_date,
            "start_date": start_date,
            "completed_date": completed_date,
            "created_date": created_date,
            "status": status_label,
            "status_code": status_code,
            "priority": priority_label,
            "priority_code": priority_code,
            "percent_complete": percent_complete,
            "complete": complete,
            "owner": owner,
            "categories": categories,
            "reminder_set": reminder_set,
            "reminder_time": reminder_time,
            "folder_path": folder_path,
        }
    except Exception as exc:
        logger.warning("Impossibile formattare l'attività: %s", exc, exc_info=True)
        return {
            "id": "",
            "subject": "(Errore nel recupero)",
            "preview": "",
            "status": "Sconosciuto",
            "priority": "Normale",
        }


def get_tasks_from_folder(
    folder,
    days: Optional[int] = None,
    include_completed: bool = False,
    search_term: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Retrieve tasks from a single folder with optional filters."""
    tasks = []

    try:
        items = folder.Items
        items.Sort("[DueDate]", False)  # Sort by due date, descending

        # Build filter
        filters = []

        if days is not None and days > 0:
            cutoff = datetime.datetime.now() - datetime.timedelta(days=days)
            cutoff_str = cutoff.strftime("%m/%d/%Y")
            filters.append(f"[CreationTime] >= '{cutoff_str}'")

        if not include_completed:
            filters.append("[Complete] = False")

        if search_term:
            # Simple search in subject and body
            search_filter = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{search_term}%' OR \"urn:schemas:httpmail:textdescription\" LIKE '%{search_term}%'"
            filters.append(search_filter)

        # Apply combined filter
        if filters:
            combined_filter = " AND ".join(f"({f})" for f in filters)
            try:
                items = items.Restrict(combined_filter)
            except Exception as exc:
                logger.warning("Filtro attività non riuscito, uso raccolta completa: %s", exc)

        count = 0
        max_items = 500  # Safety limit

        for item in items:
            if count >= max_items:
                logger.warning("Limite di %s attività raggiunto per la cartella.", max_items)
                break

            try:
                task_data = format_task_item(item)
                tasks.append(task_data)
                count += 1
            except Exception as exc:
                logger.debug("Errore nel processamento di un'attività: %s", exc)
                continue

        logger.debug("Recuperate %s attività dalla cartella.", count)

    except Exception as exc:
        logger.exception("Errore nel recupero delle attività dalla cartella.")

    return tasks


def collect_tasks_across_folders(
    folders: Sequence,
    days: Optional[int] = None,
    include_completed: bool = False,
    search_term: Optional[str] = None,
    target_total: Optional[int] = None,
) -> List[Dict[str, Any]]:
    """Collect tasks from multiple folders and merge them."""
    all_tasks: List[Dict[str, Any]] = []

    for folder in folders:
        try:
            folder_tasks = get_tasks_from_folder(folder, days, include_completed, search_term)
            all_tasks.extend(folder_tasks)

            if target_total and len(all_tasks) >= target_total:
                break
        except Exception as exc:
            logger.warning("Errore nel processamento cartella attività: %s", exc)
            continue

    # Sort by due date (tasks without due date go to the end)
    def sort_key(task: Dict[str, Any]):
        due_date = task.get("due_date")
        if not due_date:
            return ("9999-99-99", task.get("subject", ""))
        return (due_date, task.get("subject", ""))

    all_tasks.sort(key=sort_key)

    logger.info("Raccolte %s attività totali da %s cartelle.", len(all_tasks), len(folders))
    return all_tasks


def present_task_listing(
    tasks: List[Dict[str, Any]],
    folder_display: str,
    max_results: int,
    include_preview: bool,
    log_context: str,
) -> str:
    """Format tasks for presentation to the user."""
    from outlook_mcp import task_cache, clear_task_cache

    clear_task_cache()

    if not tasks:
        return f"Nessuna attività trovata in {folder_display}."

    # Limit results
    total_found = len(tasks)
    tasks_to_show = tasks[:max_results]

    lines = [
        f"Attività in {folder_display} (visualizzate {len(tasks_to_show)} di {total_found}):",
        "",
    ]

    for idx, task in enumerate(tasks_to_show, start=1):
        task_cache[idx] = task

        subject = task.get("subject", "(Senza oggetto)")
        status = task.get("status", "Sconosciuto")
        priority = task.get("priority", "Normale")
        due_date = task.get("due_date", "Nessuna scadenza")
        percent = task.get("percent_complete", 0)
        categories = task.get("categories", "")

        line_parts = [f"{idx}. {subject}"]

        details = []
        details.append(f"Stato: {status}")
        if percent > 0:
            details.append(f"{percent}%")
        details.append(f"Priorità: {priority}")
        details.append(f"Scadenza: {due_date}")

        if categories:
            details.append(f"Categorie: {categories}")

        lines.append(f"{line_parts[0]}")
        lines.append(f"   {' | '.join(details)}")

        if include_preview and task.get("preview"):
            lines.append(f"   Anteprima: {task['preview']}")

        lines.append("")

    if total_found > max_results:
        lines.append(f"(Altre {total_found - max_results} attività non visualizzate)")

    logger.info("%s ha presentato %s attività.", log_context, len(tasks_to_show))
    return "\n".join(lines)


def parse_task_status(status_input: Optional[str]) -> Optional[int]:
    """Parse a task status string to Outlook status code."""
    if status_input is None:
        return None
    normalized = status_input.strip().lower()
    return TASK_STATUS_REVERSE_MAP.get(normalized)


def parse_task_priority(priority_input: Optional[str]) -> Optional[int]:
    """Parse a task priority string to Outlook priority code."""
    if priority_input is None:
        return None
    normalized = priority_input.strip().lower()
    return TASK_PRIORITY_REVERSE_MAP.get(normalized)
