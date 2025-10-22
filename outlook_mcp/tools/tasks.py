"""MCP tools for managing Outlook tasks (to-do items)."""

from __future__ import annotations

from typing import Any, Optional, List, Dict
import datetime

from ..features import feature_gate
from outlook_mcp.toolkit import mcp_tool

from outlook_mcp import logger
from outlook_mcp.utils import coerce_bool, safe_entry_id
from outlook_mcp import MAX_TASK_DAYS, DEFAULT_TASK_MAX_RESULTS

# Reuse shared helpers from services
from outlook_mcp.services.tasks import (
    get_all_task_folders,
    get_task_folder_by_name,
    get_tasks_from_folder,
    collect_tasks_across_folders,
    present_task_listing,
    parse_task_status,
    parse_task_priority,
)


def _connect():
    from outlook_mcp import connect_to_outlook
    return connect_to_outlook()


@mcp_tool()
@feature_gate(group="tasks")
def list_tasks(
    days: Optional[int] = None,
    folder_name: Optional[str] = None,
    include_completed: bool = False,
    max_results: int = DEFAULT_TASK_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
) -> str:
    """Elenca le attività con filtri su giorni/cartelle/stato completamento."""
    if days is not None and (not isinstance(days, int) or days < 1 or days > MAX_TASK_DAYS):
        logger.warning("Valore 'days' non valido passato a list_tasks: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_TASK_DAYS} oppure omesso"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a list_tasks: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"

    include_completed_bool = coerce_bool(include_completed)
    include_preview_bool = coerce_bool(include_preview)
    include_all_bool = coerce_bool(include_all_folders)

    logger.info(
        "list_tasks chiamato con giorni=%s cartella=%s max_risultati=%s "
        "completate=%s anteprima=%s tutte_le_cartelle=%s",
        days,
        folder_name,
        max_results,
        include_completed_bool,
        include_preview_bool,
        include_all_bool,
    )

    try:
        _, namespace = _connect()

        tasks: List[Dict[str, Any]]
        folder_display: str

        if include_all_bool:
            if folder_name:
                logger.info("Parametro folder_name ignorato perché include_all_folders=True.")
            folders = get_all_task_folders(namespace)
            tasks = collect_tasks_across_folders(
                folders,
                days=days,
                include_completed=include_completed_bool,
                target_total=max_results,
            )
            folder_display = "Tutte le cartelle attività"
        else:
            if folder_name:
                folder = get_task_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella attività '{folder_name}' non trovata"
                folder_display = f"'{folder_name}'"
            else:
                folder = namespace.GetDefaultFolder(13)  # olFolderTasks
                folder_display = "Attività"

            tasks = get_tasks_from_folder(folder, days, include_completed_bool)

        return present_task_listing(
            tasks=tasks,
            folder_display=folder_display,
            max_results=max_results,
            include_preview=include_preview_bool,
            log_context="list_tasks",
        )
    except Exception as exc:
        logger.exception("Errore nel recupero delle attività per la cartella '%s'.", folder_name or "Attività")
        return f"Errore durante il recupero delle attività: {exc}"


@mcp_tool()
@feature_gate(group="tasks")
def search_tasks(
    search_term: str,
    days: Optional[int] = None,
    folder_name: Optional[str] = None,
    include_completed: bool = False,
    max_results: int = DEFAULT_TASK_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
) -> str:
    """Cerca attività per parole chiave con gli stessi filtri dell'elenco."""
    if not search_term:
        logger.warning("search_tasks chiamato senza termine di ricerca.")
        return "Errore: inserisci un termine di ricerca per le attività"

    if days is not None and (not isinstance(days, int) or days < 1 or days > MAX_TASK_DAYS):
        logger.warning("Valore 'days' non valido passato a search_tasks: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_TASK_DAYS} oppure omesso"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a search_tasks: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"

    include_completed_bool = coerce_bool(include_completed)
    include_preview_bool = coerce_bool(include_preview)
    include_all_bool = coerce_bool(include_all_folders)

    logger.info(
        "search_tasks chiamato con termine='%s' giorni=%s cartella=%s max_risultati=%s "
        "completate=%s anteprima=%s tutte_le_cartelle=%s",
        search_term,
        days,
        folder_name,
        max_results,
        include_completed_bool,
        include_preview_bool,
        include_all_bool,
    )

    try:
        _, namespace = _connect()

        tasks: List[Dict[str, Any]]
        folder_display: str

        if include_all_bool:
            if folder_name:
                logger.info("Parametro folder_name ignorato perché include_all_folders=True.")
            folders = get_all_task_folders(namespace)
            tasks = collect_tasks_across_folders(
                folders,
                days=days,
                include_completed=include_completed_bool,
                search_term=search_term,
                target_total=max_results,
            )
            folder_display = "Tutte le cartelle attività"
        else:
            if folder_name:
                folder = get_task_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella attività '{folder_name}' non trovata"
                folder_display = f"'{folder_name}'"
            else:
                folder = namespace.GetDefaultFolder(13)
                folder_display = "Attività"

            tasks = get_tasks_from_folder(folder, days, include_completed_bool, search_term)

        return present_task_listing(
            tasks=tasks,
            folder_display=folder_display,
            max_results=max_results,
            include_preview=include_preview_bool,
            log_context="search_tasks",
        )
    except Exception as exc:
        logger.exception(
            "Errore durante la ricerca di attività con termine '%s' nella cartella '%s'.",
            search_term,
            folder_name or "Attività",
        )
        return f"Errore nella ricerca delle attività: {exc}"


@mcp_tool()
@feature_gate(group="tasks")
def get_task_by_number(task_number: int) -> str:
    """Recupera i dettagli completi di un'attività dall'ultimo elenco."""
    try:
        from outlook_mcp import task_cache

        if not task_cache:
            logger.warning("get_task_by_number chiamato ma la cache attività è vuota.")
            return "Errore: nessun elenco attività attivo. Chiedimi di aggiornare le attività e poi ripeti la richiesta."

        if task_number not in task_cache:
            logger.warning("Attività numero %s non presente in cache per get_task_by_number.", task_number)
            return f"Errore: l'attività #{task_number} non è presente nell'elenco corrente."

        task = task_cache[task_number]
        logger.info("Recupero dettagli completi per l'attività #%s.", task_number)

        lines = [
            f"Dettagli attività #{task_number}:",
            "",
            f"Oggetto: {task.get('subject', '(Senza oggetto)')}",
            f"Stato: {task.get('status', 'Sconosciuto')}",
            f"Priorità: {task.get('priority', 'Normale')}",
            f"Completamento: {task.get('percent_complete', 0)}%",
        ]

        if task.get("due_date"):
            lines.append(f"Scadenza: {task['due_date']}")
        if task.get("start_date"):
            lines.append(f"Inizio: {task['start_date']}")
        if task.get("completed_date"):
            lines.append(f"Data completamento: {task['completed_date']}")
        if task.get("owner"):
            lines.append(f"Proprietario: {task['owner']}")
        if task.get("categories"):
            lines.append(f"Categorie: {task['categories']}")
        if task.get("reminder_set"):
            reminder_time = task.get("reminder_time", "Nessun orario")
            lines.append(f"Promemoria: Attivo ({reminder_time})")
        if task.get("folder_path"):
            lines.append(f"Cartella: {task['folder_path']}")

        body_content = task.get("body", "")
        if body_content and len(body_content) > 4000:
            body_content = body_content[:4000].rstrip() + "\n[Descrizione troncata per brevità]"

        lines.append("")
        lines.append("Descrizione completa:")
        lines.append(body_content or "(Nessuna descrizione)")

        return "\n".join(lines)
    except Exception as e:
        logger.exception("Errore nel recupero dei dettagli per l'attività #%s.", task_number)
        return f"Errore durante il recupero dei dettagli dell'attività: {str(e)}"


@mcp_tool()
@feature_gate(group="tasks")
def create_task(
    subject: str,
    body: Optional[str] = None,
    due_date: Optional[str] = None,
    start_date: Optional[str] = None,
    priority: Optional[str] = None,
    status: Optional[str] = None,
    reminder_time: Optional[str] = None,
    categories: Optional[str] = None,
    folder_name: Optional[str] = None,
) -> str:
    """Crea una nuova attività in Outlook."""
    try:
        if not subject.strip():
            return "Errore: specifica un oggetto per l'attività."

        logger.info(
            "create_task chiamato con oggetto='%s' scadenza=%s priorità=%s cartella=%s",
            subject,
            due_date,
            priority,
            folder_name,
        )

        outlook, namespace = _connect()

        # Create task item
        task = outlook.CreateItem(3)  # olTaskItem
        task.Subject = subject

        if body:
            task.Body = body

        # Parse and set dates
        if due_date:
            try:
                due_dt = datetime.datetime.fromisoformat(due_date.replace("T", " ").replace("Z", ""))
                task.DueDate = due_dt
            except Exception as exc:
                logger.warning("Formato data scadenza non valido: %s (%s)", due_date, exc)
                return f"Errore: formato data scadenza non valido. Usa formato ISO (es: 2025-10-22 o 2025-10-22T14:30)"

        if start_date:
            try:
                start_dt = datetime.datetime.fromisoformat(start_date.replace("T", " ").replace("Z", ""))
                task.StartDate = start_dt
            except Exception as exc:
                logger.warning("Formato data inizio non valido: %s (%s)", start_date, exc)
                return f"Errore: formato data inizio non valido. Usa formato ISO (es: 2025-10-22 o 2025-10-22T14:30)"

        # Set priority
        if priority:
            priority_code = parse_task_priority(priority)
            if priority_code is not None:
                task.Importance = priority_code
            else:
                return f"Errore: priorità non valida. Usa: bassa, normale, alta"

        # Set status
        if status:
            status_code = parse_task_status(status)
            if status_code is not None:
                task.Status = status_code
            else:
                return f"Errore: stato non valido. Usa: non iniziata, in corso, completata, in attesa, differita"

        # Set reminder
        if reminder_time:
            try:
                reminder_dt = datetime.datetime.fromisoformat(reminder_time.replace("T", " ").replace("Z", ""))
                task.ReminderSet = True
                task.ReminderTime = reminder_dt
            except Exception as exc:
                logger.warning("Formato data promemoria non valido: %s (%s)", reminder_time, exc)
                return f"Errore: formato data promemoria non valido. Usa formato ISO (es: 2025-10-22T14:30)"

        # Set categories
        if categories:
            task.Categories = categories

        # Move to specific folder if requested
        if folder_name:
            folder = get_task_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Errore: cartella attività '{folder_name}' non trovata"
            task.Save()
            task = task.Move(folder)
        else:
            task.Save()

        entry_id = safe_entry_id(task)
        return f"Attività creata: '{subject}' (task_id={entry_id or 'N/D'})"

    except Exception as exc:
        logger.exception("Errore durante create_task per oggetto '%s'.", subject)
        return f"Errore durante la creazione dell'attività: {exc}"


@mcp_tool()
@feature_gate(group="tasks")
def update_task(
    task_number: Optional[int] = None,
    task_id: Optional[str] = None,
    subject: Optional[str] = None,
    body: Optional[str] = None,
    due_date: Optional[str] = None,
    start_date: Optional[str] = None,
    priority: Optional[str] = None,
    status: Optional[str] = None,
    percent_complete: Optional[int] = None,
    reminder_time: Optional[str] = None,
    categories: Optional[str] = None,
) -> str:
    """Aggiorna un'attività esistente."""
    try:
        if task_number is None and task_id is None:
            return "Errore: specifica task_number oppure task_id"

        logger.info(
            "update_task chiamato con numero=%s id=%s",
            task_number,
            task_id,
        )

        _, namespace = _connect()

        # Resolve task
        if task_number is not None:
            from outlook_mcp import task_cache
            if not task_cache or task_number not in task_cache:
                return f"Errore: attività #{task_number} non presente in cache. Elenca prima le attività."
            cached_task = task_cache[task_number]
            task_id = cached_task.get("id")

        if not task_id:
            return "Errore: impossibile determinare l'ID dell'attività"

        # Get task item
        try:
            task = namespace.GetItemFromID(task_id)
        except Exception as exc:
            logger.exception("Impossibile recuperare l'attività con ID %s.", task_id)
            return f"Errore: impossibile recuperare l'attività ({exc})"

        # Update fields
        updates = []

        if subject is not None:
            task.Subject = subject
            updates.append("oggetto")

        if body is not None:
            task.Body = body
            updates.append("descrizione")

        if due_date is not None:
            try:
                due_dt = datetime.datetime.fromisoformat(due_date.replace("T", " ").replace("Z", ""))
                task.DueDate = due_dt
                updates.append("scadenza")
            except Exception as exc:
                return f"Errore: formato data scadenza non valido ({exc})"

        if start_date is not None:
            try:
                start_dt = datetime.datetime.fromisoformat(start_date.replace("T", " ").replace("Z", ""))
                task.StartDate = start_dt
                updates.append("data inizio")
            except Exception as exc:
                return f"Errore: formato data inizio non valido ({exc})"

        if priority is not None:
            priority_code = parse_task_priority(priority)
            if priority_code is not None:
                task.Importance = priority_code
                updates.append("priorità")
            else:
                return "Errore: priorità non valida. Usa: bassa, normale, alta"

        if status is not None:
            status_code = parse_task_status(status)
            if status_code is not None:
                task.Status = status_code
                updates.append("stato")
            else:
                return "Errore: stato non valido. Usa: non iniziata, in corso, completata, in attesa, differita"

        if percent_complete is not None:
            if not isinstance(percent_complete, int) or percent_complete < 0 or percent_complete > 100:
                return "Errore: percent_complete deve essere tra 0 e 100"
            task.PercentComplete = percent_complete
            if percent_complete == 100:
                task.Status = 2  # olTaskComplete
            updates.append("percentuale completamento")

        if reminder_time is not None:
            try:
                reminder_dt = datetime.datetime.fromisoformat(reminder_time.replace("T", " ").replace("Z", ""))
                task.ReminderSet = True
                task.ReminderTime = reminder_dt
                updates.append("promemoria")
            except Exception as exc:
                return f"Errore: formato data promemoria non valido ({exc})"

        if categories is not None:
            task.Categories = categories
            updates.append("categorie")

        if not updates:
            return "Nessuna modifica specificata"

        task.Save()

        reference = f"#{task_number}" if task_number is not None else f"ID {task_id}"
        return f"Attività {reference} aggiornata: {', '.join(updates)}"

    except Exception as exc:
        logger.exception("Errore durante update_task.")
        return f"Errore durante l'aggiornamento dell'attività: {exc}"


@mcp_tool()
@feature_gate(group="tasks")
def mark_task_complete(
    task_number: Optional[int] = None,
    task_id: Optional[str] = None,
) -> str:
    """Contrassegna un'attività come completata."""
    try:
        if task_number is None and task_id is None:
            return "Errore: specifica task_number oppure task_id"

        logger.info("mark_task_complete chiamato con numero=%s id=%s", task_number, task_id)

        _, namespace = _connect()

        # Resolve task
        if task_number is not None:
            from outlook_mcp import task_cache
            if not task_cache or task_number not in task_cache:
                return f"Errore: attività #{task_number} non presente in cache. Elenca prima le attività."
            cached_task = task_cache[task_number]
            task_id = cached_task.get("id")

        if not task_id:
            return "Errore: impossibile determinare l'ID dell'attività"

        # Get task item
        try:
            task = namespace.GetItemFromID(task_id)
        except Exception as exc:
            logger.exception("Impossibile recuperare l'attività con ID %s.", task_id)
            return f"Errore: impossibile recuperare l'attività ({exc})"

        # Mark complete
        task.Status = 2  # olTaskComplete
        task.PercentComplete = 100
        task.DateCompleted = datetime.datetime.now()
        task.Save()

        reference = f"#{task_number}" if task_number is not None else f"ID {task_id}"
        return f"Attività {reference} contrassegnata come completata."

    except Exception as exc:
        logger.exception("Errore durante mark_task_complete.")
        return f"Errore durante il completamento dell'attività: {exc}"


@mcp_tool()
@feature_gate(group="tasks")
def delete_task(
    task_number: Optional[int] = None,
    task_id: Optional[str] = None,
) -> str:
    """Elimina un'attività."""
    try:
        if task_number is None and task_id is None:
            return "Errore: specifica task_number oppure task_id"

        logger.info("delete_task chiamato con numero=%s id=%s", task_number, task_id)

        _, namespace = _connect()

        # Resolve task
        if task_number is not None:
            from outlook_mcp import task_cache
            if not task_cache or task_number not in task_cache:
                return f"Errore: attività #{task_number} non presente in cache. Elenca prima le attività."
            cached_task = task_cache[task_number]
            task_id = cached_task.get("id")

        if not task_id:
            return "Errore: impossibile determinare l'ID dell'attività"

        # Get task item
        try:
            task = namespace.GetItemFromID(task_id)
        except Exception as exc:
            logger.exception("Impossibile recuperare l'attività con ID %s.", task_id)
            return f"Errore: impossibile recuperare l'attività ({exc})"

        # Delete
        task_subject = getattr(task, "Subject", "(Senza oggetto)")
        task.Delete()

        # Remove from cache
        if task_number is not None:
            from outlook_mcp import task_cache
            task_cache.pop(task_number, None)

        reference = f"#{task_number}" if task_number is not None else f"ID {task_id}"
        return f"Attività {reference} ('{task_subject}') eliminata."

    except Exception as exc:
        logger.exception("Errore durante delete_task.")
        return f"Errore durante l'eliminazione dell'attività: {exc}"
