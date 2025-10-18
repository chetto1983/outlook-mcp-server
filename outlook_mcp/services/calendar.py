"""Calendar-oriented reusable helpers for Outlook MCP tools."""

from __future__ import annotations

import datetime
from typing import Any, Dict, Iterable, List, Optional, Sequence

from outlook_mcp import calendar_cache, clear_calendar_cache, logger
from outlook_mcp.utils import build_body_preview, safe_folder_path, to_python_datetime

from .common import format_yes_no

__all__ = [
    "get_all_calendar_folders",
    "get_calendar_folder_by_name",
    "format_calendar_event",
    "get_events_from_folder",
    "collect_events_across_calendars",
    "present_event_listing",
]


def get_all_calendar_folders(namespace) -> List:
    """Return every Outlook folder that stores appointments."""
    calendar_folders: List = []
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

        if default_item_type == 1:  # olAppointmentItem
            calendar_folders.append(folder)

        try:
            for sub in folder.Folders:
                visit(sub)
        except Exception:
            return

    try:
        default_calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar
        visit(default_calendar)
    except Exception:
        logger.warning("Impossibile ottenere la cartella Calendario predefinita.")

    try:
        for root in namespace.Folders:
            visit(root)
    except Exception:
        logger.warning("Impossibile enumerare le radici per i calendari.")

    logger.debug("Rilevate %s cartelle calendario totali.", len(calendar_folders))
    return calendar_folders


def get_calendar_folder_by_name(namespace, calendar_name: str):
    """Find a calendar folder by its display name."""
    if not calendar_name:
        return namespace.GetDefaultFolder(9)
    target = calendar_name.lower()
    for folder in get_all_calendar_folders(namespace):
        try:
            if folder.Name.lower() == target:
                return folder
        except Exception:
            continue
    return None


def format_calendar_event(appointment) -> Dict[str, Any]:
    """Generate a structured representation of an Outlook appointment."""
    try:
        start_dt = to_python_datetime(getattr(appointment, "Start", None))
        end_dt = to_python_datetime(getattr(appointment, "End", None))

        def fmt(dt: Optional[datetime.datetime]) -> Optional[str]:
            if not dt:
                return None
            return dt.strftime("%Y-%m-%d %H:%M")

        start_iso = start_dt.strftime("%Y-%m-%dT%H:%M") if start_dt else None
        end_iso = end_dt.strftime("%Y-%m-%dT%H:%M") if end_dt else None

        required = getattr(appointment, "RequiredAttendees", "") or ""
        optional = getattr(appointment, "OptionalAttendees", "") or ""
        body = getattr(appointment, "Body", "") or ""
        preview = build_body_preview(body, max_chars=320)

        event_data = {
            "id": getattr(appointment, "EntryID", None),
            "subject": getattr(appointment, "Subject", ""),
            "location": getattr(appointment, "Location", ""),
            "start_time": fmt(start_dt),
            "end_time": fmt(end_dt),
            "start_iso": start_iso,
            "end_iso": end_iso,
            "organizer": getattr(appointment, "Organizer", ""),
            "required_attendees": required,
            "optional_attendees": optional,
            "all_day": getattr(appointment, "AllDayEvent", False),
            "is_recurring": getattr(appointment, "IsRecurring", False),
            "body": body,
            "preview": preview,
            "folder_path": safe_folder_path(appointment),
            "categories": getattr(appointment, "Categories", ""),
            "duration_minutes": int((end_dt - start_dt).total_seconds() / 60) if start_dt and end_dt else None,
        }
        return event_data
    except Exception as exc:
        raise Exception(f"Impossibile formattare l'evento di calendario: {exc}")


def get_events_from_folder(folder, days: int, search_term: Optional[str] = None) -> List[Dict[str, Any]]:
    """Retrieve upcoming events from a calendar folder."""
    now = datetime.datetime.now()
    horizon = now + datetime.timedelta(days=days)
    events: List[Dict[str, Any]] = []

    try:
        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
    except Exception:
        logger.warning("Impossibile accedere o ordinare gli eventi della cartella calendario '%s'.", getattr(folder, "Name", folder))
        return events

    search_terms: List[str] = []
    if search_term:
        search_terms = [term.strip().lower() for term in search_term.split(" OR ") if term.strip()]

    def fmt(dt: datetime.datetime) -> str:
        return dt.strftime("%m/%d/%Y %I:%M %p")

    start_bound = now - datetime.timedelta(days=1)
    find_filter = f"[End] >= '{fmt(start_bound)}'"

    try:
        current = items.Find(find_filter)
        if current:
            logger.info(
                "Find iniziale per la cartella calendario '%s' ha trovato l'evento '%s'.",
                getattr(folder, "Name", folder),
                getattr(current, "Subject", None),
            )
        else:
            logger.info(
                "Find iniziale per la cartella calendario '%s' non ha trovato elementi (filtro=%s).",
                getattr(folder, "Name", folder),
                find_filter,
            )
    except Exception:
        logger.debug(
            "Errore durante l'esecuzione di Find sugli eventi della cartella calendario '%s'.",
            getattr(folder, "Name", folder),
            exc_info=True,
        )
        current = None

    scanned = 0
    max_scan = 500
    skip_counters = {
        "no_start": 0,
        "past": 0,
        "beyond": 0,
        "search": 0,
        "errors": 0,
    }

    def process_appointment(appointment) -> str:
        nonlocal scanned
        scanned += 1
        try:
            start_dt = to_python_datetime(getattr(appointment, "Start", None))
            end_dt = to_python_datetime(getattr(appointment, "End", None))

            if not start_dt:
                skip_counters["no_start"] += 1
                return "skip"
            if end_dt and end_dt < now:
                skip_counters["past"] += 1
                return "skip"
            if start_dt > horizon:
                skip_counters["beyond"] += 1
                return "break"

            if search_terms:
                haystack = " ".join(
                    filter(
                        None,
                        [
                            getattr(appointment, "Subject", ""),
                            getattr(appointment, "Location", ""),
                            getattr(appointment, "Organizer", ""),
                            getattr(appointment, "Body", ""),
                            getattr(appointment, "RequiredAttendees", "") or "",
                            getattr(appointment, "OptionalAttendees", "") or "",
                        ],
                    )
                ).lower()
                if not any(term in haystack for term in search_terms):
                    skip_counters["search"] += 1
                    return "skip"

            event_data = format_calendar_event(appointment)
            events.append(event_data)
            return "added"
        except Exception:
            skip_counters["errors"] += 1
            logger.debug("Evento calendario ignorato per errore di elaborazione.", exc_info=True)
            return "skip"

    if current:
        while current:
            outcome = process_appointment(current)
            if outcome == "break":
                break
            if scanned >= max_scan and events:
                break
            try:
                current = items.FindNext()
            except Exception:
                break
    else:
        logger.info(
            "Iterazione completa sugli eventi della cartella calendario '%s' (filtro Find non disponibile).",
            getattr(folder, "Name", folder),
        )
        for appointment in items:
            outcome = process_appointment(appointment)
            if outcome == "break":
                break
            if scanned >= max_scan and events:
                break

    logger.info(
        "Recuperati %s eventi dalla cartella calendario '%s'.",
        len(events),
        getattr(folder, "Name", folder),
    )
    logger.info(
        "Dettagli scansione calendario '%s': analizzati=%s salti=%s",
        getattr(folder, "Name", folder),
        scanned,
        skip_counters,
    )
    return events


def collect_events_across_calendars(
    folders: Sequence,
    days: int,
    search_term: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Aggregate calendar events across multiple folders."""
    aggregated: Dict[str, Dict[str, Any]] = {}
    for folder in folders:
        folder_events = get_events_from_folder(folder, days, search_term)
        for event in folder_events:
            event_id = event.get("id")
            if not event_id:
                continue
            if event_id not in aggregated:
                aggregated[event_id] = event

    sorted_events = sorted(
        aggregated.values(),
        key=lambda evt: evt.get("start_iso") or "",
    )
    logger.info(
        "Aggregati %s eventi provenienti da %s cartelle calendario.",
        len(sorted_events),
        len(folders),
    )
    return sorted_events


def present_event_listing(
    events: Sequence[Dict[str, Any]],
    calendar_display: str,
    days: int,
    max_results: int,
    include_description: bool,
    log_context: str,
) -> str:
    """Common presenter for calendar events."""
    clear_calendar_cache()

    if not events:
        logger.info(
            "%s: nessun evento trovato (calendario=%s giorni=%s).",
            log_context,
            calendar_display,
            days,
        )
        return f"Nessun evento in {calendar_display} nei prossimi {days} giorni."

    visible_events = list(events)[:max_results]
    visible_count = len(visible_events)
    total_count = len(events)

    if total_count > visible_count:
        header = (
            f"Trovati {total_count} eventi in {calendar_display} nei prossimi {days} giorni. "
            f"Mostro i primi {visible_count} risultati."
        )
    else:
        header = f"Trovati {visible_count} eventi in {calendar_display} nei prossimi {days} giorni."

    logger.info(
        "%s: restituiti %s eventi su %s (calendario=%s).",
        log_context,
        visible_count,
        total_count,
        calendar_display,
    )

    result = header + "\n\n"

    for idx, event in enumerate(visible_events, 1):
        calendar_cache[idx] = event

        result += f"Evento #{idx}\n"
        result += f"Oggetto: {event.get('subject', '(Senza oggetto)')}\n"
        result += f"Inizio: {event.get('start_time', 'Sconosciuto')}\n"
        result += f"Fine: {event.get('end_time', 'Sconosciuto')}\n"
        result += f"Calendario: {event.get('folder_path') or calendar_display}\n"
        result += f"Luogo: {event.get('location', '') or 'Non specificato'}\n"
        result += f"Organizzatore: {event.get('organizer', 'Non disponibile')}\n"
        result += f"Giornata intera: {format_yes_no(event.get('all_day'))}\n"
        if event.get("required_attendees"):
            result += f"Partecipanti obbligatori: {event['required_attendees']}\n"
        if event.get("optional_attendees"):
            result += f"Partecipanti facoltativi: {event['optional_attendees']}\n"
        if event.get("categories"):
            result += f"Categorie: {event['categories']}\n"
        if include_description and event.get("preview"):
            result += f"Anteprima: {event['preview']}\n"
        result += "\n"

    return result.rstrip()


# Legacy aliases
_present_event_listing = present_event_listing
