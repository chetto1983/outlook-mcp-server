"""MCP tools for reading Outlook calendar events across calendars."""

from __future__ import annotations

from typing import Optional, Any, Dict, List

from ..features import feature_gate
from outlook_mcp_server import mcp

from outlook_mcp import logger
from outlook_mcp.utils import coerce_bool
from outlook_mcp import MAX_EVENT_LOOKAHEAD_DAYS

# Reuse server helpers to avoid duplication
from outlook_mcp_server import (
    get_all_calendar_folders,
    collect_events_across_calendars,
    get_events_from_folder,
    _present_event_listing,
)


def _connect():
    from outlook_mcp_server import connect_to_outlook

    return connect_to_outlook()


def _get_calendar_folder(namespace, calendar_name: Optional[str]):
    from outlook_mcp_server import get_calendar_folder_by_name

    return get_calendar_folder_by_name(namespace, calendar_name) if calendar_name else namespace.GetDefaultFolder(9)


@mcp.tool()
@feature_gate(group="calendar.read")
def list_upcoming_events(
    days: int = 7,
    calendar_name: Optional[str] = None,
    max_results: int = 50,
    include_description: bool = False,
    include_all_calendars: bool = False,
) -> str:
    """Elenca i prossimi eventi (fino a 90 giorni), con descrizione opzionale."""
    if not isinstance(days, int) or days < 1 or days > MAX_EVENT_LOOKAHEAD_DAYS:
        logger.warning("Valore 'days' non valido passato a list_upcoming_events: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_EVENT_LOOKAHEAD_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a list_upcoming_events: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"

    include_desc = coerce_bool(include_description)
    include_all = coerce_bool(include_all_calendars)
    logger.info(
        "list_upcoming_events chiamato con giorni=%s calendario=%s max_risultati=%s descrizione=%s tutti_i_calendari=%s",
        days,
        calendar_name,
        max_results,
        include_desc,
        include_all,
    )

    try:
        _, namespace = _connect()

        if include_all:
            calendars = get_all_calendar_folders(namespace)
            events = collect_events_across_calendars(calendars, days)
            calendar_display = "Tutti i calendari"
        else:
            calendar_folder = _get_calendar_folder(namespace, calendar_name)
            if not calendar_folder:
                return f"Errore: calendario '{calendar_name}' non trovato"
            calendar_display = calendar_folder.Name if calendar_name else "Calendario"
            events = get_events_from_folder(calendar_folder, days)

        return _present_event_listing(
            events=events,
            calendar_display=calendar_display,
            days=days,
            max_results=max_results,
            include_description=include_desc,
            log_context="list_upcoming_events",
        )
    except Exception as e:
        logger.exception("Errore durante il recupero degli eventi per il calendario '%s'.", calendar_name or "Calendario")
        return f"Errore durante il recupero degli eventi: {str(e)}"


@mcp.tool()
@feature_gate(group="calendar.read")
def search_calendar_events(
    search_term: str,
    days: int = 30,
    calendar_name: Optional[str] = None,
    max_results: int = 50,
    include_description: bool = False,
    include_all_calendars: bool = False,
) -> str:
    """Cerca eventi per parole chiave con gli stessi filtri dell'elenco."""
    if not search_term:
        logger.warning("search_calendar_events chiamato senza termine di ricerca.")
        return "Errore: inserisci un termine di ricerca per il calendario"

    if not isinstance(days, int) or days < 1 or days > MAX_EVENT_LOOKAHEAD_DAYS:
        logger.warning("Valore 'days' non valido passato a search_calendar_events: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_EVENT_LOOKAHEAD_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a search_calendar_events: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"

    include_desc = coerce_bool(include_description)
    include_all = coerce_bool(include_all_calendars)
    logger.info(
        "search_calendar_events chiamato con termine='%s' giorni=%s calendario=%s max_risultati=%s descrizione=%s tutti_i_calendari=%s",
        search_term,
        days,
        calendar_name,
        max_results,
        include_desc,
        include_all,
    )

    try:
        _, namespace = _connect()

        if include_all:
            calendars = get_all_calendar_folders(namespace)
            events = collect_events_across_calendars(calendars, days, search_term)
            calendar_display = "Tutti i calendari"
        else:
            calendar_folder = _get_calendar_folder(namespace, calendar_name)
            if not calendar_folder:
                return f"Errore: calendario '{calendar_name}' non trovato"
            calendar_display = calendar_folder.Name if calendar_name else "Calendario"
            events = get_events_from_folder(calendar_folder, days, search_term)

        return _present_event_listing(
            events=events,
            calendar_display=calendar_display,
            days=days,
            max_results=max_results,
            include_description=include_desc,
            log_context="search_calendar_events",
        )
    except Exception as e:
        logger.exception(
            "Errore durante la ricerca di eventi con termine '%s' nel calendario '%s'.",
            search_term,
            calendar_name or "Calendario",
        )
        return f"Errore durante la ricerca degli eventi: {str(e)}"


@mcp.tool()
@feature_gate(group="calendar.read")
def get_event_by_number(event_number: int) -> str:
    """Recupera i dettagli completi di un evento dall'ultimo elenco."""
    try:
        from outlook_mcp import calendar_cache

        if not calendar_cache:
            logger.warning("get_event_by_number chiamato ma la cache eventi e vuota.")
            return "Errore: nessun elenco eventi attivo. Chiedimi di aggiornare gli eventi e poi ripeti la richiesta."

        if event_number not in calendar_cache:
            logger.warning("Evento numero %s non presente in cache per get_event_by_number.", event_number)
            return f"Errore: l'evento #{event_number} non e presente nell'elenco corrente."

        event = calendar_cache[event_number]
        logger.info("Recupero dettagli completi per l'evento #%s.", event_number)

        lines = [
            f"Dettagli evento #{event_number}:",
            "",
            f"Oggetto: {event.get('subject', '(Senza oggetto)')}",
            f"Inizio: {event.get('start_time', 'Sconosciuto')}",
            f"Fine: {event.get('end_time', 'Sconosciuto')}",
            f"Luogo: {event.get('location', '') or 'Non specificato'}",
            f"Organizzatore: {event.get('organizer', 'Non disponibile')}",
            f"Calendario: {event.get('folder_path', '')}",
            f"Giornata intera: {'Si' if event.get('all_day') else 'No'}",
        ]

        if event.get("required_attendees"):
            lines.append(f"Partecipanti obbligatori: {event['required_attendees']}")
        if event.get("optional_attendees"):
            lines.append(f"Partecipanti facoltativi: {event['optional_attendees']}")
        if event.get("categories"):
            lines.append(f"Categorie: {event['categories']}")
        if event.get("preview"):
            lines.append(f"Anteprima descrizione: {event['preview']}")

        body_content = event.get("body", "")
        if body_content and len(body_content) > 4000:
            body_content = body_content[:4000].rstrip() + "\n[Descrizione troncata per brevita]"

        lines.append("")
        lines.append("Descrizione completa:")
        lines.append(body_content or "(Nessuna descrizione)")

        return "\n".join(lines)
    except Exception as e:
        logger.exception("Errore nel recupero dei dettagli per l'evento #%s.", event_number)
        return f"Errore durante il recupero dei dettagli dell'evento: {str(e)}"
