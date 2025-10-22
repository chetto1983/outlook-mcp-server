"""MCP tools for accessing Free/Busy calendar information."""

from __future__ import annotations

from typing import Optional, List, Dict, Any
import datetime

from ..features import feature_gate
from outlook_mcp.toolkit import mcp_tool

from outlook_mcp import logger


def _connect():
    from outlook_mcp import connect_to_outlook
    return connect_to_outlook()


def _parse_freebusy_string(fb_string: str, interval_minutes: int, start_time: datetime.datetime) -> List[Dict[str, Any]]:
    """Parse the free/busy string into time slots."""
    # Free/Busy status codes:
    # 0 = Free
    # 1 = Tentative
    # 2 = Busy
    # 3 = Out of Office

    status_map = {
        "0": "Libero",
        "1": "Provvisorio",
        "2": "Occupato",
        "3": "Fuori ufficio",
    }

    slots = []
    current_time = start_time

    for char in fb_string:
        status = status_map.get(char, f"Sconosciuto ({char})")
        end_time = current_time + datetime.timedelta(minutes=interval_minutes)

        slots.append({
            "start": current_time.strftime("%Y-%m-%d %H:%M"),
            "end": end_time.strftime("%Y-%m-%d %H:%M"),
            "status": status,
            "status_code": char,
        })

        current_time = end_time

    return slots


def _merge_consecutive_slots(slots: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Merge consecutive slots with the same status for better readability."""
    if not slots:
        return []

    merged = []
    current = slots[0].copy()

    for slot in slots[1:]:
        if slot["status"] == current["status"]:
            # Extend current slot
            current["end"] = slot["end"]
        else:
            # Save current and start new
            merged.append(current)
            current = slot.copy()

    # Add last slot
    merged.append(current)

    return merged


@mcp_tool()
@feature_gate(group="freebusy")
def get_freebusy_info(
    recipient_email: str,
    start_date: str,
    end_date: str,
    interval_minutes: int = 30,
    merge_slots: bool = True,
) -> str:
    """Recupera le informazioni di disponibilità (Free/Busy) per un destinatario.

    Args:
        recipient_email: Indirizzo email del destinatario
        start_date: Data/ora inizio in formato ISO (es: 2025-10-22 o 2025-10-22T08:00)
        end_date: Data/ora fine in formato ISO
        interval_minutes: Intervallo in minuti per ogni slot (default: 30)
        merge_slots: Se True, unisce slot consecutivi con stesso stato (default: True)
    """
    try:
        if not recipient_email.strip():
            return "Errore: specifica l'indirizzo email del destinatario"

        # Parse dates
        try:
            start_dt = datetime.datetime.fromisoformat(start_date.replace("T", " ").replace("Z", ""))
        except Exception as exc:
            return f"Errore: formato start_date non valido. Usa formato ISO (es: 2025-10-22 o 2025-10-22T08:00)"

        try:
            end_dt = datetime.datetime.fromisoformat(end_date.replace("T", " ").replace("Z", ""))
        except Exception as exc:
            return f"Errore: formato end_date non valido. Usa formato ISO (es: 2025-10-22 o 2025-10-22T18:00)"

        if start_dt >= end_dt:
            return "Errore: start_date deve essere precedente a end_date"

        if not isinstance(interval_minutes, int) or interval_minutes < 1 or interval_minutes > 1440:
            return "Errore: interval_minutes deve essere un intero tra 1 e 1440 (24 ore)"

        logger.info(
            "get_freebusy_info chiamato per %s dal %s al %s (intervallo: %s min)",
            recipient_email,
            start_date,
            end_date,
            interval_minutes,
        )

        _, namespace = _connect()

        # Create recipient
        try:
            recipient = namespace.CreateRecipient(recipient_email)
            recipient.Resolve()

            if not recipient.Resolved:
                return f"Errore: impossibile risolvere il destinatario '{recipient_email}'. Verifica l'indirizzo."
        except Exception as exc:
            logger.exception("Impossibile creare il destinatario.")
            return f"Errore: impossibile risolvere il destinatario ({exc})"

        # Get FreeBusy data
        # CompleteFormat parameter:
        # True = more detailed format with 4 states (Free, Tentative, Busy, OOF)
        # False = simple format with 2 states (Free, Busy)
        try:
            fb_string = recipient.FreeBusy(start_dt, interval_minutes, True)
        except Exception as exc:
            logger.exception("Impossibile recuperare i dati Free/Busy.")
            return f"Errore: impossibile recuperare i dati di disponibilità ({exc})"

        if not fb_string:
            return f"Nessun dato di disponibilità trovato per {recipient_email} nel periodo specificato."

        # Parse the free/busy string
        slots = _parse_freebusy_string(fb_string, interval_minutes, start_dt)

        # Merge consecutive slots if requested
        if merge_slots:
            slots = _merge_consecutive_slots(slots)

        # Build output
        lines = [
            f"Disponibilità per {recipient_email}:",
            f"Periodo: {start_dt.strftime('%Y-%m-%d %H:%M')} - {end_dt.strftime('%Y-%m-%d %H:%M')}",
            f"Intervallo: {interval_minutes} minuti",
            "",
        ]

        # Count statuses
        status_counts = {}
        for slot in slots:
            status = slot["status"]
            status_counts[status] = status_counts.get(status, 0) + 1

        lines.append("Riepilogo:")
        for status, count in sorted(status_counts.items()):
            lines.append(f"  - {status}: {count} slot")
        lines.append("")

        lines.append("Dettaglio slot:")
        for slot in slots:
            lines.append(f"  {slot['start']} - {slot['end']}: {slot['status']}")

        return "\n".join(lines)

    except Exception as exc:
        logger.exception("Errore durante get_freebusy_info.")
        return f"Errore durante il recupero dei dati di disponibilità: {exc}"


@mcp_tool()
@feature_gate(group="freebusy")
def find_free_time_slots(
    attendees: str,
    duration_minutes: int,
    start_date: str,
    end_date: str,
    working_hours_start: str = "08:00",
    working_hours_end: str = "18:00",
    interval_minutes: int = 30,
    max_results: int = 10,
) -> str:
    """Trova slot temporali liberi comuni per più partecipanti.

    Args:
        attendees: Indirizzi email separati da virgola o punto e virgola
        duration_minutes: Durata richiesta del meeting in minuti
        start_date: Data inizio ricerca (es: 2025-10-22)
        end_date: Data fine ricerca (es: 2025-10-25)
        working_hours_start: Ora inizio orario lavorativo (es: 08:00)
        working_hours_end: Ora fine orario lavorativo (es: 18:00)
        interval_minutes: Intervallo di ricerca in minuti (default: 30)
        max_results: Numero massimo di slot da restituire (default: 10)
    """
    try:
        if not attendees.strip():
            return "Errore: specifica almeno un partecipante"

        # Parse attendees
        attendee_list = [
            email.strip()
            for email in attendees.replace(";", ",").split(",")
            if email.strip()
        ]

        if not attendee_list:
            return "Errore: nessun indirizzo email valido fornito"

        if not isinstance(duration_minutes, int) or duration_minutes < 1:
            return "Errore: duration_minutes deve essere un intero positivo"

        if not isinstance(max_results, int) or max_results < 1 or max_results > 100:
            return "Errore: max_results deve essere un intero tra 1 e 100"

        # Parse dates
        try:
            base_start = datetime.datetime.fromisoformat(start_date.replace("T", " ").replace("Z", ""))
            base_end = datetime.datetime.fromisoformat(end_date.replace("T", " ").replace("Z", ""))
        except Exception:
            return "Errore: formato date non valido. Usa formato ISO (es: 2025-10-22)"

        # Parse working hours
        try:
            work_start_time = datetime.datetime.strptime(working_hours_start, "%H:%M").time()
            work_end_time = datetime.datetime.strptime(working_hours_end, "%H:%M").time()
        except Exception:
            return "Errore: formato orario lavorativo non valido. Usa formato HH:MM (es: 08:00)"

        logger.info(
            "find_free_time_slots chiamato per %s partecipanti, durata %s min, dal %s al %s",
            len(attendee_list),
            duration_minutes,
            start_date,
            end_date,
        )

        _, namespace = _connect()

        # Get FreeBusy data for all attendees
        attendee_fb_data = {}

        for email in attendee_list:
            try:
                recipient = namespace.CreateRecipient(email)
                recipient.Resolve()

                if not recipient.Resolved:
                    logger.warning("Impossibile risolvere destinatario: %s", email)
                    continue

                # Get full period FB data
                fb_string = recipient.FreeBusy(base_start, interval_minutes, True)
                if fb_string:
                    attendee_fb_data[email] = fb_string

            except Exception as exc:
                logger.warning("Errore recupero FB per %s: %s", email, exc)
                continue

        if not attendee_fb_data:
            return f"Errore: impossibile recuperare dati di disponibilità per nessuno dei partecipanti"

        # Find free slots
        free_slots = []
        current_date = base_start.date()

        while current_date <= base_end.date():
            # Set working hours for this date
            day_start = datetime.datetime.combine(current_date, work_start_time)
            day_end = datetime.datetime.combine(current_date, work_end_time)

            if day_start < base_start:
                day_start = base_start
            if day_end > base_end:
                day_end = base_end

            # Check slots within working hours
            current_time = day_start

            while current_time + datetime.timedelta(minutes=duration_minutes) <= day_end:
                # Calculate slot index in FB string
                time_diff = (current_time - base_start).total_seconds() / 60
                slot_index = int(time_diff / interval_minutes)
                slots_needed = max(1, duration_minutes // interval_minutes)

                # Check if all attendees are free for the duration
                all_free = True

                for email, fb_string in attendee_fb_data.items():
                    # Check all required slots
                    for i in range(slots_needed):
                        idx = slot_index + i
                        if idx >= len(fb_string):
                            all_free = False
                            break
                        # '0' = Free, anything else is not free
                        if fb_string[idx] != '0':
                            all_free = False
                            break

                    if not all_free:
                        break

                if all_free:
                    end_time = current_time + datetime.timedelta(minutes=duration_minutes)
                    free_slots.append({
                        "start": current_time.strftime("%Y-%m-%d %H:%M"),
                        "end": end_time.strftime("%Y-%m-%d %H:%M"),
                        "day": current_time.strftime("%A %d/%m/%Y"),
                    })

                    if len(free_slots) >= max_results:
                        break

                current_time += datetime.timedelta(minutes=interval_minutes)

            if len(free_slots) >= max_results:
                break

            current_date += datetime.timedelta(days=1)

        # Build output
        lines = [
            f"Ricerca slot liberi comuni per meeting di {duration_minutes} minuti:",
            f"Partecipanti ({len(attendee_fb_data)} con dati disponibili): {', '.join(attendee_fb_data.keys())}",
            f"Periodo: {start_date} - {end_date}",
            f"Orario lavorativo: {working_hours_start} - {working_hours_end}",
            "",
        ]

        if not free_slots:
            lines.append("Nessuno slot libero comune trovato nel periodo specificato.")
            lines.append("")
            lines.append("Suggerimenti:")
            lines.append("  - Aumenta il periodo di ricerca (end_date)")
            lines.append("  - Riduci la durata richiesta (duration_minutes)")
            lines.append("  - Estendi l'orario lavorativo")
        else:
            lines.append(f"Trovati {len(free_slots)} slot liberi comuni:")
            lines.append("")

            for idx, slot in enumerate(free_slots, 1):
                lines.append(f"{idx}. {slot['day']}")
                lines.append(f"   {slot['start']} - {slot['end']}")
                lines.append("")

        return "\n".join(lines)

    except Exception as exc:
        logger.exception("Errore durante find_free_time_slots.")
        return f"Errore durante la ricerca degli slot liberi: {exc}"
