"""MCP tool for creating Outlook calendar events (all-day or timed)."""

from __future__ import annotations

from typing import Any, Optional
import datetime

from ..features import feature_gate
from outlook_mcp.toolkit import mcp_tool

from outlook_mcp import logger, clear_calendar_cache, calendar_cache
from outlook_mcp.utils import ensure_string_list, ensure_naive_datetime, safe_entry_id, to_python_datetime
from outlook_mcp.services.common import parse_datetime_string
from outlook_mcp.services.calendar import get_calendar_folder_by_name


def _connect():
    from outlook_mcp import connect_to_outlook

    return connect_to_outlook()


def _parse_dt(value: Optional[str]):
    return parse_datetime_string(value)


def _calendar_by_name(namespace, name: Optional[str]):
    return get_calendar_folder_by_name(namespace, name) if name else namespace.GetDefaultFolder(9)


def _local_timezone() -> datetime.tzinfo:
    tz = datetime.datetime.now().astimezone().tzinfo
    return tz or datetime.timezone.utc


def _ensure_local(dt: datetime.datetime) -> datetime.datetime:
    tz = _local_timezone()
    if dt.tzinfo is None:
        return dt.replace(tzinfo=tz)
    try:
        return dt.astimezone(tz)
    except Exception:
        return dt


@mcp_tool()
@feature_gate(group="calendar.write")
def create_calendar_event(
    subject: str,
    start_time: str,
    duration_minutes: Optional[int] = 60,
    location: Optional[str] = None,
    body: Optional[str] = None,
    attendees: Optional[Any] = None,
    reminder_minutes: Optional[int] = 15,
    calendar_name: Optional[str] = None,
    all_day: bool = False,
    send_invitations: bool = True,
) -> str:
    """Crea un nuovo evento di calendario e, se richiesto, invia gli inviti."""
    if not subject or not subject.strip():
        return "Errore: specifica un oggetto ('subject') per l'evento."

    start_dt = _parse_dt(start_time)
    if not start_dt:
        return "Errore: 'start_time' deve essere una data valida (es. '2025-10-20 10:30')."

    all_day_bool = bool(all_day)
    send_bool = bool(send_invitations)

    if not all_day_bool:
        if duration_minutes is None:
            duration_value = 60
        else:
            try:
                duration_value = int(duration_minutes)
            except (TypeError, ValueError):
                return "Errore: 'duration_minutes' deve essere un intero positivo."
            if duration_value <= 0:
                return "Errore: 'duration_minutes' deve essere un intero positivo."
    else:
        duration_value = None

    attendee_list = ensure_string_list(attendees)

    reminder_set = False
    reminder_value: Optional[int] = None
    if reminder_minutes is not None:
        try:
            reminder_value = int(reminder_minutes)
        except (TypeError, ValueError):
            return "Errore: 'reminder_minutes' deve essere un intero (minuti)."
        if reminder_value >= 0:
            reminder_set = True
        else:
            reminder_value = None

    logger.info(
        "create_calendar_event chiamato (subject=%s start=%s durata=%s luogo=%s partecipanti=%s all_day=%s invia=%s calendario=%s).",
        subject,
        start_time,
        duration_value,
        location,
        attendee_list,
        all_day_bool,
        send_bool,
        calendar_name,
    )

    try:
        outlook, namespace = _connect()
        if calendar_name:
            target_calendar = _calendar_by_name(namespace, calendar_name)
            if not target_calendar:
                return f"Errore: calendario '{calendar_name}' non trovato."
        else:
            target_calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar

        calendar_display = getattr(target_calendar, "Name", "Calendario")
        appointment = outlook.CreateItem(1)  # olAppointmentItem

        subject_clean = subject.strip()
        appointment.Subject = subject_clean

        local_start = ensure_naive_datetime(_ensure_local(start_dt)) or start_dt

        normalized_start: Optional[datetime.datetime] = None
        if all_day_bool:
            normalized_start = local_start.replace(hour=0, minute=0, second=0, microsecond=0)
            appointment.Start = normalized_start
            appointment.End = normalized_start + datetime.timedelta(days=1)
            appointment.AllDayEvent = True
        else:
            appointment.Start = local_start
            appointment.AllDayEvent = False
            appointment.Duration = duration_value or 60
            appointment.End = local_start + datetime.timedelta(minutes=appointment.Duration)

        if location:
            appointment.Location = location
        if body:
            appointment.Body = body

        if attendee_list:
            appointment.MeetingStatus = 1  # olMeeting
            for email in attendee_list:
                if not email:
                    continue
                try:
                    recipient = appointment.Recipients.Add(email)
                    if hasattr(recipient, "Type"):
                        recipient.Type = 1  # Required attendee
                except Exception as exc:
                    logger.warning("Impossibile aggiungere il destinatario '%s': %s", email, exc)

        appointment.ReminderSet = reminder_set
        if reminder_set and reminder_value is not None:
            appointment.ReminderMinutesBeforeStart = reminder_value

        try:
            appointment.Save()
        except Exception as exc:
            logger.exception("Salvataggio dell'appuntamento fallito.")
            return f"Errore: impossibile salvare l'evento ({exc})."

        if calendar_name:
            try:
                moved = appointment.Move(target_calendar)
                if moved:
                    appointment = moved
            except Exception as exc:
                logger.exception("Impossibile spostare l'evento nel calendario '%s'.", calendar_display)
                return (
                    "Errore: salvataggio completato ma non è stato possibile spostare l'evento nel "
                    f"calendario '{calendar_display}' ({exc})."
                )

        if send_bool and attendee_list:
            try:
                appointment.Send()
            except Exception as exc:
                logger.exception("Invio degli inviti fallito.")
                return f"Errore: evento creato ma invio degli inviti fallito ({exc})."

        clear_calendar_cache()

        entry_id = safe_entry_id(appointment) or "N/D"
        display_start = normalized_start if all_day_bool and normalized_start else local_start
        start_display = display_start.strftime("%Y-%m-%d") if all_day_bool else display_start.strftime("%Y-%m-%d %H:%M")

        lines = [
            f"Evento '{subject_clean}' creato per {start_display} nel calendario '{calendar_display}'.",
            f"Message ID: {entry_id}",
        ]
        if attendee_list:
            lines.append(f"Partecipanti: {', '.join(attendee_list)}")
        if send_bool and attendee_list:
            lines.append("Inviti inviati ai partecipanti.")
        elif attendee_list:
            lines.append("Inviti non inviati (send_invitations impostato a False).")

        return "\n".join(lines)
    except Exception as exc:
        logger.exception("Errore durante create_calendar_event.")
        return f"Errore durante la creazione dell'evento: {exc}"


@mcp_tool()
@feature_gate(group="calendar.write")
def move_calendar_event(
    event_number: Optional[int] = None,
    entry_id: Optional[str] = None,
    new_start_time: Optional[str] = None,
    new_duration_minutes: Optional[int] = None,
    new_location: Optional[str] = None,
    new_calendar_name: Optional[str] = None,
    send_updates: bool = False,
) -> str:
    """Aggiorna orario/posizione di un evento esistente (identificato da numero elenco o EntryID)."""
    if entry_id:
        target_id = entry_id.strip()
        if not target_id:
            return "Errore: 'entry_id' non può essere vuoto."
    elif event_number is not None:
        try:
            number = int(event_number)
        except (TypeError, ValueError):
            return "Errore: 'event_number' deve essere un intero."
        if number not in calendar_cache:
            return "Errore: evento non trovato nella cache corrente. Elenca gli eventi e riprova."
        cached_event = calendar_cache[number]
        target_id = cached_event.get("id")
        if not target_id:
            return "Errore: l'evento selezionato non espone un EntryID valido."
    else:
        return "Errore: specifica almeno 'event_number' (dall'ultimo elenco) oppure 'entry_id'."

    if new_duration_minutes is not None:
        try:
            duration_value = int(new_duration_minutes)
        except (TypeError, ValueError):
            return "Errore: 'new_duration_minutes' deve essere un intero positivo."
        if duration_value <= 0:
            return "Errore: 'new_duration_minutes' deve essere un intero positivo."
    else:
        duration_value = None

    send_bool = bool(send_updates)

    logger.info(
        "move_calendar_event chiamato (entry_id=%s numero=%s nuovo_start=%s nuova_durata=%s nuova_location=%s nuovo_calendario=%s invia=%s).",
        entry_id or target_id,
        event_number,
        new_start_time,
        duration_value,
        new_location,
        new_calendar_name,
        send_bool,
    )

    try:
        outlook, namespace = _connect()
        try:
            appointment = namespace.GetItemFromID(target_id)
        except Exception as exc:
            logger.exception("Impossibile recuperare l'evento con EntryID=%s.", target_id)
            return f"Errore: impossibile recuperare l'evento (EntryID={target_id}): {exc}"

        if new_start_time:
            parsed_start = _parse_dt(new_start_time)
            if not parsed_start:
                return "Errore: 'new_start_time' non è in un formato valido (usa es. '2025-10-23 09:00')."
            local_start = ensure_naive_datetime(_ensure_local(parsed_start))
            if not local_start:
                local_start = ensure_naive_datetime(parsed_start) or parsed_start
            if getattr(appointment, "AllDayEvent", False):
                normalized = local_start.replace(hour=0, minute=0, second=0, microsecond=0)
                appointment.Start = normalized
                appointment.End = normalized + datetime.timedelta(days=1)
            else:
                appointment.Start = local_start

        if duration_value is not None and not getattr(appointment, "AllDayEvent", False):
            appointment.Duration = duration_value
            try:
                current_start = ensure_naive_datetime(to_python_datetime(getattr(appointment, "Start", None)))
                if current_start:
                    appointment.End = current_start + datetime.timedelta(minutes=duration_value)
            except Exception:
                pass

        if new_location:
            appointment.Location = new_location

        moved_calendar_display = None
        if new_calendar_name:
            target_calendar = _calendar_by_name(namespace, new_calendar_name)
            if not target_calendar:
                return f"Errore: calendario '{new_calendar_name}' non trovato."
            try:
                moved = appointment.Move(target_calendar)
                if moved:
                    appointment = moved
                    moved_calendar_display = getattr(target_calendar, "Name", new_calendar_name)
            except Exception as exc:
                logger.exception("Impossibile spostare l'evento nel calendario '%s'.", new_calendar_name)
                return f"Errore: impossibile spostare l'evento nel calendario '{new_calendar_name}': {exc}"

        try:
            appointment.Save()
        except Exception as exc:
            logger.exception("Salvataggio dell'evento aggiornato fallito.")
            return f"Errore: impossibile salvare le modifiche all'evento ({exc})."

        if send_bool and getattr(appointment, "MeetingStatus", 0) in (1, 3):  # olMeeting / olMeetingReceivedAndCanceled
            try:
                appointment.Send()
            except Exception as exc:
                logger.warning("Invio aggiornamenti meeting fallito: %s", exc)
                return (
                    "Evento aggiornato ma invio degli aggiornamenti ai partecipanti non riuscito "
                    f"({exc})."
                )

        clear_calendar_cache()

        summary_lines = ["Evento aggiornato con successo."]
        current_start = ensure_naive_datetime(to_python_datetime(getattr(appointment, "Start", None)))
        if current_start:
            summary_lines.append(f"Nuovo inizio: {current_start.strftime('%Y-%m-%d %H:%M')}")
        if not getattr(appointment, "AllDayEvent", False):
            summary_lines.append(f"Durata: {getattr(appointment, 'Duration', duration_value or 0)} minuti")
        if new_location:
            summary_lines.append(f"Nuova posizione: {appointment.Location}")
        if moved_calendar_display:
            summary_lines.append(f"Calendario: {moved_calendar_display}")

        return "\n".join(summary_lines)
    except Exception as exc:
        logger.exception("Errore durante move_calendar_event.")
        return f"Errore durante l'aggiornamento dell'evento: {exc}"
