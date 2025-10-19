"""MCP tool for creating Outlook calendar events (all-day or timed)."""

from __future__ import annotations

from typing import Any, Optional
import datetime

from ..features import feature_gate
from outlook_mcp.toolkit import mcp_tool

from outlook_mcp import logger, clear_calendar_cache
from outlook_mcp.utils import ensure_string_list, to_python_datetime, safe_entry_id
from outlook_mcp.services.common import parse_datetime_string
from outlook_mcp.services.calendar import get_calendar_folder_by_name


def _connect():
    from outlook_mcp import connect_to_outlook

    return connect_to_outlook()


def _parse_dt(value: Optional[str]):
    return parse_datetime_string(value)


def _calendar_by_name(namespace, name: Optional[str]):
    return get_calendar_folder_by_name(namespace, name) if name else namespace.GetDefaultFolder(9)


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

        if all_day_bool:
            normalized_start = start_dt.replace(hour=0, minute=0, second=0, microsecond=0)
            appointment.Start = normalized_start
            appointment.End = normalized_start + datetime.timedelta(days=1)
            appointment.AllDayEvent = True
        else:
            appointment.Start = start_dt
            appointment.AllDayEvent = False
            appointment.Duration = duration_value or 60
            appointment.End = start_dt + datetime.timedelta(minutes=appointment.Duration)

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
        start_display = start_dt.strftime("%Y-%m-%d") if all_day_bool else start_dt.strftime("%Y-%m-%d %H:%M")

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
