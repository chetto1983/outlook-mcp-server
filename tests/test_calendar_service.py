import datetime
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from outlook_mcp.services import calendar as calendar_service


class FixedDateTime(datetime.datetime):
    """Frozen datetime to provide deterministic 'now' in tests."""

    @classmethod
    def now(cls, tz=None):
        return cls(2025, 10, 18, 10, 0, 0, tzinfo=tz)


class MockAppointment:
    def __init__(
        self,
        subject,
        start,
        end,
        *,
        entry_id=None,
        location="",
        organizer="Test Organizer",
        required="",
        optional="",
        all_day=False,
        is_recurring=False,
        body="",
        categories="",
        folder_path="\\\\Calendario",
    ):
        self.Subject = subject
        self.Start = start
        self.End = end
        self.EntryID = entry_id or subject
        self.Location = location
        self.Organizer = organizer
        self.RequiredAttendees = required
        self.OptionalAttendees = optional
        self.AllDayEvent = all_day
        self.IsRecurring = is_recurring
        self.Body = body
        self.Categories = categories
        self.Parent = SimpleNamespace(FolderPath=folder_path)


def _to_naive_datetime(value):
    if isinstance(value, datetime.datetime) and value.tzinfo:
        try:
            return value.astimezone().replace(tzinfo=None)
        except Exception:
            return value.replace(tzinfo=None)
    return value


class MockItems:
    def __init__(self, appointments):
        self._appointments = list(appointments)
        self.IncludeRecurrences = False
        self.last_sort_key = None
        self.last_restriction = None
        self.last_find_filter = None
        self._find_results = []
        self._find_index = -1

    def Sort(self, key):
        self.last_sort_key = key
        attribute = key.strip("[]")
        self._appointments.sort(key=lambda appt: getattr(appt, attribute))

    def Restrict(self, restriction):
        self.last_restriction = restriction
        start_match = None
        end_match = None
        if restriction:
            start_match = _extract_datetime(restriction, r"\[End\] >= '([^']+)'")
            end_match = _extract_datetime(restriction, r"\[Start\] <= '([^']+)'")

        filtered = self._appointments
        if start_match and end_match:
            filtered = [
                appt
                for appt in self._appointments
                if _to_naive_datetime(appt.End) >= start_match
                and _to_naive_datetime(appt.Start) <= end_match
            ]
        restricted_items = MockItems(filtered)
        restricted_items.IncludeRecurrences = self.IncludeRecurrences
        restricted_items.last_restriction = restriction
        return restricted_items

    def Find(self, filter_query):
        self.last_find_filter = filter_query
        start_match = _extract_datetime(filter_query, r"\[End\] >= '([^']+)'")
        if start_match:
            self._find_results = [
                appt for appt in self._appointments if _to_naive_datetime(appt.End) >= start_match
            ]
        else:
            self._find_results = list(self._appointments)
        self._find_index = 0
        if self._find_results:
            return self._find_results[0]
        return None

    def FindNext(self):
        if not self._find_results:
            return None
        self._find_index += 1
        if self._find_index < len(self._find_results):
            return self._find_results[self._find_index]
        return None

    def __iter__(self):
        return iter(self._appointments)


class MockFolder:
    def __init__(self, name, appointments):
        self.Name = name
        self.Items = MockItems(appointments)


def _extract_datetime(restriction, pattern):
    import re

    match = re.search(pattern, restriction)
    if not match:
        return None
    return datetime.datetime.strptime(match.group(1), "%m/%d/%Y %I:%M %p")


def _freeze_now(monkeypatch):
    monkeypatch.setattr(calendar_service.datetime, "datetime", FixedDateTime)


def test_get_events_from_folder_filters_time_window(monkeypatch):
    _freeze_now(monkeypatch)

    now = FixedDateTime.now()
    upcoming = MockAppointment(
        "Evento imminente",
        start=now + datetime.timedelta(hours=2),
        end=now + datetime.timedelta(hours=3),
    )
    distant = MockAppointment(
        "Evento lontano",
        start=now + datetime.timedelta(days=20),
        end=now + datetime.timedelta(days=20, hours=1),
    )
    past = MockAppointment(
        "Evento passato",
        start=now - datetime.timedelta(days=2),
        end=now - datetime.timedelta(days=1, hours=1),
    )

    folder = MockFolder("Calendario", [past, upcoming, distant])
    events = calendar_service.get_events_from_folder(folder, days=14)

    assert [event["subject"] for event in events] == ["Evento imminente"]

    expected_start = (now - datetime.timedelta(days=1)).strftime("%m/%d/%Y %I:%M %p")
    expected_filter = f"[End] >= '{expected_start}'"
    assert folder.Items.last_find_filter == expected_filter
    assert folder.Items.IncludeRecurrences is True


def test_get_events_from_folder_filters_search_term(monkeypatch):
    _freeze_now(monkeypatch)

    now = FixedDateTime.now()
    matching = MockAppointment(
        "Quarterly meeting",
        start=now + datetime.timedelta(days=1),
        end=now + datetime.timedelta(days=1, hours=1),
        body="Discussione strategica",
    )
    non_matching = MockAppointment(
        "Aggiornamento",
        start=now + datetime.timedelta(days=1),
        end=now + datetime.timedelta(days=1, hours=1),
        body="Note varie",
    )

    folder = MockFolder("Calendario", [matching, non_matching])
    events = calendar_service.get_events_from_folder(folder, days=14, search_term="meeting")

    assert [event["subject"] for event in events] == ["Quarterly meeting"]
    assert folder.Items.last_find_filter is not None


def test_get_events_from_folder_handles_timezone_aware(monkeypatch):
    _freeze_now(monkeypatch)

    aware_now = FixedDateTime.now(datetime.timezone.utc)
    start = aware_now + datetime.timedelta(hours=1)
    end = start + datetime.timedelta(hours=1)
    appointment = MockAppointment(
        "Evento con timezone",
        start=start,
        end=end,
    )

    folder = MockFolder("Calendario", [appointment])
    events = calendar_service.get_events_from_folder(folder, days=1)

    assert [event["subject"] for event in events] == ["Evento con timezone"]


def test_collect_events_across_calendars_deduplicates(monkeypatch):
    event_a = {"id": "A", "start_iso": "2025-10-18T10:00"}
    event_b = {"id": "B", "start_iso": "2025-10-19T09:00"}

    folder_one = SimpleNamespace(events=[event_a, event_b])
    folder_two = SimpleNamespace(events=[{"id": "A", "start_iso": "2025-10-18T10:00"}])

    def fake_get_events(folder, days, search_term):
        return folder.events

    monkeypatch.setattr(calendar_service, "get_events_from_folder", fake_get_events)

    aggregated = calendar_service.collect_events_across_calendars([folder_one, folder_two], days=14)

    assert aggregated == [event_a, event_b]
