import datetime
import unittest
from unittest.mock import MagicMock, patch

import outlook_mcp_server as server


class HelperFunctionsTest(unittest.TestCase):
    def test_normalize_email_address_handles_wrapped_smtp(self):
        raw = "Display Name <SMTP:USER@Example.com>"
        self.assertEqual(server._normalize_email_address(raw), "user@example.com")

    def test_parse_datetime_string_accepts_multiple_formats(self):
        iso_input = "2025-10-08T12:30:15"
        legacy_input = "2025-10-08 12:30"
        self.assertEqual(
            server._parse_datetime_string(iso_input),
            datetime.datetime(2025, 10, 8, 12, 30, 15),
        )
        self.assertEqual(
            server._parse_datetime_string(legacy_input),
            datetime.datetime(2025, 10, 8, 12, 30),
        )

    def test_extract_best_timestamp_prefers_received_iso(self):
        entry = {
            "received_iso": "2025-10-08T09:00:00",
            "received_time": "2025-10-08 08:30:00",
        }
        extracted = server._extract_best_timestamp(entry)
        self.assertEqual(extracted, datetime.datetime(2025, 10, 8, 9, 0))

    def test_coerce_bool_handles_strings(self):
        self.assertTrue(server._coerce_bool("true"))
        self.assertFalse(server._coerce_bool("off"))

    def test_mail_item_marked_replied_lastverb(self):
        baseline = datetime.datetime(2025, 10, 8, 9, 0, 0)

        class DummyMail:
            LastVerbExecuted = 102
            LastVerbExecutionTime = datetime.datetime(2025, 10, 8, 9, 5, 0)

        self.assertTrue(server._mail_item_marked_replied(DummyMail(), baseline))

    def test_mail_item_marked_replied_property_accessor(self):
        baseline = datetime.datetime(2025, 10, 8, 9, 0, 0)

        class DummyAccessor:
            def GetProperty(self, name):
                if name == server.PR_LAST_VERB_EXECUTED:
                    return 103
                if name == server.PR_LAST_VERB_EXECUTION_TIME:
                    return datetime.datetime(2025, 10, 8, 10, 0, 0)
                raise AssertionError("Unexpected property")

        class DummyMail:
            LastVerbExecuted = None
            LastVerbExecutionTime = None
            PropertyAccessor = DummyAccessor()

        self.assertTrue(server._mail_item_marked_replied(DummyMail(), baseline))


class GetCurrentDatetimeTest(unittest.TestCase):
    def test_get_current_datetime_includes_local_and_utc(self):
        class FixedDateTime(datetime.datetime):
            @classmethod
            def now(cls):
                return cls(2025, 10, 9, 10, 30, 45)

            @classmethod
            def utcnow(cls):
                return cls(2025, 10, 9, 8, 30, 45)

        with patch("outlook_mcp_server.datetime.datetime", FixedDateTime):
            result = server.get_current_datetime(include_utc=True)

        self.assertIn("2025-10-09 10:30:45", result)
        self.assertIn("2025-10-09T10:30:45", result)
        self.assertIn("2025-10-09 08:30:45", result)
        self.assertIn("UTC ISO", result)

    def test_get_current_datetime_can_skip_utc(self):
        class FixedDateTime(datetime.datetime):
            @classmethod
            def now(cls):
                return cls(2025, 10, 9, 10, 30, 45)

            @classmethod
            def utcnow(cls):
                return cls(2025, 10, 9, 8, 30, 45)

        with patch("outlook_mcp_server.datetime.datetime", FixedDateTime):
            result = server.get_current_datetime(include_utc=False)

        self.assertIn("2025-10-09 10:30:45", result)
        self.assertNotIn("UTC:", result)


class EmailReplyInferenceTest(unittest.TestCase):
    def setUp(self):
        server.email_cache.clear()

    @patch("outlook_mcp_server._mail_item_marked_replied", return_value=False)
    @patch("outlook_mcp_server.get_related_conversation_emails")
    def test_email_has_user_reply_detects_related_sender(self, mock_related, mock_marked):
        namespace = MagicMock()
        namespace.GetItemFromID.return_value = object()

        email_data = {
            "id": "abc",
            "received_iso": "2025-10-08T09:00:00",
        }
        related_entry = {
            "sender_email": "user@example.com",
            "received_iso": "2025-10-08T10:00:00",
        }
        mock_related.return_value = [related_entry]

        result = server._email_has_user_reply(
            namespace=namespace,
            email_data=email_data,
            user_addresses={"user@example.com"},
            conversation_limit=5,
            lookback_days=30,
        )

        self.assertTrue(result)
        namespace.GetItemFromID.assert_called_once_with("abc")
        mock_related.assert_called_once()

    @patch("outlook_mcp_server._mail_item_marked_replied", return_value=True)
    @patch("outlook_mcp_server.get_related_conversation_emails", return_value=[])
    def test_email_has_user_reply_uses_lastverb(self, mock_related, mock_marked):
        namespace = MagicMock()
        namespace.GetItemFromID.return_value = MagicMock()
        email_data = {"id": "xyz", "received_iso": "2025-10-08T09:00:00"}

        result = server._email_has_user_reply(
            namespace=namespace,
            email_data=email_data,
            user_addresses=set(),
            conversation_limit=5,
            lookback_days=30,
        )
        self.assertTrue(result)
        mock_related.assert_not_called()
        mock_marked.assert_called_once()


class ListPendingRepliesTest(unittest.TestCase):
    def setUp(self):
        server.email_cache.clear()

    @patch("outlook_mcp_server._email_has_user_reply", side_effect=[False, True])
    @patch("outlook_mcp_server.get_emails_from_folder")
    @patch("outlook_mcp_server._collect_user_addresses", return_value={"user@example.com"})
    @patch("outlook_mcp_server.connect_to_outlook")
    def test_list_pending_replies_filters_and_formats(
        self,
        mock_connect,
        mock_collect_addresses,
        mock_get_emails,
        mock_has_reply,
    ):
        namespace = MagicMock()
        mock_connect.return_value = (None, namespace)
        namespace.GetDefaultFolder.return_value = MagicMock()

        candidate_emails = [
            {
                "id": "1",
                "subject": "Follow up",
                "sender": "Partner",
                "sender_email": "partner@example.com",
                "unread": True,
                "received_time": "2025-10-08 09:00:00",
                "importance": 1,
                "folder_path": "Inbox",
            },
            {
                "id": "2",
                "subject": "Done",
                "sender": "User Self",
                "sender_email": "user@example.com",
                "unread": True,
                "received_time": "2025-10-08 08:00:00",
                "importance": 1,
                "folder_path": "Inbox",
            },
        ]
        mock_get_emails.return_value = candidate_emails

        result = server.list_pending_replies(
            days=7,
            max_results=5,
            include_preview=False,
            include_all_folders=False,
            include_unread_only=True,
        )

        self.assertIn("Messaggio #1", result)
        self.assertIn("Follow up", result)
        self.assertNotIn("Messaggio #2", result)
        mock_has_reply.assert_called()
        self.assertEqual(len(server.email_cache), 1)

    def test_list_pending_replies_validates_days(self):
        message = server.list_pending_replies(days=31)
        self.assertTrue(message.startswith("Errore"))


class CalendarToolsTest(unittest.TestCase):
    def setUp(self):
        server.calendar_cache.clear()

    @patch("outlook_mcp_server.get_events_from_folder")
    @patch("outlook_mcp_server.connect_to_outlook")
    def test_list_upcoming_events_default_calendar(self, mock_connect, mock_get_events):
        namespace = MagicMock()
        calendar_folder = MagicMock()
        calendar_folder.Name = "Calendario"
        namespace.GetDefaultFolder.return_value = calendar_folder
        mock_connect.return_value = (None, namespace)

        mock_get_events.return_value = [
            {
                "subject": "Riunione team",
                "start_time": "2025-10-10 09:00",
                "end_time": "2025-10-10 10:00",
                "folder_path": "\\Calendario",
            }
        ]

        output = server.list_upcoming_events(days=3, max_results=5, include_description=False)

        self.assertIn("Evento #1", output)
        self.assertIn("Riunione team", output)
        self.assertEqual(len(server.calendar_cache), 1)
        namespace.GetDefaultFolder.assert_called_once()
        mock_get_events.assert_called_once_with(calendar_folder, 3)

    @patch("outlook_mcp_server.collect_events_across_calendars")
    @patch("outlook_mcp_server.get_all_calendar_folders")
    @patch("outlook_mcp_server.connect_to_outlook")
    def test_list_upcoming_events_include_all(self, mock_connect, mock_get_all, mock_collect):
        namespace = MagicMock()
        mock_connect.return_value = (None, namespace)
        mock_get_all.return_value = ["calA", "calB"]
        mock_collect.return_value = [
            {
                "subject": "All Hands",
                "start_time": "2025-10-11 15:00",
                "end_time": "2025-10-11 16:00",
                "folder_path": "\\All",
            }
        ]

        output = server.list_upcoming_events(
            days=7,
            max_results=3,
            include_description=True,
            include_all_calendars=True,
        )

        self.assertIn("Tutti i calendari", output)
        self.assertIn("All Hands", output)
        mock_get_all.assert_called_once_with(namespace)
        mock_collect.assert_called_once_with(["calA", "calB"], 7)

    def test_list_upcoming_events_validates_days(self):
        too_many_days = server.MAX_EVENT_LOOKAHEAD_DAYS + 1
        message = server.list_upcoming_events(days=too_many_days)
        self.assertTrue(message.startswith("Errore"))

    @patch("outlook_mcp_server.get_events_from_folder")
    @patch("outlook_mcp_server.get_calendar_folder_by_name")
    @patch("outlook_mcp_server.connect_to_outlook")
    def test_search_calendar_events_single_calendar(
        self,
        mock_connect,
        mock_get_calendar,
        mock_get_events,
    ):
        namespace = MagicMock()
        mock_connect.return_value = (None, namespace)
        calendar_folder = MagicMock()
        calendar_folder.Name = "Progetti"
        mock_get_calendar.return_value = calendar_folder

        mock_get_events.return_value = [
            {"subject": "Workshop", "start_time": "2025-10-12 11:00", "end_time": "2025-10-12 12:00"}
        ]

        output = server.search_calendar_events(
            search_term="Workshop",
            days=5,
            calendar_name="Progetti",
            include_description=False,
        )

        self.assertIn("Workshop", output)
        mock_get_calendar.assert_called_once_with(namespace, "Progetti")
        mock_get_events.assert_called_once_with(calendar_folder, 5, "Workshop")

    @patch("outlook_mcp_server.collect_events_across_calendars")
    @patch("outlook_mcp_server.get_all_calendar_folders")
    @patch("outlook_mcp_server.connect_to_outlook")
    def test_search_calendar_events_include_all(
        self,
        mock_connect,
        mock_get_all,
        mock_collect,
    ):
        namespace = MagicMock()
        mock_connect.return_value = (None, namespace)
        mock_get_all.return_value = ["cal1"]
        mock_collect.return_value = [{"subject": "Demo", "start_time": "2025-10-13 14:00", "end_time": "2025-10-13 15:00"}]

        output = server.search_calendar_events(
            search_term="Demo",
            days=10,
            include_all_calendars=True,
        )

        self.assertIn("Demo", output)
        mock_get_all.assert_called_once_with(namespace)
        mock_collect.assert_called_once_with(["cal1"], 10, "Demo")

    def test_search_calendar_events_requires_term(self):
        message = server.search_calendar_events(search_term="")
        self.assertTrue(message.startswith("Errore"))

    def test_get_event_by_number_uses_cache(self):
        server.calendar_cache.clear()
        server.calendar_cache[1] = {
            "subject": "Demo giornaliero",
            "start_time": "2025-10-10 09:00",
            "end_time": "2025-10-10 10:00",
            "folder_path": "\\Calendario",
            "location": "Sala A",
            "organizer": "Davide",
            "all_day": False,
        }
        details = server.get_event_by_number(1)
        self.assertIn("Demo giornaliero", details)
        self.assertIn("Sala A", details)

    def test_get_event_by_number_validates_cache(self):
        server.calendar_cache.clear()
        message = server.get_event_by_number(1)
        self.assertTrue(message.startswith("Errore"))


if __name__ == "__main__":
    unittest.main()
