import datetime
import types
from pathlib import Path

import outlook_mcp_server as server


def test_present_email_listing_with_offset():
    emails = [
        {
            "id": "m1",
            "subject": "Alpha",
            "sender": "Alice",
            "sender_email": "alice@example.com",
            "received_time": "2025-01-01 09:00",
            "folder_path": "Inbox",
            "unread": False,
        },
        {
            "id": "m2",
            "subject": "Beta",
            "sender": "Bob",
            "sender_email": "bob@example.com",
            "received_time": "2025-01-01 10:00",
            "folder_path": "Inbox",
            "unread": True,
        },
        {
            "id": "m3",
            "subject": "Gamma",
            "sender": "Carol",
            "sender_email": "carol@example.com",
            "received_time": "2025-01-01 11:00",
            "folder_path": "Inbox",
            "unread": True,
        },
    ]

    result = server._present_email_listing(
        emails=emails,
        folder_display="Posta in arrivo",
        days=7,
        max_results=2,
        include_preview=False,
        log_context="test_list",
        offset=1,
    )

    assert "Mostro i risultati 2-3" in result
    assert "Messaggio #1" in result
    assert "Messaggio #2" in result
    assert server.email_cache[1]["id"] == "m2"
    assert server.email_cache[2]["id"] == "m3"
    server.clear_email_cache()


def test_mark_email_read_unread_updates_cache(monkeypatch):
    mail_item = types.SimpleNamespace(UnRead=True, Save=lambda: None)
    server.email_cache = {
        1: {"id": "abc", "unread": True},
    }

    monkeypatch.setattr(server, "connect_to_outlook", lambda: (None, None))

    def fake_resolver(namespace, *, email_number=None, message_id=None):
        assert email_number == 1
        return server.email_cache[email_number], mail_item

    monkeypatch.setattr(server, "_resolve_mail_item", fake_resolver)

    result = server.mark_email_read_unread(email_number=1, unread=False)

    assert "Letta" in result
    assert mail_item.UnRead is False
    assert server.email_cache[1]["unread"] is False
    server.clear_email_cache()


def test_get_attachments_download(tmp_path, monkeypatch):
    class DummyAttachment:
        def __init__(self, name, size):
            self.FileName = name
            self.Size = size

        def SaveAsFile(self, destination: str) -> None:
            Path(destination).write_text("dummy")

    class DummyAttachments:
        def __init__(self, attachments):
            self._attachments = attachments
            self.Count = len(attachments)

        def __call__(self, index: int):
            return self._attachments[index - 1]

    class DummyMail:
        def __init__(self, attachments):
            self.Attachments = DummyAttachments(attachments)

    class DummyNamespace:
        def __init__(self, mail):
            self._mail = mail

        def GetItemFromID(self, message_id):
            self.last_id = message_id
            return self._mail

    attachments = [
        DummyAttachment("report.pdf", 1024),
        DummyAttachment("data.csv", 2048),
    ]
    mail_item = DummyMail(attachments)
    namespace = DummyNamespace(mail_item)

    monkeypatch.setattr(server, "connect_to_outlook", lambda: (None, namespace))

    result = server.get_attachments(
        message_id="MSG1",
        download=True,
        save_to=str(tmp_path),
    )

    assert "report.pdf" in result
    assert "data.csv" in result

    saved_files = sorted(p.name for p in tmp_path.iterdir())
    assert saved_files == ["data.csv", "report.pdf"]


def test_move_email_to_folder_updates_cache(monkeypatch):
    class DummyFolder:
        def __init__(self):
            self.FolderPath = "\\\\Mailbox\\Processed"
            self.Name = "Processed"

    class DummyMail:
        def __init__(self):
            self.EntryID = "OLD"
            self.moved_to = None

        def Move(self, target):
            self.moved_to = target
            return types.SimpleNamespace(EntryID="NEW")

    dummy_folder = DummyFolder()
    mail_item = DummyMail()
    namespace = object()

    server.email_cache = {1: {"id": "OLD"}}

    def fake_connect():
        return None, namespace

    def fake_resolve_mail_item(ns, *, email_number=None, message_id=None):
        assert ns is namespace
        assert email_number == 1
        return server.email_cache[email_number], mail_item

    def fake_resolve_folder(ns, *, folder_id=None, folder_path=None, folder_name=None):
        assert ns is namespace
        assert folder_id == "FOLDER-ID"
        return dummy_folder, []

    monkeypatch.setattr(server, "connect_to_outlook", fake_connect)
    monkeypatch.setattr(server, "_resolve_mail_item", fake_resolve_mail_item)
    monkeypatch.setattr(server.folder_service, "resolve_folder", fake_resolve_folder)

    result = server.move_email_to_folder(target_folder_id="FOLDER-ID", email_number=1)

    assert "Messaggio #1 spostato" in result
    assert mail_item.moved_to is dummy_folder
    assert server.email_cache[1]["id"] == "NEW"
    assert server.email_cache[1]["folder_path"] == dummy_folder.Name
    server.clear_email_cache()


def test_create_folder_delegates_to_service(monkeypatch):
    namespace = object()
    parent_folder = types.SimpleNamespace(Name="Inbox", FolderPath="\\\\Mailbox\\Inbox")
    created_folder = types.SimpleNamespace(
        Name="Reports",
        FolderPath="\\\\Mailbox\\Inbox\\Reports",
        EntryID="ENTRY",
        DefaultItemType=0,
    )
    calls = {}

    def fake_connect():
        return None, namespace

    def fake_resolve(ns, *, folder_id=None, folder_path=None, folder_name=None):
        calls["resolve"] = (folder_id, folder_path, folder_name)
        return parent_folder, []

    def fake_create(parent, **kwargs):
        calls["create_kwargs"] = kwargs
        assert parent is parent_folder
        return created_folder, "Cartella 'Reports' creata con successo."

    monkeypatch.setattr(server, "connect_to_outlook", fake_connect)
    monkeypatch.setattr(server.folder_service, "resolve_folder", fake_resolve)
    monkeypatch.setattr(server.folder_service, "create_folder", fake_create)

    message = server.create_folder("Reports", parent_folder_id="PID")

    assert "Reports" in message
    assert calls["resolve"] == ("PID", None, None)
    assert calls["create_kwargs"]["new_folder_name"] == "Reports"


def test_delete_folder_invokes_service(monkeypatch):
    namespace = object()
    parent = types.SimpleNamespace(FolderPath="\\\\Mailbox\\Inbox")
    target_folder = types.SimpleNamespace(Name="Old", Parent=parent)
    deleted = {}

    def fake_connect():
        return None, namespace

    def fake_resolve(ns, *, folder_id=None, folder_path=None, folder_name=None):
        deleted["resolve_args"] = (folder_id, folder_path, folder_name)
        return target_folder, []

    def fake_delete(folder):
        deleted["folder"] = folder

    monkeypatch.setattr(server, "connect_to_outlook", fake_connect)
    monkeypatch.setattr(server.folder_service, "resolve_folder", fake_resolve)
    monkeypatch.setattr(server.folder_service, "delete_folder", fake_delete)

    result = server.delete_folder(folder_id="DEL-ID", confirm=True)

    assert "Cartella eliminata" in result
    assert deleted["resolve_args"] == ("DEL-ID", None, None)
    assert deleted["folder"] is target_folder


def test_compose_email_html_body(monkeypatch):
    class DummyAttachments:
        def Add(self, path):
            raise AssertionError("Non dovrebbero essere aggiunti allegati in questo test.")

    class DummyMail:
        def __init__(self):
            self.Attachments = DummyAttachments()
            self.HTMLBody = ""
            self.Body = ""
            self.sent = False

        def Send(self):
            self.sent = True

        def Save(self):
            self.saved = True

    class DummyOutlook:
        def __init__(self):
            self.created = None

        def CreateItem(self, item_type):
            assert item_type == 0
            self.created = DummyMail()
            return self.created

    dummy_outlook = DummyOutlook()
    monkeypatch.setattr(server, "connect_to_outlook", lambda: (dummy_outlook, None))

    response = server.compose_email(
        recipient_email="dest@example.com",
        subject="HTML Test",
        body="<h1>Hello</h1>",
        use_html=True,
    )

    assert response == "Email inviata a: dest@example.com"
    assert dummy_outlook.created.HTMLBody == "<h1>Hello</h1>"
    assert dummy_outlook.created.sent is True
    assert dummy_outlook.created.Body == ""


def test_search_contacts_filters_results(monkeypatch):
    class DummyItems:
        def __init__(self, contacts):
            self._contacts = contacts
            self.Count = len(contacts)

        def __call__(self, index):
            return self._contacts[index - 1]

    class DummyFolder:
        def __init__(self, items):
            self.Items = items

    class DummyNamespace:
        def __init__(self, folder):
            self._folder = folder

        def GetDefaultFolder(self, folder_id):
            assert folder_id == 10
            return self._folder

    contacts = [
        types.SimpleNamespace(
            FullName="Alice Rossi",
            Email1Address="alice@example.com",
            CompanyName="Acme S.p.A.",
            MobileTelephoneNumber="123",
            Categories="Clienti",
        ),
        types.SimpleNamespace(
            FullName="Bob Bianchi",
            Email1Address="bob@example.com",
            CompanyName="BetaCorp",
            MobileTelephoneNumber="456",
            Categories="Partner",
        ),
    ]

    items = DummyItems(contacts)
    namespace = DummyNamespace(DummyFolder(items))

    monkeypatch.setattr(server, "connect_to_outlook", lambda: (None, namespace))

    full_listing = server.search_contacts()
    assert "Alice Rossi" in full_listing
    assert "Bob Bianchi" in full_listing

    filtered_listing = server.search_contacts(search_term="Acme")
    assert "Alice Rossi" in filtered_listing
    assert "Bob Bianchi" not in filtered_listing


def test_search_contacts_iter_fallback_list(monkeypatch):
    class DummyItems:
        def __init__(self, contacts):
            self._items = contacts

    class DummyNamespace:
        def __init__(self, contacts):
            self.folder = types.SimpleNamespace(Items=DummyItems(contacts))

        def GetDefaultFolder(self, folder_id):
            assert folder_id == 10
            return self.folder

    contacts = [
        types.SimpleNamespace(FullName="Mario Rossi", Email1Address="mario@example.com"),
    ]

    namespace = DummyNamespace(contacts)
    monkeypatch.setattr(server, "connect_to_outlook", lambda: (None, namespace))

    result = server.search_contacts()
    assert "Mario Rossi" in result


def test_search_contacts_iter_fallback_getnext(monkeypatch):
    class DummyItems:
        def __init__(self, contacts):
            self._contacts = contacts
            self._index = 0

        def GetFirst(self):
            self._index = 0
            return self._contacts[self._index] if self._contacts else None

        def GetNext(self):
            self._index += 1
            if self._index >= len(self._contacts):
                return None
            return self._contacts[self._index]

    class DummyNamespace:
        def __init__(self, contacts):
            self.folder = types.SimpleNamespace(Items=DummyItems(contacts))

        def GetDefaultFolder(self, folder_id):
            assert folder_id == 10
            return self.folder

    contacts = [
        types.SimpleNamespace(FullName="Lucia Verdi", Email1Address="lucia@example.com"),
    ]

    namespace = DummyNamespace(contacts)
    monkeypatch.setattr(server, "connect_to_outlook", lambda: (None, namespace))

    result = server.search_contacts()
    assert "Lucia Verdi" in result


def test_create_calendar_event_basic(monkeypatch):
    class DummyRecipients:
        def __init__(self):
            self.addresses = []

        def Add(self, address):
            self.addresses.append(address)
            recipient = types.SimpleNamespace(Type=1)
            return recipient

    class DummyAppointment:
        def __init__(self):
            self.Recipients = DummyRecipients()
            self.EntryID = "EVT-1"
            self.ReminderSet = False
            self.ReminderMinutesBeforeStart = None
            self.Duration = None
            self.Start = None
            self.End = None
            self.AllDayEvent = False
            self.sent = False
            self.saved = False

        def Save(self):
            self.saved = True

        def Send(self):
            self.sent = True

        def Move(self, folder):
            self.moved_to = folder
            return self

    class DummyOutlook:
        def CreateItem(self, item_type):
            assert item_type == 1
            self.created = DummyAppointment()
            return self.created

    class DummyNamespace:
        def __init__(self):
            self.default_calendar = types.SimpleNamespace(Name="Calendario")

        def GetDefaultFolder(self, folder_id):
            assert folder_id == 9
            return self.default_calendar

    dummy_outlook = DummyOutlook()
    dummy_namespace = DummyNamespace()

    monkeypatch.setattr(server, "connect_to_outlook", lambda: (dummy_outlook, dummy_namespace))

    response = server.create_calendar_event(
        subject="Riunione settimanale ",
        start_time="2025-10-20 10:30",
        duration_minutes=90,
        location="Sala Riunioni",
        body="Agenda:\n- Stato progetti",
        attendees=["alice@example.com", "bob@example.com"],
        reminder_minutes=30,
        send_invitations=True,
    )

    appointment = dummy_outlook.created
    assert appointment.Subject == "Riunione settimanale"
    assert appointment.Start == datetime.datetime(2025, 10, 20, 10, 30)
    assert appointment.Duration == 90
    assert appointment.Location == "Sala Riunioni"
    assert appointment.Body.startswith("Agenda")
    assert appointment.ReminderSet is True
    assert appointment.ReminderMinutesBeforeStart == 30
    assert appointment.saved is True
    assert appointment.sent is True
    assert appointment.Recipients.addresses == ["alice@example.com", "bob@example.com"]
    assert "Evento 'Riunione settimanale' creato" in response
    assert "Inviti inviati" in response


def test_create_calendar_event_all_day(monkeypatch):
    class DummyAppointment:
        def __init__(self):
            self.EntryID = "EVT-2"
            self.ReminderSet = False
            self.Start = None
            self.End = None
            self.AllDayEvent = False

        def Save(self):
            self.saved = True

        def Move(self, folder):
            return self

    class DummyOutlook:
        def CreateItem(self, item_type):
            assert item_type == 1
            self.created = DummyAppointment()
            return self.created

    class DummyNamespace:
        def __init__(self):
            self.default_calendar = types.SimpleNamespace(Name="Calendario")

        def GetDefaultFolder(self, folder_id):
            assert folder_id == 9
            return self.default_calendar

    dummy_outlook = DummyOutlook()
    dummy_namespace = DummyNamespace()
    monkeypatch.setattr(server, "connect_to_outlook", lambda: (dummy_outlook, dummy_namespace))

    response = server.create_calendar_event(
        subject="Giornata intera",
        start_time="2025-12-24 15:30",
        all_day=True,
    )

    appointment = dummy_outlook.created
    assert appointment.AllDayEvent is True
    assert appointment.Start == datetime.datetime(2025, 12, 24, 0, 0)
    assert appointment.End == datetime.datetime(2025, 12, 25, 0, 0)
    assert "Evento 'Giornata intera' creato" in response


def test_create_calendar_event_invalid_start():
    result = server.create_calendar_event(subject="Evento", start_time="non-data")
    assert "Errore" in result
    assert "start_time" in result


def test_create_calendar_event_calendar_not_found(monkeypatch):
    dummy_outlook = types.SimpleNamespace()
    dummy_namespace = types.SimpleNamespace()

    monkeypatch.setattr(server, "connect_to_outlook", lambda: (dummy_outlook, dummy_namespace))
    monkeypatch.setattr(server, "get_calendar_folder_by_name", lambda ns, name: None)

    result = server.create_calendar_event(
        subject="Evento",
        start_time="2025-11-01",
        calendar_name="Agenda condivisa",
    )

    assert "calendario 'agenda condivisa' non trovato" in result.lower()
