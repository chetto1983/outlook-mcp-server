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
