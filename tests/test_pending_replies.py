import sys
import types

import pytest


# Provide a lightweight stub for win32com when tests run outside Windows.
win32com_module = types.ModuleType("win32com")
win32com_client_module = types.ModuleType("win32com.client")
win32com_client_module.Dispatch = lambda *args, **kwargs: None
win32com_module.client = win32com_client_module
sys.modules.setdefault("win32com", win32com_module)
sys.modules.setdefault("win32com.client", win32com_client_module)

# Provide minimal stubs for FastMCP so the module can be imported without the
# actual dependency during unit tests.
mcp_module = types.ModuleType("mcp")
mcp_server_module = types.ModuleType("mcp.server")
fastmcp_module = types.ModuleType("mcp.server.fastmcp")


class DummyFastMCP:
    def __init__(self, *args, **kwargs):
        self._tool_manager = types.SimpleNamespace(list_tools=lambda: [])
        self.settings = types.SimpleNamespace()

    def run(self, *args, **kwargs):
        return None

    def tool(self, *args, **kwargs):
        def decorator(func):
            return func

        return decorator


fastmcp_module.FastMCP = DummyFastMCP
fastmcp_module.Context = object

exceptions_module = types.ModuleType("mcp.server.fastmcp.exceptions")


class DummyToolError(Exception):
    pass


exceptions_module.ToolError = DummyToolError

mcp_module.server = types.SimpleNamespace(fastmcp=fastmcp_module)
mcp_server_module.fastmcp = fastmcp_module

sys.modules.setdefault("mcp", mcp_module)
sys.modules.setdefault("mcp.server", mcp_server_module)
sys.modules.setdefault("mcp.server.fastmcp", fastmcp_module)
sys.modules.setdefault("mcp.server.fastmcp.exceptions", exceptions_module)

import outlook_mcp_server as server


class DummyFolder:
    def __init__(self, name: str):
        self.Name = name


class DummyNamespace:
    def __init__(self, folder: DummyFolder):
        self._folder = folder

    def GetDefaultFolder(self, folder_id: int):
        # 6 is Inbox; we only support that in the dummy
        assert folder_id == 6
        return self._folder


@pytest.fixture
def sample_emails():
    base = {
        "conversation_id": "conv-1",
        "received_iso": "2025-10-10T09:00:00",
        "folder_path": "\\\\account\\Posta in arrivo",
        "importance": 1,
        "importance_label": "Normale",
        "has_attachments": False,
        "attachment_names": [],
        "preview": "Messaggio di prova",
        "categories": "",
    }
    read_mail = {
        **base,
        "id": "mail-read",
        "subject": "Richiesta conferma ordine",
        "sender": "Cliente A",
        "sender_email": "cliente@example.com",
        "received_time": "2025-10-10 09:00",
        "unread": False,
    }
    unread_mail = {
        **base,
        "id": "mail-unread",
        "subject": "Nuova richiesta preventivo",
        "sender": "Partner B",
        "sender_email": "partner@example.com",
        "received_time": "2025-10-11 11:30",
        "unread": True,
    }
    return [read_mail, unread_mail]


def test_list_pending_replies_includes_read_messages(monkeypatch, sample_emails):
    inbox_folder = DummyFolder("Posta in arrivo")
    namespace = DummyNamespace(inbox_folder)

    monkeypatch.setattr(server, "connect_to_outlook", lambda: (None, namespace))
    monkeypatch.setattr(
        server,
        "_collect_user_addresses",
        lambda ns: {"user@example.com"},
    )

    def fake_get_emails(folder, days, search_term=None):
        assert folder == inbox_folder
        return [dict(email) for email in sample_emails]

    monkeypatch.setattr(server, "get_emails_from_folder", fake_get_emails)

    monkeypatch.setattr(
        server,
        "_email_has_user_reply_with_context",
        lambda namespace, email_data, user_addresses, conversation_limit, lookback_days, collect_related: (False, [], None),
    )
    monkeypatch.setattr(
        server,
        "_build_conversation_outline",
        lambda namespace, email_data, lookback_days, max_items, preloaded_entries=None, mail_item=None: None,
    )

    result = server.list_pending_replies(days=14, include_all_folders=False)

    assert "Stato lettura: Letta" in result
    assert "Stato lettura: Non letta" in result
