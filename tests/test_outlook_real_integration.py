import datetime
import datetime
import os
import re
from pathlib import Path

import pytest


def _is_real_outlook_enabled() -> bool:
    value = os.environ.get("OUTLOOK_MCP_REAL", "")
    return value.strip().lower() in {"1", "true", "yes", "on"}


if not _is_real_outlook_enabled():
    pytest.skip(
        "Outlook integration tests disabilitati. Imposta OUTLOOK_MCP_REAL=1 per eseguirli con Outlook reale.",
        allow_module_level=True,
    )


import outlook_mcp_server as server  # noqa: E402  (import dopo skip condizionale)


def _extract_message_id(text: str) -> str | None:
    match = re.search(r"message_id=([^\)\s]+)", text)
    return match.group(1) if match else None


def _extract_event_id(text: str) -> str | None:
    match = re.search(r"Message ID:\s*([^\s]+)", text)
    return match.group(1) if match else None


def _get_self_email(namespace) -> str:
    try:
        current_user = getattr(namespace, "CurrentUser", None)
        if current_user:
            address = getattr(current_user, "Address", None)
            if address and "@" in address:
                return address
            address_entry = getattr(current_user, "AddressEntry", None)
            if address_entry:
                get_exchange = getattr(address_entry, "GetExchangeUser", None)
                if callable(get_exchange):
                    exchange_user = get_exchange()
                    primary = getattr(exchange_user, "PrimarySmtpAddress", None)
                    if primary:
                        return primary
    except Exception:
        pass

    try:
        accounts = getattr(namespace, "Accounts", None)
        if accounts:
            first_account = accounts.Item(1)
            smtp_address = getattr(first_account, "SmtpAddress", None)
            if smtp_address:
                return smtp_address
    except Exception:
        pass

    raise RuntimeError("Impossibile determinare l'indirizzo email dell'account corrente.")


@pytest.mark.slow
def test_real_outlook_end_to_end(tmp_path: Path):
    unique_token = datetime.datetime.now().strftime("MCPINT_%Y%m%d_%H%M%S")
    base_folder_name = f"{unique_token}_Base"
    renamed_folder_name = f"{base_folder_name}_Renamed"
    move_folder_name = f"{unique_token}_Move"
    contact_name = f"{unique_token} Contact"
    contact_email = f"{unique_token.lower()}@example.com"
    event_subject = f"{unique_token} Meeting"
    html_body = "<h1>Integrazione MCP</h1><p>Messaggio di test.</p>"

    attachment_source = tmp_path / "mcp_integration_attachment.txt"
    attachment_source.write_text("Contenuto allegato MCP integration test", encoding="utf-8")
    attachment_download_dir = tmp_path / "downloads"
    attachment_download_dir.mkdir()

    message_id: str | None = None
    contact_item = None
    event_id: str | None = None

    outlook, namespace = server.connect_to_outlook()
    self_email = _get_self_email(namespace)
    inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
    drafts = namespace.GetDefaultFolder(16)  # olFolderDrafts
    drafts_path = drafts.FolderPath

    created_folders = set()

    try:
        create_resp = server.create_folder(
            new_folder_name=base_folder_name,
            parent_folder_id=inbox.EntryID,
            allow_existing=False,
        )
        assert base_folder_name in create_resp
        created_folders.add(base_folder_name)

        rename_resp = server.rename_folder(folder_name=base_folder_name, new_name=renamed_folder_name)
        assert "Cartella rinominata" in rename_resp
        created_folders.add(renamed_folder_name)

        move_folder_resp = server.create_folder(
            new_folder_name=move_folder_name,
            parent_folder_id=inbox.EntryID,
            allow_existing=True,
        )
        assert move_folder_name in move_folder_resp
        created_folders.add(move_folder_name)

        move_folder_obj = server.get_folder_by_name(namespace, move_folder_name)
        assert move_folder_obj is not None
        move_folder_path = move_folder_obj.FolderPath

        compose_resp = server.compose_email(
            recipient_email=self_email,
            subject=f"{unique_token} Subject",
            body=html_body,
            send=False,
            use_html=True,
        )
        message_id = _extract_message_id(compose_resp)
        assert message_id is not None

        mail_item = namespace.GetItemFromID(message_id)
        assert html_body in getattr(mail_item, "HTMLBody", "")

        move_resp = server.move_email_to_folder(message_id=message_id, target_folder_path=move_folder_path)
        message_id = _extract_message_id(move_resp)
        assert message_id is not None

        move_back_resp = server.move_email_to_folder(message_id=message_id, target_folder_path=drafts_path)
        message_id = _extract_message_id(move_back_resp)
        assert message_id is not None

        attach_resp = server.attach_to_email(
            attachments=[str(attachment_source)],
            message_id=message_id,
            send=False,
        )
        assert "Allegati aggiunti" in attach_resp

        download_resp = server.get_attachments(
            message_id=message_id,
            download=True,
            save_to=str(attachment_download_dir),
        )
        assert "Allegati salvati" in download_resp
        assert (attachment_download_dir / attachment_source.name).exists()

        contact_item = outlook.CreateItem(2)  # olContactItem
        contact_item.FullName = contact_name
        contact_item.Email1Address = contact_email
        contact_item.Save()

        contacts_resp = server.search_contacts(search_term=unique_token)
        assert contact_name in contacts_resp
        assert contact_email in contacts_resp

        start_dt = (datetime.datetime.now() + datetime.timedelta(hours=1)).replace(minute=0, second=0, microsecond=0)
        event_resp = server.create_calendar_event(
            subject=event_subject,
            start_time=start_dt.strftime("%Y-%m-%d %H:%M"),
            duration_minutes=45,
            location="Sala Test MCP",
            body="Evento di integrazione MCP.",
            reminder_minutes=10,
            send_invitations=False,
        )
        assert event_subject in event_resp
        event_id = _extract_event_id(event_resp)
        assert event_id is not None

    finally:
        if event_id:
            try:
                event_item = namespace.GetItemFromID(event_id)
                event_item.Delete()
            except Exception:
                pass

        if contact_item:
            try:
                contact_item.Delete()
            except Exception:
                pass

        if message_id:
            try:
                server.batch_manage_emails(message_ids=[message_id], delete=True)
            except Exception:
                pass

        for name in created_folders:
            try:
                server.delete_folder(folder_name=name, confirm=True)
            except Exception:
                pass
