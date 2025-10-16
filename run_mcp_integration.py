import re
import sys
import builtins
from pathlib import Path

import outlook_mcp_server as server


def _safe_print(*args, **kwargs):
    try:
        builtins.print(*args, **kwargs)
    except UnicodeEncodeError:
        encoding = sys.stdout.encoding or "utf-8"
        fallback_args = []
        for value in args:
            if isinstance(value, str):
                fallback_args.append(value.encode(encoding, errors="replace").decode(encoding, errors="replace"))
            else:
                fallback_args.append(value)
        builtins.print(*fallback_args, **kwargs)


print = _safe_print  # type: ignore


def section(title: str) -> None:
    print("\n" + "=" * 80)
    print(title)
    print("=" * 80 + "\n")


def extract_message_id(text: str) -> str | None:
    match = re.search(r"message_id=([^\)\s]+)", text)
    if not match:
        return None
    token = match.group(1).rstrip(").,")
    return token


def main() -> None:
    attachment_path = Path("mcp_attachment_test.txt")
    reply_draft_id: str | None = None
    compose_draft_id: str | None = None
    move_folder_name = "MCP_AUTOTEST_MOVE"
    base_test_folder = "MCP_AUTOTEST"
    renamed_folder = base_test_folder + "_RENAMED"

    try:
        section("List Folders (depth=1)")
        print(server.list_folders(max_depth=1, include_counts=False, include_paths=False))

        outlook, namespace = server.connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)
        inbox_id = inbox.EntryID

        section("Create Test Folder")
        print(server.create_folder(new_folder_name=base_test_folder, parent_folder_id=inbox_id, allow_existing=True))

        section("Folder Metadata")
        print(server.get_folder_metadata(folder_name=base_test_folder, include_children=False))

        section("Rename Folder")
        print(server.rename_folder(folder_name=base_test_folder, new_name=renamed_folder))

        section("Delete Folder")
        print(server.delete_folder(folder_name=renamed_folder, confirm=True))

        section("Create Move Target Folder")
        print(server.create_folder(new_folder_name=move_folder_name, parent_folder_id=inbox_id, allow_existing=True))
        move_folder_obj = server.get_folder_by_name(namespace, move_folder_name)
        move_folder_path = getattr(move_folder_obj, "FolderPath", None) if move_folder_obj else None

        section("List Recent Emails (latest)")
        recent_output = server.list_recent_emails(days=7, max_results=5, include_preview=False)
        print(recent_output)

        if server.email_cache:
            email_number = min(server.email_cache.keys())
            email_info = server.email_cache[email_number]
            message_id = email_info.get("id")
            original_folder_path = email_info.get("folder_path")
            original_unread = bool(email_info.get("unread"))
            original_categories = email_info.get("categories") or ""
            search_term = (
                email_info.get("subject") or email_info.get("sender") or email_info.get("sender_email") or "test"
            )

            section("Get Email By Number")
            print(server.get_email_by_number(email_number=email_number, include_body=False))

            section("Get Email Context")
            print(server.get_email_context(email_number=email_number, include_thread=False, thread_limit=3))

            section("Mark Email Read/Unread (toggle)")
            print(server.mark_email_read_unread(email_number=email_number, unread=not original_unread))

            section("Mark Email Read/Unread (revert)")
            print(server.mark_email_read_unread(email_number=email_number, unread=original_unread))

            section("Apply Category")
            print(server.apply_category(categories=["MCP-Test"], email_number=email_number))

            if original_categories:
                original_list = [c.strip() for c in original_categories.split(";") if c.strip()]
                section("Apply Category (revert)")
                print(server.apply_category(categories=original_list, email_number=email_number, overwrite=True))
            else:
                _, mail_item = server._resolve_mail_item(namespace, email_number=email_number)
                mail_item.Categories = ""
                mail_item.Save()
                print("Categorie ripristinate (vuote).")

            if move_folder_path and original_folder_path:
                section("Move Email To Folder (test)")
                print(server.move_email_to_folder(email_number=email_number, target_folder_path=move_folder_path))

                section("Move Email To Folder (revert)")
                print(server.move_email_to_folder(message_id=message_id, target_folder_path=original_folder_path))

            section("Get Attachments (message)")
            print(server.get_attachments(email_number=email_number, download=False, limit=3))

            section("Reply To Email (draft only)")
            reply_output = server.reply_to_email_by_number(
                email_number=email_number,
                reply_text="Test reply via MCP automation.",
                send=False,
            )
            print(reply_output)
            reply_draft_id = extract_message_id(reply_output)

            search_seed = search_term
        else:
            search_seed = "test"
            print("Nessuna email disponibile per i test delle operazioni sui messaggi.")

        server.clear_email_cache()

        section("List Recent Emails (offset=1)")
        print(server.list_recent_emails(days=7, max_results=3, include_preview=False, offset=1))

        section("Search Emails")
        print(server.search_emails(search_term=search_seed or "test", days=7, max_results=3, include_preview=False))

        section("List Sent Emails")
        print(server.list_sent_emails(days=14, max_results=5, include_preview=False, offset=1))

        section("List Pending Replies")
        print(server.list_pending_replies(days=7, max_results=5, include_preview=False))

        attachment_path.write_text("Allegato di prova MCP")

        section("Compose Email Draft (send=False)")
        compose_output = server.compose_email(
            recipient_email="recipient@example.com",
            subject="MCP Draft Test",
            body="Questo è un messaggio di prova (non inviato).",
            send=False,
        )
        print(compose_output)
        compose_draft_id = extract_message_id(compose_output)

        section("Attach File To Draft")
        print(
            server.attach_to_email(
                attachments=[str(attachment_path)],
                message_id=compose_draft_id,
                send=False,
            )
        )

        section("Get Attachments (draft)")
        if compose_draft_id:
            print(server.get_attachments(message_id=compose_draft_id, download=False, limit=5))
        else:
            print("Nessuna bozza disponibile per il recupero allegati.")

        cleanup_ids = [mid for mid in [reply_draft_id, compose_draft_id] if mid]
        if cleanup_ids:
            section("Batch Manage Emails (delete drafts)")
            print(server.batch_manage_emails(message_ids=cleanup_ids, delete=True))

    finally:
        if attachment_path.exists():
            attachment_path.unlink()
        try:
            server.delete_folder(folder_name=move_folder_name, confirm=True)
        except Exception:
            pass
        try:
            server.delete_folder(folder_name=base_test_folder, confirm=True)
        except Exception:
            pass
        try:
            server.delete_folder(folder_name=renamed_folder, confirm=True)
        except Exception:
            pass


if __name__ == "__main__":
    main()
