"""MCP tools for detailed message inspection and context outlines."""

from __future__ import annotations

from typing import Optional, Any, Dict, List

from ..features import feature_gate
from outlook_mcp_server import mcp  # FastMCP instance

from outlook_mcp import logger
from outlook_mcp import folders as folder_service
from outlook_mcp import email_cache
from outlook_mcp.utils import (
    coerce_bool,
    trim_conversation_id,
    safe_entry_id,
    safe_folder_path,
)

# Reuse helpers from the service layer
from outlook_mcp.services.email import resolve_mail_item, format_email, build_conversation_outline


@mcp.tool()
@feature_gate(group="email.detail")
def get_email_by_number(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    index: Optional[int] = None,
    include_body: bool = True,
) -> str:
    """Recupera i dettagli completi di un messaggio via cache/id o posizione cartella."""
    try:
        include_body_bool = coerce_bool(include_body)
        logger.info(
            "get_email_by_number chiamato (numero=%s id=%s folder_id=%s folder_path=%s index=%s corpo=%s).",
            email_number,
            message_id,
            folder_id,
            folder_path,
            index,
            include_body_bool,
        )

        from outlook_mcp import connect_to_outlook
        _, namespace = connect_to_outlook()

        email_data: Optional[Dict[str, Any]] = None
        mail_item = None
        number_ref = email_number

        if email_number is not None:
            if not email_cache:
                return "Errore: nessun elenco messaggi attivo. Mostra prima le email e poi ripeti la richiesta."
            if email_number not in email_cache:
                return f"Errore: il messaggio #{email_number} non e presente nell'elenco corrente."
            email_data = dict(email_cache[email_number])
            message_id = message_id or email_data.get("id")

        if message_id and not mail_item:
            try:
                mail_item = namespace.GetItemFromID(message_id)
            except Exception:
                mail_item = None

        if mail_item is None and (folder_id or folder_path):
            if index is None:
                return "Errore: specifica 'index' (posizione 1-based) quando usi folder_id o folder_path."
            if not isinstance(index, int) or index < 1:
                return "Errore: 'index' deve essere un intero positivo (1-based)."
            folder, attempts = folder_service.resolve_folder(
                namespace,
                folder_id=folder_id,
                folder_path=folder_path,
                folder_name=None,
            )
            if not folder:
                detail = "; ".join(attempts) if attempts else "cartella non trovata."
                return f"Errore: impossibile individuare la cartella specificata ({detail})."
            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                if index > items.Count:
                    return f"Errore: la cartella contiene solo {items.Count} elementi."
                mail_item = items(index)
                message_id = message_id or safe_entry_id(mail_item)
            except Exception as exc:
                logger.exception("Impossibile recuperare il messaggio %s dalla cartella.", index)
                return f"Errore: impossibile recuperare il messaggio in posizione {index} ({exc})."

        if mail_item and email_data is None:
            try:
                email_data = format_email(mail_item)
            except Exception as exc:
                logger.warning("Impossibile formattare il messaggio recuperato: %s", exc)
                email_data = None

        if email_data is None:
            return "Errore: impossibile recuperare i dettagli del messaggio richiesto."

        trimmed_conv = trim_conversation_id(email_data.get("conversation_id"), max_chars=32)
        importance_label = email_data.get("importance_label") or str(email_data.get("importance", ""))
        attachment_names_preview = email_data.get("attachment_names") or []
        to_line = ", ".join(email_data.get("to_recipients", []))
        cc_line = ", ".join(email_data.get("cc_recipients", []))
        bcc_line = ", ".join(email_data.get("bcc_recipients", []))

        identifier_line = (
            f"Dettagli del messaggio #{number_ref}:" if number_ref is not None else "Dettagli del messaggio richiesto:"
        )
        result_lines = [
            identifier_line,
            "",
            f"Oggetto: {email_data.get('subject', '(Senza oggetto)')}",
            f"Da: {email_data.get('sender', 'Sconosciuto')} <{email_data.get('sender_email', '')}>",
        ]

        if to_line:
            result_lines.append(f"A: {to_line}")
        if cc_line:
            result_lines.append(f"Cc: {cc_line}")
        if bcc_line:
            result_lines.append(f"Ccn: {bcc_line}")

        folder_display = email_data.get("folder_path") or ""
        if not folder_display and mail_item:
            folder_display = safe_folder_path(getattr(mail_item, "Parent", None))

        result_lines.extend(
            [
                f"Ricevuto: {email_data.get('received_time', 'Sconosciuto')}",
                f"Cartella: {folder_display or 'Cartella sconosciuta'}",
                f"Importanza: {importance_label}",
                # Stato lettura già stampato a elenco, qui non indispensabile
            ]
        )

        if email_data.get("categories"):
            result_lines.append(f"Categorie: {email_data['categories']}")
        if trimmed_conv:
            result_lines.append(f"ID conversazione: {trimmed_conv}")
        if email_data.get("preview"):
            result_lines.append(f"Anteprima corpo: {email_data['preview']}")

        if message_id:
            result_lines.append(f"MessageID: {message_id}")

        result_lines.append(f"Allegati: {'Si' if email_data.get('has_attachments') else 'No'}")
        if attachment_names_preview:
            result_lines.append(f"Nomi allegati: {', '.join(attachment_names_preview)}")

        attachment_lines: List[str] = []
        if mail_item and email_data.get("has_attachments") and hasattr(mail_item, "Attachments"):
            try:
                for i in range(1, mail_item.Attachments.Count + 1):
                    attachment = mail_item.Attachments(i)
                    attachment_lines.append(f"  - {attachment.FileName}")
            except Exception:
                attachment_lines = []

        if attachment_lines:
            result_lines.append("")
            result_lines.append("Allegati dettagliati:")
            result_lines.extend(attachment_lines)

        if include_body_bool:
            body_content = email_data.get("body")
            if not body_content and mail_item:
                try:
                    body_content = getattr(mail_item, "Body", "")
                except Exception:
                    body_content = ""
            result_lines.append("")
            result_lines.append("Corpo:")
            result_lines.append(body_content or "(Nessun contenuto)")

        result_lines.append("")
        result_lines.append(
            "Puoi chiedermi di rispondere o inoltrare questo messaggio indicando il numero o il message_id."
        )

        return "\n".join(result_lines)

    except Exception as exc:
        logger.exception("Errore nel recupero dei dettagli del messaggio (numero=%s id=%s).", email_number, message_id)
        return f"Errore durante il recupero dei dettagli del messaggio: {exc}"


@mcp.tool()
@feature_gate(group="email.detail")
def get_email_context(
    email_number: int,
    include_thread: bool = True,
    thread_limit: int = 6,
    lookback_days: int = 45,
) -> str:
    """Restituisce un contesto sintetico per la conversazione dell'email indicata."""
    try:
        from outlook_mcp import connect_to_outlook
        _, namespace = connect_to_outlook()
        include_thread_bool = coerce_bool(include_thread)

        if not email_cache or email_number not in email_cache:
            return (
                "Errore: nessun elenco messaggi attivo o numero non valido. "
                "Elenca prima le email per costruire il contesto."
            )

        focus_email = email_cache[email_number]
        outline = build_conversation_outline(
            namespace=namespace,
            email_data=focus_email,
            lookback_days=lookback_days,
            max_items=max(1, int(thread_limit)),
        )

        lines = [f"Contesto per messaggio #{email_number}:"]
        if outline:
            lines.append("")
            lines.append(outline)
        else:
            lines.append("(Nessun altro elemento di conversazione trovato nei limiti impostati)")

        if include_thread_bool:
            lines.append("")
            lines.append("Suggerimento: usa 'reply_to_email_by_number' per rispondere con testo mirato.")

        return "\n".join(lines)
    except Exception as exc:
        logger.exception("Errore durante get_email_context per #%s.", email_number)
        return f"Errore durante il recupero del contesto: {exc}"
