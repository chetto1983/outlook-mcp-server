"""MCP tools for inspecting and adding Outlook email attachments.

This module exposes two tools:
- `get_attachments`: list and optionally download attachments of a message
- `attach_to_email`: add one or more files to a message or draft (and optionally send)
"""

from __future__ import annotations

from typing import Any, Optional, List
import os

from ..features import feature_gate
from outlook_mcp_server import mcp  # FastMCP

from outlook_mcp import logger
from outlook_mcp.utils import coerce_bool, ensure_string_list, safe_filename, safe_entry_id
from outlook_mcp_server import _resolve_mail_item
from mcp.server.fastmcp.exceptions import ToolError


@mcp.tool()
@feature_gate(group="attachments")
def get_attachments(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    save_to: Optional[str] = None,
    download: bool = False,
    limit: Optional[int] = None,
) -> str:
    """Elenca (ed eventualmente scarica) gli allegati di un messaggio.

    Args:
        email_number: Numero del messaggio dall'ultimo elenco/ricerca (cache)
        message_id: EntryID del messaggio (alternativa a email_number)
        save_to: Cartella locale di destinazione per il download
        download: Se True, salva i file su disco in `save_to`
        limit: Massimo numero di allegati da processare
    """
    try:
        download_bool = coerce_bool(download)
        if limit is not None:
            try:
                limit_value = int(limit)
                if limit_value < 1:
                    return "Errore: 'limit' deve essere un intero positivo."
            except (TypeError, ValueError):
                return "Errore: 'limit' deve essere un intero positivo."
        else:
            limit_value = None

        if download_bool and not save_to:
            return "Errore: specifica 'save_to' quando download=True."

        logger.info(
            "get_attachments chiamato (numero=%s id=%s download=%s limit=%s destinazione=%s).",
            email_number,
            message_id,
            download_bool,
            limit_value,
            save_to,
        )

        from outlook_mcp_server import connect_to_outlook
        _, namespace = connect_to_outlook()
        try:
            _, mail_item = _resolve_mail_item(namespace, email_number=email_number, message_id=message_id)
        except ToolError as exc:
            return f"Errore: {exc}"

        if not hasattr(mail_item, "Attachments") or mail_item.Attachments.Count == 0:
            return "Questo messaggio non contiene allegati."

        total_attachments = mail_item.Attachments.Count
        max_index = total_attachments if limit_value is None else min(total_attachments, limit_value)

        if download_bool and save_to:
            os.makedirs(save_to, exist_ok=True)

        lines = [
            f"Allegati trovati: {total_attachments} (mostrati {max_index}).",
            "",
        ]
        saved_paths: List[str] = []

        for position in range(1, max_index + 1):
            attachment = mail_item.Attachments(position)
            name = getattr(attachment, "FileName", f"Allegato {position}")
            size = getattr(attachment, "Size", None)
            size_text = f"{size} byte" if isinstance(size, int) else "dimensione sconosciuta"
            lines.append(f"- {name} ({size_text})")

            if download_bool and save_to:
                safe_name = safe_filename(name or f"allegato_{position}")
                base, ext = os.path.splitext(safe_name)
                destination = os.path.join(save_to, safe_name)
                counter = 1
                while os.path.exists(destination):
                    destination = os.path.join(save_to, f"{base}_{counter}{ext}")
                    counter += 1
                try:
                    attachment.SaveAsFile(destination)
                    saved_paths.append(destination)
                except Exception as exc:
                    logger.exception("Impossibile salvare l'allegato %s.", name)
                    lines.append(f"  * Errore nel salvataggio: {exc}")

        if download_bool and saved_paths:
            lines.append("")
            lines.append("Allegati salvati in:")
            lines.extend(f"- {path}" for path in saved_paths)

        return "\n".join(lines)
    except Exception as exc:
        logger.exception("Errore durante get_attachments.")
        return f"Errore durante la gestione degli allegati: {exc}"


@mcp.tool()
@feature_gate(group="attachments")
def attach_to_email(
    attachments: Any,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    send: bool = False,
) -> str:
    """Aggiunge file come allegati a una bozza o risposta.

    Args:
        attachments: Percorso o lista di percorsi file da allegare
        email_number: Numero del messaggio dalla cache corrente
        message_id: EntryID del messaggio (se non si usa email_number)
        send: Se True, invia subito il messaggio dopo l'allegato
    """
    try:
        attachment_paths = ensure_string_list(attachments)
        if not attachment_paths:
            return "Errore: specifica almeno un percorso di file da allegare."

        send_bool = coerce_bool(send)
        logger.info(
            "attach_to_email chiamato (numero=%s id=%s allegati=%s invia=%s).",
            email_number,
            message_id,
            attachment_paths,
            send_bool,
        )

        from outlook_mcp_server import connect_to_outlook
        _, namespace = connect_to_outlook()
        try:
            _, mail_item = _resolve_mail_item(namespace, email_number=email_number, message_id=message_id)
        except ToolError as exc:
            return f"Errore: {exc}"

        attached_files: List[str] = []
        for path_value in attachment_paths:
            absolute = os.path.abspath(path_value)
            if not os.path.exists(absolute):
                return f"Errore: file '{absolute}' non trovato."
            try:
                mail_item.Attachments.Add(absolute)
                attached_files.append(absolute)
            except Exception as exc:
                logger.exception("Impossibile allegare il file %s.", absolute)
                return f"Errore: impossibile allegare '{absolute}' ({exc})."

        reference_id = message_id or safe_entry_id(mail_item) or "N/D"

        if send_bool:
            try:
                mail_item.Send()
            except Exception as exc:
                logger.exception("Invio del messaggio fallito dopo l'aggiunta degli allegati.")
                try:
                    mail_item.Save()
                except Exception:
                    pass
                return (
                    f"Allegati aggiunti ({len(attached_files)}), ma invio non riuscito: {exc}. "
                    f"(message_id={reference_id})"
                )
            return f"{len(attached_files)} allegati aggiunti e messaggio inviato (message_id={reference_id})."

        try:
            mail_item.Save()
        except Exception:
            pass
        return (
            f"Allegati aggiunti al messaggio ({', '.join(os.path.basename(p) for p in attached_files)}). "
            f"(message_id={reference_id})"
        )
    except Exception as exc:
        logger.exception("Errore durante attach_to_email.")
        return f"Errore durante l'aggiunta degli allegati: {exc}"
