"""MCP tools for acting on emails: reply, compose, move, mark, categorize, batch."""

from __future__ import annotations

from typing import Any, Optional, List, Dict
import os

from ..features import feature_gate
from outlook_mcp.toolkit import mcp_tool  # FastMCP instance

from outlook_mcp import logger
from outlook_mcp.utils import (
    coerce_bool,
    ensure_string_list,
    ensure_int_list,
    obfuscate_identifier,
    safe_entry_id,
    safe_folder_path,
    normalize_folder_path,
)
from outlook_mcp import folders as folder_service
from outlook_mcp.services.email import (
    resolve_mail_item,
    update_cached_email,
    apply_categories_to_item,
)
from outlook_mcp.services.common import (
    parse_importance,
    parse_sensitivity,
)

# Import runtime helpers lazily to avoid circular imports
def _connect():
    from outlook_mcp import connect_to_outlook

    return connect_to_outlook()


def _resolve(namespace, *, email_number: Optional[int], message_id: Optional[str]):
    return resolve_mail_item(namespace, email_number=email_number, message_id=message_id)


def _update_cache(number: Optional[int], **updates):
    update_cached_email(number, **updates)


def _apply_cats(mail_item, categories: List[str], overwrite: bool, append: bool) -> List[str]:
    return apply_categories_to_item(mail_item, categories, overwrite, append)


@mcp_tool()
@feature_gate(group="email.actions")
def move_email_to_folder(
    target_folder_id: Optional[str] = None,
    target_folder_path: Optional[str] = None,
    target_folder_name: Optional[str] = None,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    create_if_missing: bool = False,
) -> str:
    """Sposta un messaggio in una cartella risolta (id/path/nome), con creazione opzionale."""
    try:
        if not (target_folder_id or target_folder_path or target_folder_name):
            return "Errore: specifica una cartella di destinazione tramite id, path o nome."

        create_if_missing_bool = coerce_bool(create_if_missing)
        logger.info(
            "move_email_to_folder chiamato (numero=%s id=%s target_id=%s target_path=%s target_nome=%s crea=%s).",
            email_number,
            message_id,
            target_folder_id,
            target_folder_path,
            target_folder_name,
            create_if_missing_bool,
        )

        _, namespace = _connect()
        try:
            cached_entry, mail_item = _resolve(namespace, email_number=email_number, message_id=message_id)
        except Exception as exc:
            return f"Errore: {exc}"

        target_folder, attempts = folder_service.resolve_folder(
            namespace,
            folder_id=target_folder_id,
            folder_path=target_folder_path,
            folder_name=target_folder_name,
        )

        if not target_folder and create_if_missing_bool and target_folder_path:
            normalized = normalize_folder_path(target_folder_path)
            if normalized:
                segments = [segment for segment in normalized.split("\\") if segment]
                if segments:
                    parent_segments = segments[:-1]
                    leaf_name = segments[-1]
                    parent_folder = None
                    if parent_segments:
                        parent_path = "\\\\" + "\\".join(parent_segments)
                        parent_folder = folder_service.get_folder_by_path(namespace, parent_path)
                    if parent_folder:
                        try:
                            target_folder = parent_folder.Folders.Add(leaf_name)
                            attempts = []
                        except Exception as exc:
                            attempts.append(f"Creazione automatica '{leaf_name}' fallita: {exc}")
                    else:
                        attempts.append("Cartella padre non trovata per la creazione automatica.")

        if not target_folder:
            detail = "; ".join(attempts) if attempts else "cartella di destinazione non trovata."
            return f"Errore: {detail}"

        try:
            moved_item = mail_item.Move(target_folder)
        except Exception as exc:
            logger.exception("Outlook ha rifiutato lo spostamento del messaggio.")
            return f"Errore: impossibile spostare il messaggio ({exc})."

        destination_path = safe_folder_path(target_folder) or getattr(target_folder, "Name", "(destinazione)")
        new_entry_id = safe_entry_id(moved_item) or safe_entry_id(mail_item)
        if email_number is not None:
            _update_cache(
                email_number,
                folder_path=destination_path,
                id=new_entry_id,
            )

        reference = f"#{email_number}" if email_number is not None else (message_id or new_entry_id or "N/D")
        return (
            f"Messaggio {reference} spostato nella cartella '{destination_path}'. "
            f"(message_id={new_entry_id or 'N/D'})"
        )
    except Exception as exc:
        logger.exception("Errore durante move_email_to_folder.")
        return f"Errore durante lo spostamento del messaggio: {exc}"


@mcp_tool()
@feature_gate(group="email.actions")
def mark_email_read_unread(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    unread: Optional[bool] = None,
    flag: Optional[str] = None,
) -> str:
    """Imposta lo stato di lettura (letto/non letto) di un messaggio."""
    try:
        if unread is None and flag is None:
            return "Errore: specifica 'unread' True/False oppure flag='read'/'unread'."

        if unread is not None:
            target_unread = bool(coerce_bool(unread))
        else:
            normalized = str(flag).strip().lower()
            if normalized in {"read", "letto", "letta"}:
                target_unread = False
            elif normalized in {"unread", "non letto", "non letta"}:
                target_unread = True
            else:
                return "Errore: flag deve essere 'read' o 'unread'."

        logger.info(
            "mark_email_read_unread chiamato (numero=%s id=%s unread=%s).",
            email_number,
            message_id,
            target_unread,
        )

        _, namespace = _connect()
        try:
            _, mail_item = _resolve(namespace, email_number=email_number, message_id=message_id)
        except Exception as exc:
            return f"Errore: {exc}"

        try:
            mail_item.UnRead = target_unread
            mail_item.Save()
        except Exception as exc:
            logger.exception("Outlook ha rifiutato l'aggiornamento dello stato lettura.")
            return f"Errore: impossibile aggiornare lo stato lettura ({exc})."

        _update_cache(email_number, unread=target_unread)
        status_label = "Non letta" if target_unread else "Letta"
        reference = f"#{email_number}" if email_number is not None else (message_id or safe_entry_id(mail_item) or "messaggio")
        return f"Messaggio {reference} contrassegnato come {status_label}."
    except Exception as exc:
        logger.exception("Errore durante mark_email_read_unread.")
        return f"Errore durante l'aggiornamento dello stato di lettura: {exc}"


@mcp_tool()
@feature_gate(group="email.actions")
def apply_category(
    categories: Optional[Any] = None,
    category: Optional[str] = None,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    overwrite: bool = False,
    append: bool = False,
) -> str:
    """Applica una o piu' categorie Outlook a un messaggio (unisci/sovrascrivi)."""
    try:
        category_list = ensure_string_list(categories)
        if category:
            category_list.extend(ensure_string_list(category))
        category_list = [cat for cat in category_list if cat]
        if not category_list:
            return "Errore: specifica almeno una categoria da applicare."

        overwrite_bool = coerce_bool(overwrite)
        append_bool = coerce_bool(append)
        logger.info(
            "apply_category chiamato (categorie=%s numero=%s id=%s overwrite=%s append=%s).",
            category_list,
            email_number,
            message_id,
            overwrite_bool,
            append_bool,
        )

        _, namespace = _connect()
        try:
            _, mail_item = _resolve(namespace, email_number=email_number, message_id=message_id)
        except Exception as exc:
            return f"Errore: {exc}"

        try:
            final_categories = _apply_cats(mail_item, category_list, overwrite_bool, append_bool)
        except ValueError as exc:
            return f"Errore: {exc}"

        if email_number is not None:
            _update_cache(email_number, categories="; ".join(final_categories))

        reference = f"#{email_number}" if email_number is not None else (message_id or safe_entry_id(mail_item) or "messaggio")
        return f"Categorie applicate al messaggio {reference}: {', '.join(final_categories) if final_categories else '(nessuna)'}."
    except Exception as exc:
        logger.exception("Errore durante apply_category.")
        return f"Errore durante l'aggiornamento delle categorie: {exc}"


@mcp_tool()
@feature_gate(group="email.actions")
def set_email_category(
    email_number: int,
    category: str,
    overwrite: bool = False,
) -> str:
    """Compatibilita' retro: applica una singola categoria (alias di apply_category)."""
    overwrite_bool = coerce_bool(overwrite)
    append_flag = not overwrite_bool
    return apply_category(
        categories=[category],
        email_number=email_number,
        overwrite=overwrite_bool,
        append=append_flag,
    )


@mcp_tool()
@feature_gate(group="email.actions")
def reply_to_email_by_number(
    email_number: Optional[int] = None,
    reply_text: str = "",
    message_id: Optional[str] = None,
    reply_all: bool = False,
    send: bool = True,
    attachments: Optional[Any] = None,
    use_html: bool = False,
    importance: Optional[str] = None,
    sensitivity: Optional[str] = None,
    request_read_receipt: bool = False,
    request_delivery_receipt: bool = False,
) -> str:
    """Risponde a un messaggio (reply/reply-all), con allegati, proprietà avanzate e invio opzionale."""
    try:
        if not reply_text.strip():
            return "Errore: specifica il testo della risposta."

        reply_all_bool = coerce_bool(reply_all)
        send_bool = coerce_bool(send)
        use_html_bool = coerce_bool(use_html)
        request_read_bool = coerce_bool(request_read_receipt)
        request_delivery_bool = coerce_bool(request_delivery_receipt)
        attachment_paths = ensure_string_list(attachments)
        logger.info(
            "reply_to_email_by_number chiamato (numero=%s id=%s reply_all=%s invia=%s allegati=%s html=%s importance=%s sensitivity=%s).",
            email_number,
            message_id,
            reply_all_bool,
            send_bool,
            attachment_paths,
            use_html_bool,
            importance,
            sensitivity,
        )

        _, namespace = _connect()
        try:
            _, mail_item = _resolve(namespace, email_number=email_number, message_id=message_id)
        except Exception as exc:
            return f"Errore: {exc}"

        reply = mail_item.ReplyAll() if reply_all_bool else mail_item.Reply()
        try:
            if use_html_bool:
                original_body = getattr(reply, "HTMLBody", "")
                reply.HTMLBody = f"{reply_text}<br><br>{original_body}"
            else:
                original_body = getattr(reply, "Body", "")
                reply.Body = f"{reply_text}\n\n{original_body}"
        except Exception:
            reply.Body = reply_text

        # Set advanced properties
        if importance:
            importance_code = parse_importance(importance)
            if importance_code is not None:
                reply.Importance = importance_code
            else:
                return "Errore: importanza non valida. Usa: bassa, normale, alta"

        if sensitivity:
            sensitivity_code = parse_sensitivity(sensitivity)
            if sensitivity_code is not None:
                reply.Sensitivity = sensitivity_code
            else:
                return "Errore: sensibilità non valida. Usa: normale, personale, privato, confidenziale"

        if request_read_bool:
            reply.ReadReceiptRequested = True

        if request_delivery_bool:
            reply.OriginatorDeliveryReportRequested = True

        for path_value in attachment_paths:
            absolute = os.path.abspath(path_value)
            if not os.path.exists(absolute):
                return f"Errore: file '{absolute}' non trovato."
            try:
                reply.Attachments.Add(absolute)
            except Exception as exc:
                logger.exception("Impossibile allegare il file %s alla risposta.", absolute)
                return f"Errore: impossibile allegare '{absolute}' ({exc})."

        if send_bool:
            reply.Send()
        else:
            try:
                reply.Save()
            except Exception:
                pass

        sender_name = getattr(mail_item, "SenderName", "Destinatario")
        entry_id = safe_entry_id(mail_item)
        action = "inviata" if send_bool else "salvata in Bozze"
        return (
            f"Risposta {action} per {sender_name}. "
            f"(message_id={entry_id or message_id or 'N/D'})"
        )
    except Exception as exc:
        logger.exception("Errore durante reply_to_email_by_number (numero=%s id=%s).", email_number, obfuscate_identifier(message_id))
        return f"Errore durante l'invio della risposta: {exc}"


@mcp_tool()
@feature_gate(group="email.actions")
def compose_email(
    recipient_email: str,
    subject: str,
    body: str,
    cc_email: Optional[str] = None,
    bcc_email: Optional[str] = None,
    attachments: Optional[Any] = None,
    send: bool = True,
    use_html: bool = False,
    importance: Optional[str] = None,
    sensitivity: Optional[str] = None,
    request_read_receipt: bool = False,
    request_delivery_receipt: bool = False,
    voting_options: Optional[str] = None,
) -> str:
    """Crea e invia/archivia una nuova email con CC/BCC, allegati e proprietà avanzate."""
    try:
        if not recipient_email.strip():
            return "Errore: specifica almeno un destinatario."

        send_bool = coerce_bool(send)
        use_html_bool = coerce_bool(use_html)
        request_read_bool = coerce_bool(request_read_receipt)
        request_delivery_bool = coerce_bool(request_delivery_receipt)
        attachment_paths = ensure_string_list(attachments)
        logger.info(
            "compose_email chiamato (destinatario=%s cc=%s bcc=%s oggetto='%s' invia=%s allegati=%s html=%s importance=%s sensitivity=%s).",
            recipient_email,
            cc_email,
            bcc_email,
            subject,
            send_bool,
            attachment_paths,
            use_html_bool,
            importance,
            sensitivity,
        )

        outlook, _ = _connect()
        mail = outlook.CreateItem(0)
        mail.To = recipient_email
        if cc_email:
            mail.CC = cc_email
        if bcc_email:
            mail.BCC = bcc_email
        mail.Subject = subject

        if use_html_bool:
            existing = getattr(mail, "HTMLBody", "")
            mail.HTMLBody = f"{body}{existing}"
        else:
            mail.Body = body

        # Set advanced properties
        if importance:
            importance_code = parse_importance(importance)
            if importance_code is not None:
                mail.Importance = importance_code
            else:
                return "Errore: importanza non valida. Usa: bassa, normale, alta"

        if sensitivity:
            sensitivity_code = parse_sensitivity(sensitivity)
            if sensitivity_code is not None:
                mail.Sensitivity = sensitivity_code
            else:
                return "Errore: sensibilità non valida. Usa: normale, personale, privato, confidenziale"

        if request_read_bool:
            mail.ReadReceiptRequested = True

        if request_delivery_bool:
            mail.OriginatorDeliveryReportRequested = True

        if voting_options:
            # Voting options format: "Approve;Reject;Review"
            mail.VotingOptions = voting_options

        for path_value in attachment_paths:
            absolute = os.path.abspath(path_value)
            if not os.path.exists(absolute):
                return f"Errore: file '{absolute}' non trovato."
            try:
                mail.Attachments.Add(absolute)
            except Exception as exc:
                logger.exception("Impossibile allegare il file %s alla bozza.", absolute)
                return f"Errore: impossibile allegare '{absolute}' ({exc})."

        if send_bool:
            mail.Send()
            return f"Email inviata a: {recipient_email}"

        try:
            mail.Save()
        except Exception:
            pass
        entry_id = safe_entry_id(mail)
        return f"Bozza salvata (message_id={entry_id or 'N/D'})."
    except Exception as exc:
        logger.exception("Errore durante compose_email per destinatario %s.", recipient_email)
        return f"Errore durante la composizione dell'email: {exc}"


@mcp_tool()
@feature_gate(group="batch")
def batch_manage_emails(
    email_numbers: Optional[Any] = None,
    message_ids: Optional[Any] = None,
    move_to_folder_id: Optional[str] = None,
    move_to_folder_path: Optional[str] = None,
    move_to_folder_name: Optional[str] = None,
    mark_as: Optional[str] = None,
    delete: bool = False,
) -> str:
    """Operazioni batch: sposta, marca come letto/non letto o elimina per numeri/ID."""
    try:
        numbers = ensure_int_list(email_numbers)
        ids = ensure_string_list(message_ids)
        if not numbers and not ids:
            return "Errore: specifica almeno un email_number o message_id."

        delete_bool = coerce_bool(delete)

        mark_target: Optional[bool] = None
        if mark_as is not None:
            normalized = str(mark_as).strip().lower()
            if normalized in {"read", "letto", "letta"}:
                mark_target = False
            elif normalized in {"unread", "non letto", "non letta"}:
                mark_target = True
            else:
                return "Errore: 'mark_as' deve essere 'read' o 'unread'."

        move_requested = any([move_to_folder_id, move_to_folder_path, move_to_folder_name])
        masked_ids = [obfuscate_identifier(entry) for entry in ids]
        logger.info(
            "batch_manage_emails chiamato (numeri=%s ids=%s move=%s mark=%s delete=%s).",
            numbers,
            masked_ids,
            move_requested,
            mark_target,
            delete_bool,
        )

        _, namespace = _connect()
        target_folder = None
        target_attempts: List[str] = []
        if move_requested:
            target_folder, target_attempts = folder_service.resolve_folder(
                namespace,
                folder_id=move_to_folder_id,
                folder_path=move_to_folder_path,
                folder_name=move_to_folder_name,
            )
            if not target_folder:
                detail = "; ".join(target_attempts) if target_attempts else "cartella di destinazione non trovata."
                return f"Errore: {detail}"

        successes: List[str] = []
        failures: List[str] = []

        def process_email(number: Optional[int], entry_id: Optional[str], label: str) -> None:
            try:
                _, mail_item = _resolve(namespace, email_number=number, message_id=entry_id)
            except Exception as exc:
                failures.append(f"{label}: {exc}")
                return

            reference_id = safe_entry_id(mail_item) or entry_id or "N/D"
            operations: List[str] = []

            if move_requested and target_folder:
                try:
                    mail_item = mail_item.Move(target_folder)
                    operations.append(
                        f"spostato in {safe_folder_path(target_folder) or getattr(target_folder, 'Name', '')}"
                    )
                except Exception as exc:
                    failures.append(f"{label} (id={reference_id}): errore nello spostamento ({exc})")
                    return

            if mark_target is not None:
                try:
                    mail_item.UnRead = mark_target
                    operations.append(
                        "contrassegnato come non letto" if mark_target else "contrassegnato come letto"
                    )
                except Exception as exc:
                    failures.append(
                        f"{label} (id={reference_id}): impossibile aggiornare lo stato lettura ({exc})"
                    )
                    return

            if delete_bool:
                try:
                    mail_item.Delete()
                    operations.append("eliminato")
                except Exception as exc:
                    failures.append(f"{label} (id={reference_id}): eliminazione non riuscita ({exc})")
                    return
            else:
                try:
                    mail_item.Save()
                except Exception:
                    pass

            if number is not None:
                if delete_bool:
                    try:
                        from outlook_mcp import email_cache

                        email_cache.pop(number, None)
                    except Exception:
                        pass
                else:
                    updates: Dict[str, Any] = {}
                    if move_requested and target_folder:
                        updates["folder_path"] = safe_folder_path(target_folder)
                        updates["id"] = safe_entry_id(mail_item) or reference_id
                    if mark_target is not None:
                        updates["unread"] = mark_target
                    if updates:
                        _update_cache(number, **updates)

            final_ref = safe_entry_id(mail_item) or reference_id
            if not operations:
                operations.append("nessuna modifica")
            successes.append(f"{label} (id={final_ref}): {', '.join(operations)}")

        for number in numbers:
            process_email(number, None, f"numero={number}")
        for entry_id in ids:
            process_email(None, entry_id, f"id={entry_id}")

        result_lines = [
            f"Operazioni riuscite: {len(successes)}",
            f"Operazioni fallite: {len(failures)}",
        ]
        if successes:
            result_lines.append("")
            result_lines.append("Dettagli riusciti:")
            result_lines.extend(f"- {line}" for line in successes)
        if failures:
            result_lines.append("")
            result_lines.append("Errori:")
            result_lines.extend(f"- {line}" for line in failures)

        return "\n".join(result_lines)
    except Exception as exc:
        logger.exception("Errore durante batch_manage_emails.")
        return f"Errore durante le operazioni batch sui messaggi: {exc}"


@mcp_tool()
@feature_gate(group="email.actions")
def flag_email(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    flag_status: Optional[str] = None,
    due_date: Optional[str] = None,
    reminder_time: Optional[str] = None,
    clear_flag: bool = False,
) -> str:
    """Imposta o rimuove il contrassegno follow-up su un'email con promemoria opzionale."""
    try:
        import datetime

        clear_flag_bool = coerce_bool(clear_flag)

        logger.info(
            "flag_email chiamato (numero=%s id=%s flag_status=%s due_date=%s reminder=%s clear=%s).",
            email_number,
            message_id,
            flag_status,
            due_date,
            reminder_time,
            clear_flag_bool,
        )

        _, namespace = _connect()
        try:
            _, mail_item = _resolve(namespace, email_number=email_number, message_id=message_id)
        except Exception as exc:
            return f"Errore: {exc}"

        if clear_flag_bool:
            # Clear flag
            try:
                mail_item.FlagRequest = ""
                mail_item.ReminderSet = False
                mail_item.Save()
                reference = f"#{email_number}" if email_number is not None else (message_id or "messaggio")
                return f"Contrassegno rimosso dal messaggio {reference}."
            except Exception as exc:
                logger.exception("Impossibile rimuovere il contrassegno.")
                return f"Errore: impossibile rimuovere il contrassegno ({exc})"

        # Set flag
        flag_request = flag_status or "Follow up"
        try:
            mail_item.FlagRequest = flag_request
        except Exception as exc:
            logger.exception("Impossibile impostare il contrassegno.")
            return f"Errore: impossibile impostare il contrassegno ({exc})"

        # Set due date
        if due_date:
            try:
                due_dt = datetime.datetime.fromisoformat(due_date.replace("T", " ").replace("Z", ""))
                mail_item.FlagDueBy = due_dt
            except Exception as exc:
                logger.warning("Formato data scadenza non valido: %s (%s)", due_date, exc)
                return f"Errore: formato data scadenza non valido. Usa formato ISO (es: 2025-10-22 o 2025-10-22T14:30)"

        # Set reminder
        if reminder_time:
            try:
                reminder_dt = datetime.datetime.fromisoformat(reminder_time.replace("T", " ").replace("Z", ""))
                mail_item.ReminderSet = True
                mail_item.ReminderTime = reminder_dt
            except Exception as exc:
                logger.warning("Formato data promemoria non valido: %s (%s)", reminder_time, exc)
                return f"Errore: formato data promemoria non valido. Usa formato ISO (es: 2025-10-22T14:30)"
        elif due_date:
            # If due_date is set but no reminder, set reminder to same time
            try:
                due_dt = datetime.datetime.fromisoformat(due_date.replace("T", " ").replace("Z", ""))
                mail_item.ReminderSet = True
                mail_item.ReminderTime = due_dt
            except Exception:
                pass

        try:
            mail_item.Save()
        except Exception as exc:
            logger.exception("Impossibile salvare il messaggio con il contrassegno.")
            return f"Errore: impossibile salvare il messaggio ({exc})"

        reference = f"#{email_number}" if email_number is not None else (message_id or "messaggio")
        parts = [f"Messaggio {reference} contrassegnato: {flag_request}"]
        if due_date:
            parts.append(f"Scadenza: {due_date}")
        if reminder_time:
            parts.append(f"Promemoria: {reminder_time}")

        return " | ".join(parts)

    except Exception as exc:
        logger.exception("Errore durante flag_email.")
        return f"Errore durante l'impostazione del contrassegno: {exc}"
