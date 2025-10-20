"""Service layer with reusable email helpers for Outlook MCP tools and server."""

from __future__ import annotations

import datetime
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from mcp.server.fastmcp.exceptions import ToolError

from outlook_mcp import (
    DEFAULT_DOMAIN_ROOT_NAME,
    DEFAULT_DOMAIN_SUBFOLDERS,
    LAST_VERB_REPLY_CODES,
    MAX_EMAIL_SCAN_PER_FOLDER,
    PR_LAST_VERB_EXECUTED,
    PR_LAST_VERB_EXECUTION_TIME,
    clear_email_cache,
    connect_to_outlook,
    email_cache,
    logger,
)
from outlook_mcp.utils import (
    build_body_preview,
    coerce_bool,
    extract_attachment_names,
    extract_recipients,
    normalize_folder_path,
    safe_entry_id,
    safe_folder_path,
    to_python_datetime,
    trim_conversation_id,
)
from outlook_mcp import folders as folder_service

from .common import (
    describe_importance,
    format_read_status,
    format_yes_no,
    parse_datetime_string,
)

__all__ = [
    "resolve_mail_item",
    "update_cached_email",
    "normalize_email_address",
    "extract_email_domain",
    "derive_sender_email",
    "ensure_domain_folder_structure",
    "collect_user_addresses",
    "mail_item_marked_replied",
    "format_email",
    "get_emails_from_folder",
    "resolve_additional_folders",
    "get_all_mail_folders",
    "collect_emails_across_folders",
    "get_related_conversation_emails",
    "email_has_user_reply",
    "email_has_user_reply_with_context",
    "build_conversation_outline",
    "present_email_listing",
    "apply_categories_to_item",
    "get_email_context",
]


# ---------------------------------------------------------------------------
# Fundamental helpers
# ---------------------------------------------------------------------------
def resolve_mail_item(
    namespace,
    *,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
) -> Tuple[Optional[Dict[str, Any]], Any]:
    """Return an Outlook mail item and cached metadata if available."""
    cached_entry: Optional[Dict[str, Any]] = None

    if email_number is not None:
        if not email_cache or email_number not in email_cache:
            raise ToolError(
                "Messaggio non presente nella cache corrente. Elenca prima le email o specifica un message_id."
            )
        cached_entry = email_cache[email_number]
        if not message_id:
            message_id = cached_entry.get("id")

    if not message_id:
        raise ToolError("Specificare 'message_id' oppure un 'email_number' precedentemente elencato.")

    try:
        mail_item = namespace.GetItemFromID(message_id)
    except Exception as exc:  # pragma: no cover - Outlook COM guarded
        raise ToolError(f"Impossibile recuperare il messaggio con ID '{message_id}': {exc}") from exc

    if not mail_item:
        raise ToolError("Outlook ha restituito un elemento vuoto per l'ID specificato.")

    return cached_entry, mail_item


def update_cached_email(email_number: Optional[int], **updates: Any) -> None:
    """Apply in-place updates to the cached representation of an email."""
    if email_number is None:
        return
    if not email_cache or email_number not in email_cache:
        return
    email_cache[email_number].update({key: value for key, value in updates.items() if value is not None})


def normalize_email_address(value: Optional[str]) -> Optional[str]:
    """Return a normalized lowercase email address when possible."""
    if not value:
        return None
    text = value.strip()
    if not text:
        return None
    if "<" in text and ">" in text:
        start = text.find("<")
        end = text.rfind(">")
        if start < end:
            text = text[start + 1 : end]
    text = text.replace(",", ";").strip()
    segments = [segment.strip() for segment in text.split(";") if segment.strip()]
    if not segments:
        segments = [text]
    for segment in segments:
        candidate = segment
        if ":" in candidate:
            prefix, remainder = candidate.split(":", 1)
            if prefix.lower() in {"smtp", "sip", "mailto"}:
                candidate = remainder
        candidate = candidate.strip().strip("<>").strip().lower()
        if not candidate:
            continue
        if "@" in candidate:
            return candidate
    fallback = segments[0].strip().lower() if segments else text.lower()
    return fallback or None


def extract_email_domain(address: Optional[str]) -> Optional[str]:
    """Return email domain portion."""
    normalized = normalize_email_address(address)
    if not normalized or "@" not in normalized:
        return None
    return normalized.split("@", 1)[1]


def derive_sender_email(entry: Dict[str, Any]) -> Optional[str]:
    """Extract sender email from cached entry."""
    return entry.get("sender_email") or entry.get("sender")


def _get_or_create_subfolder(parent, name: str):
    """Return existing Outlook subfolder or create it."""
    try:
        for sub in parent.Folders:
            if sub.Name.lower() == name.lower():
                return sub, False
    except Exception:
        logger.debug("Impossibile enumerare le sottocartelle di '%s'.", getattr(parent, "Name", parent), exc_info=True)
    try:
        return parent.Folders.Add(name), True
    except Exception as exc:  # pragma: no cover - Outlook COM guarded
        raise Exception(f"Impossibile creare la cartella '{name}': {exc}") from exc


def ensure_domain_folder_structure(
    namespace,
    domain: str,
    root_folder_name: str = DEFAULT_DOMAIN_ROOT_NAME,
    subfolders: Optional[Sequence[str]] = None,
):
    """Ensure domain folder and optional subfolders exist under Inbox."""
    if not subfolders:
        subfolders = DEFAULT_DOMAIN_SUBFOLDERS
    inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
    root_folder, _ = _get_or_create_subfolder(inbox, root_folder_name)
    domain_folder, domain_created = _get_or_create_subfolder(root_folder, domain)
    created_subfolders: List[str] = []
    for name in subfolders:
        existing = False
        try:
            for sub in domain_folder.Folders:
                if sub.Name.lower() == name.lower():
                    existing = True
                    break
        except Exception:
            existing = False
        if existing:
            continue
        try:
            folder, created = _get_or_create_subfolder(domain_folder, name)
        except Exception as exc:
            logger.warning("Creazione della sottocartella '%s' fallita: %s", name, exc)
            continue
        if created:
            created_subfolders.append(folder.Name)
    return domain_folder, domain_created, created_subfolders


def _extract_best_timestamp(entry: Optional[Dict[str, Any]]) -> Optional[datetime.datetime]:
    """Derive the most relevant timestamp from a formatted email dictionary."""
    if not entry:
        return None
    for key in ("received_iso", "sent_iso", "last_modified_iso"):
        dt = parse_datetime_string(entry.get(key))
        if dt:
            return dt
    for key in ("received_time", "sent_time", "last_modified_time"):
        dt = parse_datetime_string(entry.get(key))
        if dt:
            return dt
    return None


def collect_user_addresses(namespace) -> Set[str]:
    """Collect SMTP-style addresses that belong to the local Outlook profile."""
    addresses: Set[str] = set()
    try:
        current_user = getattr(namespace, "CurrentUser", None)
        if current_user:
            for attr_name in ("Address", "Name"):
                attr_value = getattr(current_user, attr_name, None)
                if attr_value:
                    addresses.add(str(attr_value))
            address_entry = getattr(current_user, "AddressEntry", None)
            if address_entry:
                entry_address = getattr(address_entry, "Address", None)
                if entry_address:
                    addresses.add(str(entry_address))
                try:
                    exchange_user = address_entry.GetExchangeUser()
                    if exchange_user:
                        primary = getattr(exchange_user, "PrimarySmtpAddress", None)
                        if primary:
                            addresses.add(str(primary))
                except Exception:
                    pass
                try:
                    exchange_dl = address_entry.GetExchangeDistributionList()
                    if exchange_dl:
                        primary = getattr(exchange_dl, "PrimarySmtpAddress", None)
                        if primary:
                            addresses.add(str(primary))
                except Exception:
                    pass
    except Exception:
        logger.debug("Impossibile leggere CurrentUser da Outlook.", exc_info=True)

    try:
        session = namespace.Application.Session
        accounts = getattr(session, "Accounts", None)
        if accounts:
            for idx in range(1, accounts.Count + 1):
                account = accounts.Item(idx)
                for attr_name in ("SmtpAddress", "DisplayName"):
                    attr_value = getattr(account, attr_name, None)
                    if attr_value:
                        addresses.add(str(attr_value))
    except Exception:
        logger.debug("Impossibile enumerare gli account Outlook disponibili.", exc_info=True)

    normalized = {
        normalized
        for normalized in (normalize_email_address(addr) for addr in addresses)
        if normalized
    }
    if normalized:
        logger.debug("Rilevati indirizzi utente: %s", ", ".join(sorted(normalized)))
    else:
        logger.debug("Nessun indirizzo utente normalizzato rilevato.")
    return normalized


def mail_item_marked_replied(mail_item, baseline: Optional[datetime.datetime]) -> bool:
    """Infer reply status from last-verb metadata."""
    try:
        last_verb = getattr(mail_item, "LastVerbExecuted", None)
    except Exception:
        last_verb = None
    if isinstance(last_verb, int) and last_verb in LAST_VERB_REPLY_CODES:
        last_time = to_python_datetime(getattr(mail_item, "LastVerbExecutionTime", None))
        if not baseline or not last_time or last_time >= baseline:
            return True

    accessor = getattr(mail_item, "PropertyAccessor", None)
    if accessor:
        try:
            verb_value = accessor.GetProperty(PR_LAST_VERB_EXECUTED)
            if isinstance(verb_value, int) and verb_value in LAST_VERB_REPLY_CODES:
                time_value = accessor.GetProperty(PR_LAST_VERB_EXECUTION_TIME)
                time_dt = to_python_datetime(time_value)
                if not baseline or not time_dt or time_dt >= baseline:
                    return True
        except Exception:
            pass
    return False


# ---------------------------------------------------------------------------
# Email formatting and retrieval
# ---------------------------------------------------------------------------
def format_email(mail_item) -> Dict[str, Any]:
    """Format an Outlook mail item into a structured dictionary."""
    try:
        recipients_by_type = extract_recipients(mail_item)
        all_recipients = (
            recipients_by_type["to"]
            + recipients_by_type["cc"]
            + recipients_by_type["bcc"]
        )

        body_content = getattr(mail_item, "Body", "") or ""
        preview = build_body_preview(body_content)

        received_iso = None
        received_display = None
        if hasattr(mail_item, "ReceivedTime") and mail_item.ReceivedTime:
            received_dt = to_python_datetime(mail_item.ReceivedTime)
            if received_dt:
                received_display = received_dt.strftime("%Y-%m-%d %H:%M:%S")
                received_iso = received_dt.strftime("%Y-%m-%dT%H:%M:%S")
            else:
                received_display = str(mail_item.ReceivedTime)
                received_iso = received_display

        sent_iso = None
        sent_display = None
        if hasattr(mail_item, "SentOn") and mail_item.SentOn:
            sent_dt = to_python_datetime(mail_item.SentOn)
            if sent_dt:
                sent_display = sent_dt.strftime("%Y-%m-%d %H:%M:%S")
                sent_iso = sent_dt.strftime("%Y-%m-%dT%H:%M:%S")
            else:
                sent_display = str(mail_item.SentOn)
                sent_iso = sent_display

        last_modified_iso = None
        last_modified_display = None
        if hasattr(mail_item, "LastModificationTime") and mail_item.LastModificationTime:
            last_dt = to_python_datetime(mail_item.LastModificationTime)
            if last_dt:
                last_modified_display = last_dt.strftime("%Y-%m-%d %H:%M:%S")
                last_modified_iso = last_dt.strftime("%Y-%m-%dT%H:%M:%S")
            else:
                last_modified_display = str(mail_item.LastModificationTime)
                last_modified_iso = last_modified_display

        has_attachments = False
        attachment_count = 0
        attachment_names: List[str] = []
        if hasattr(mail_item, "Attachments"):
            try:
                attachment_count = mail_item.Attachments.Count
                has_attachments = attachment_count > 0
                if has_attachments:
                    attachment_names = extract_attachment_names(mail_item)
            except Exception:
                attachment_count = 0
                has_attachments = False

        importance_value = mail_item.Importance if hasattr(mail_item, "Importance") else None
        importance_label = describe_importance(importance_value)

        email_data = {
            "id": mail_item.EntryID,
            "conversation_id": mail_item.ConversationID if hasattr(mail_item, "ConversationID") else None,
            "subject": mail_item.Subject,
            "sender": mail_item.SenderName,
            "sender_email": mail_item.SenderEmailAddress,
            "received_time": received_display,
            "received_iso": received_iso,
            "sent_time": sent_display,
            "sent_iso": sent_iso,
            "last_modified_time": last_modified_display,
            "last_modified_iso": last_modified_iso,
            "recipients": all_recipients,
            "to_recipients": recipients_by_type["to"],
            "cc_recipients": recipients_by_type["cc"],
            "bcc_recipients": recipients_by_type["bcc"],
            "body": body_content,
            "preview": preview,
            "has_attachments": has_attachments,
            "attachment_count": attachment_count,
            "attachment_names": attachment_names,
            "unread": mail_item.UnRead if hasattr(mail_item, "UnRead") else False,
            "importance": importance_value if importance_value is not None else 1,
            "importance_label": importance_label,
            "categories": mail_item.Categories if hasattr(mail_item, "Categories") else "",
            "folder_path": safe_folder_path(mail_item),
            "message_class": getattr(mail_item, "MessageClass", ""),
        }
        return email_data
    except Exception as exc:
        logger.exception("Impossibile formattare il messaggio con EntryID=%s.", getattr(mail_item, "EntryID", "Sconosciuto"))
        raise Exception(f"Impossibile formattare il messaggio: {exc}")


def get_emails_from_folder(folder, days: int, search_term: Optional[str] = None):
    """Get emails from a folder with optional search filter."""
    emails_list: List[Dict[str, Any]] = []
    now = datetime.datetime.now()
    threshold_date = now - datetime.timedelta(days=days)
    term_groups: List[List[str]] = []
    if search_term:
        for raw_term in filter(None, (chunk.strip() for chunk in search_term.split(" OR "))):
            tokens = [token for token in raw_term.lower().split() if token]
            if tokens:
                term_groups.append(tokens)

    try:
        try:
            default_item_type = getattr(folder, "DefaultItemType", None)
        except Exception:
            default_item_type = None
        if default_item_type not in (None, 0):  # 0 == olMailItem
            logger.debug(
                "Cartella '%s' ignorata (tipo elemento predefinito=%s).",
                getattr(folder, "Name", str(folder)),
                default_item_type,
            )
            return []

        folder_items = folder.Items
        folder_items.Sort("[ReceivedTime]", True)
        logger.info(
            "Raccolta email dalla cartella '%s' con giorni=%s termine=%s.",
            getattr(folder, "Name", str(folder)),
            days,
            search_term,
        )

        def _matches_search_groups(email_data: Dict[str, Any]) -> bool:
            if not term_groups:
                return True
            haystacks = [
                email_data.get("subject") or "",
                email_data.get("sender") or "",
                email_data.get("sender_email") or "",
                " ".join(email_data.get("recipients") or []),
                " ".join(email_data.get("to_recipients") or []),
                " ".join(email_data.get("cc_recipients") or []),
                " ".join(email_data.get("bcc_recipients") or []),
                email_data.get("body") or "",
                email_data.get("preview") or "",
            ]
            normalized = [value.lower() for value in haystacks if isinstance(value, str)]
            for tokens in term_groups:
                if tokens and all(any(token in field for field in normalized) for token in tokens):
                    return True
            return False

        count = 0
        for item in folder_items:
            try:
                if hasattr(item, "ReceivedTime") and item.ReceivedTime:
                    received_time = item.ReceivedTime.replace(tzinfo=None)
                    if received_time < threshold_date:
                        break

                    if term_groups:
                        pre_fields: List[str] = []
                        try:
                            subject = getattr(item, "Subject", "")
                            sender_name = getattr(item, "SenderName", "")
                            sender_email = getattr(item, "SenderEmailAddress", "")
                            body = getattr(item, "Body", "")
                            pre_fields.extend(
                                [
                                    subject if isinstance(subject, str) else "",
                                    sender_name if isinstance(sender_name, str) else "",
                                    sender_email if isinstance(sender_email, str) else "",
                                    body if isinstance(body, str) else "",
                                ]
                            )
                        except Exception:
                            pass
                        try:
                            recipient_strings: List[str] = []
                            for recipient in getattr(item, "Recipients", []):
                                name = getattr(recipient, "Name", None)
                                address = getattr(recipient, "Address", None)
                                if name and address:
                                    recipient_strings.append(f"{name} <{address}>")
                                elif name:
                                    recipient_strings.append(str(name))
                                elif address:
                                    recipient_strings.append(str(address))
                            if recipient_strings:
                                pre_fields.append(" ".join(recipient_strings))
                        except Exception:
                            pass
                        normalized_item_fields = [
                            value.lower() for value in pre_fields if isinstance(value, str) and value
                        ]
                        if normalized_item_fields:
                            matches_term = False
                            for tokens in term_groups:
                                if tokens and all(
                                    any(token in field for field in normalized_item_fields) for token in tokens
                                ):
                                    matches_term = True
                                    break
                            if not matches_term:
                                continue

                    email_data = format_email(item)
                    if search_term and not _matches_search_groups(email_data):
                        continue
                    emails_list.append(email_data)
                    count += 1
                    if count >= MAX_EMAIL_SCAN_PER_FOLDER:
                        break
            except Exception as exc:
                logger.warning("Errore durante l'elaborazione di un messaggio: %s", exc)
                continue
    except Exception:
        logger.exception("Errore nel recupero dei messaggi dalla cartella '%s'.", getattr(folder, "Name", str(folder)))

    return emails_list


def resolve_additional_folders(namespace, folder_names: Optional[Iterable[str]]) -> List:
    """Resolve extra folder names to Outlook folder objects."""
    resolved: List = []
    if not folder_names:
        return resolved

    seen_paths: Set[str] = set()
    for name in folder_names:
        if not name:
            continue
        try:
            folder = folder_service.get_folder_by_name(namespace, name)
            if not folder:
                logger.warning("Cartella aggiuntiva '%s' non trovata per la ricerca della conversazione.", name)
                continue
            try:
                path = folder.FolderPath
            except Exception:
                path = str(folder)
            if path in seen_paths:
                continue
            seen_paths.add(path)
            resolved.append(folder)
        except Exception:
            logger.exception("Errore nel recupero della cartella aggiuntiva '%s'.", name)
    return resolved


def get_all_mail_folders(namespace) -> List:
    """Return a flat list of all accessible mail folders prioritizing the inbox tree."""
    folders: List = []
    visited_paths: Set[str] = set()

    def enqueue(folder) -> None:
        try:
            path = folder.FolderPath
        except Exception:
            path = str(folder)
        if path in visited_paths:
            return
        visited_paths.add(path)
        folders.append(folder)
        try:
            for subfolder in folder.Folders:
                enqueue(subfolder)
        except Exception:
            return

    try:
        inbox = namespace.GetDefaultFolder(6)  # Inbox
        enqueue(inbox)
    except Exception:
        logger.warning("Impossibile accedere alla Posta in arrivo predefinita durante la scansione globale.")
        inbox = None

    try:
        for root_folder in namespace.Folders:
            if inbox and root_folder.EntryID == inbox.EntryID:
                continue
            enqueue(root_folder)
    except Exception:
        logger.warning("Impossibile enumerare le cartelle principali dell'account Outlook.")

    logger.debug("Rilevate %s cartelle complessive per la scansione globale (inbox prioritaria).", len(folders))
    return folders


def collect_emails_across_folders(
    folders: Sequence,
    days: int,
    search_term: Optional[str] = None,
    target_total: Optional[int] = None,
) -> List[Dict[str, Any]]:
    """Aggregate emails from multiple folders into a single newest-first list."""
    aggregated: Dict[str, Dict[str, Any]] = {}
    total_folders = max(len(folders), 1)
    base_limit = max(150 // total_folders, 5)
    max_per_folder = MAX_EMAIL_SCAN_PER_FOLDER if search_term else base_limit
    if target_total:
        per_folder_goal = max(1, (target_total + total_folders - 1) // total_folders)
        max_per_folder = max(max_per_folder, per_folder_goal)
    for folder in folders:
        try:
            folder_emails = get_emails_from_folder(folder, days, search_term)
        except Exception:
            logger.debug("Cartella ignorata durante la raccolta globale: %s", getattr(folder, "FolderPath", folder))
            continue

        limited_emails = folder_emails if search_term else folder_emails[:max_per_folder]
        if not search_term and len(folder_emails) > len(limited_emails):
            logger.debug(
                "Cartella '%s': limitati %s messaggi su %s per contenere la scansione globale.",
                getattr(folder, "Name", str(folder)),
                len(limited_emails),
                len(folder_emails),
            )
        for email in limited_emails:
            email_id = email.get("id")
            if not email_id:
                continue
            if email_id not in aggregated:
                aggregated[email_id] = email
            if target_total and len(aggregated) >= target_total:
                break

        if target_total and len(aggregated) >= target_total:
            break

    sorted_emails = sorted(
        aggregated.values(),
        key=lambda item: item.get("received_iso") or item.get("received_time") or "",
        reverse=True,
    )
    logger.info(
        "Raccolti %s messaggi totali attraversando %s cartelle (limite per cartella=%s%s).",
        len(sorted_emails),
        len(folders),
        max_per_folder,
        f", target={target_total}" if target_total else "",
    )
    return sorted_emails


# ---------------------------------------------------------------------------
# Conversation helpers
# ---------------------------------------------------------------------------
def get_related_conversation_emails(
    namespace,
    mail_item,
    max_items: int = 5,
    lookback_days: int = 30,
    include_sent: bool = True,
    additional_folders: Optional[Iterable[str]] = None,
) -> List[Dict[str, Any]]:
    """Collect other emails from the same conversation to build context."""
    conversation_id = getattr(mail_item, "ConversationID", None)
    if not conversation_id:
        logger.debug("Nessun ID conversazione disponibile: ricerca conversazione ignorata.")
        return []

    now = datetime.datetime.now()
    threshold_date = now - datetime.timedelta(days=lookback_days)
    seen_ids = {mail_item.EntryID}
    related_entries: List[Tuple[Optional[datetime.datetime], Dict[str, Any]]] = []

    potential_folders = []
    parent_folder = getattr(mail_item, "Parent", None)
    if parent_folder:
        potential_folders.append(parent_folder)

    default_folder_ids = [6]  # Inbox
    if include_sent:
        default_folder_ids.append(5)  # Sent Items
    for folder_id in default_folder_ids:
        try:
            folder = namespace.GetDefaultFolder(folder_id)
            potential_folders.append(folder)
        except Exception:
            continue

    for extra_folder in resolve_additional_folders(namespace, additional_folders):
        potential_folders.append(extra_folder)

    folders_to_scan = []
    seen_paths: Set[str] = set()
    for folder in potential_folders:
        try:
            folder_path = folder.FolderPath
        except Exception:
            folder_path = str(folder)
        if folder_path in seen_paths:
            continue
        seen_paths.add(folder_path)
        folders_to_scan.append(folder)

    for folder in folders_to_scan:
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
        except Exception:
            logger.warning(
                "Impossibile scorrere gli elementi della cartella '%s' durante la ricerca della conversazione.",
                getattr(folder, "Name", folder),
            )
            continue

        manual_filter = False
        candidate_items = items
        try:
            filter_query = f"[ConversationID] = '{conversation_id}'"
            candidate_items = items.Restrict(filter_query)
        except Exception:
            logger.debug(
                "Filtro conversazione SQL non disponibile nella cartella '%s', uso filtraggio manuale.",
                getattr(folder, "Name", folder),
            )
            manual_filter = True

        scanned = 0
        max_scan = max(max_items * 25, 200)
        for item in candidate_items:
            scanned += 1
            if scanned > max_scan:
                break

            try:
                if manual_filter and getattr(item, "ConversationID", None) != conversation_id:
                    continue

                if not hasattr(item, "EntryID") or item.EntryID in seen_ids:
                    continue

                received_dt = None
                if hasattr(item, "ReceivedTime") and item.ReceivedTime:
                    try:
                        received_dt = datetime.datetime(
                            item.ReceivedTime.year,
                            item.ReceivedTime.month,
                            item.ReceivedTime.day,
                            item.ReceivedTime.hour,
                            item.ReceivedTime.minute,
                            item.ReceivedTime.second,
                        )
                    except Exception:
                        try:
                            received_dt = datetime.datetime.strptime(
                                item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
                                "%Y-%m-%d %H:%M:%S",
                            )
                        except Exception:
                            received_dt = None

                if received_dt and received_dt < threshold_date:
                    break

                email_data = format_email(item)
                related_entries.append((received_dt, email_data))
                seen_ids.add(item.EntryID)

                if len(related_entries) >= max_items:
                    break
            except Exception:
                logger.debug(
                    "Messaggio correlato ignorato a causa di un errore di elaborazione.",
                    exc_info=True,
                )
                continue

        if len(related_entries) >= max_items:
            break

    related_entries.sort(
        key=lambda entry: entry[0] if entry[0] else datetime.datetime.min,
        reverse=True,
    )
    return [entry[1] for entry in related_entries]


def email_has_user_reply(
    namespace,
    email_data: Dict[str, Any],
    user_addresses: Set[str],
    conversation_limit: int,
    lookback_days: int,
) -> bool:
    """Determine whether the user has already replied within a conversation."""
    result, _, _ = email_has_user_reply_with_context(
        namespace=namespace,
        email_data=email_data,
        user_addresses=user_addresses,
        conversation_limit=conversation_limit,
        lookback_days=lookback_days,
        collect_related=False,
    )
    return result


def email_has_user_reply_with_context(
    namespace,
    email_data: Dict[str, Any],
    user_addresses: Set[str],
    conversation_limit: int,
    lookback_days: int,
    collect_related: bool,
) -> Tuple[bool, Optional[List[Dict[str, Any]]], Any]:
    """Determine whether the user has already replied within a conversation."""
    if not email_data:
        return False, None, None

    normalized_user_addresses = (
        {
            addr
            for addr in (normalize_email_address(addr) for addr in user_addresses)
            if addr
        }
        if user_addresses
        else set()
    )

    baseline_dt = _extract_best_timestamp(email_data)
    mail_item = None
    captured_related: Optional[List[Dict[str, Any]]] = None
    try:
        mail_item = namespace.GetItemFromID(email_data.get("id"))
    except Exception:
        logger.debug(
            "Impossibile recuperare il messaggio %s per il controllo delle risposte.",
            email_data.get("id"),
            exc_info=True,
        )

    if normalized_user_addresses:
        related_entries: List[Dict[str, Any]] = []
        if mail_item:
            try:
                related_entries = get_related_conversation_emails(
                    namespace=namespace,
                    mail_item=mail_item,
                    max_items=conversation_limit,
                    lookback_days=lookback_days,
                    include_sent=True,
                    additional_folders=None,
                )
                if collect_related:
                    captured_related = related_entries
            except Exception:
                logger.debug(
                    "Errore durante la ricerca dei messaggi correlati per %s.",
                    email_data.get("id"),
                    exc_info=True,
                )

        for related in related_entries:
            msg_class = (related.get("message_class") or "").lower()
            if msg_class and not msg_class.startswith("ipm.note"):
                continue
            sender_email = normalize_email_address(related.get("sender_email")) or normalize_email_address(
                related.get("sender")
            )
            if not sender_email or sender_email not in normalized_user_addresses:
                continue
            related_dt = _extract_best_timestamp(related)
            if not baseline_dt or not related_dt or related_dt >= baseline_dt:
                return True, None, None

    if mail_item and mail_item_marked_replied(mail_item, baseline_dt):
        return True, None, None

    return False, captured_related, mail_item


def build_conversation_outline(
    namespace,
    email_data: Dict[str, Any],
    lookback_days: int,
    max_items: int = 4,
    preloaded_entries: Optional[List[Dict[str, Any]]] = None,
    mail_item: Any = None,
) -> Optional[str]:
    """Create a compact summary of the most recent conversation messages."""
    if max_items < 1:
        return None

    mail_id = email_data.get("id")
    if not mail_id and not mail_item:
        return None

    local_mail_item = mail_item
    if local_mail_item is None and mail_id:
        try:
            local_mail_item = namespace.GetItemFromID(mail_id)
        except Exception:
            logger.debug(
                "Impossibile recuperare il messaggio %s per costruire il riepilogo conversazione.",
                mail_id,
                exc_info=True,
            )
            return None

    if preloaded_entries is not None:
        related_entries = preloaded_entries
    else:
        try:
            related_entries = get_related_conversation_emails(
                namespace=namespace,
                mail_item=local_mail_item,
                max_items=max(1, max_items - 1),
                lookback_days=lookback_days,
                include_sent=True,
                additional_folders=None,
            )
        except Exception:
            logger.debug(
                "Errore durante la raccolta dei messaggi correlati per %s.",
                mail_id,
                exc_info=True,
            )
            related_entries = []

    timeline: List[Tuple[Optional[datetime.datetime], Dict[str, Any], bool]] = []
    main_dt = _extract_best_timestamp(email_data)
    timeline.append((main_dt, email_data, True))

    for entry in related_entries:
        timeline.append((_extract_best_timestamp(entry), entry, False))

    filtered = [
        (dt, entry, is_focus)
        for dt, entry, is_focus in timeline
        if entry
    ]
    if not filtered:
        return None

    filtered.sort(
        key=lambda value: value[0] if value[0] else datetime.datetime.min,
        reverse=True,
    )

    lines: List[str] = []
    for dt, entry, is_focus in filtered[:max_items]:
        msg_class = (entry.get("message_class") or "").lower()
        if msg_class and not msg_class.startswith("ipm.note"):
            if not is_focus:
                continue
        timestamp = (
            dt.strftime("%Y-%m-%d %H:%M")
            if isinstance(dt, datetime.datetime)
            else entry.get("received_time")
            or entry.get("sent_time")
            or entry.get("last_modified_time")
            or "Sconosciuto"
        )
        sender = entry.get("sender", "Sconosciuto")
        subject = entry.get("subject", "(Senza oggetto)")
        prefix = ">>" if is_focus else "- "
        preview = entry.get("preview") or build_body_preview(entry.get("body"), 160)
        preview_line = ""
        if preview:
            preview_line = f"\n   Anteprima: {preview}"

        lines.append(f"{prefix} {timestamp} -> {sender}: {subject}{preview_line}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Presentation helpers
# ---------------------------------------------------------------------------
def present_email_listing(
    emails: Sequence[Dict[str, Any]],
    folder_display: str,
    days: int,
    max_results: int,
    include_preview: bool,
    log_context: str,
    search_term: Optional[str] = None,
    focus_on_recipients: bool = False,
    offset: int = 0,
) -> str:
    """Common presenter for listing emails and caching them."""
    clear_email_cache()

    if not emails:
        if search_term:
            message = (
                f"Nessun messaggio corrispondente a '{search_term}' trovato in {folder_display} "
                f"negli ultimi {days} giorni."
            )
        else:
            message = f"Nessun messaggio trovato in {folder_display} negli ultimi {days} giorni."
        logger.info(
            "%s: nessun messaggio (termine=%s cartella=%s giorni=%s).",
            log_context,
            search_term,
            folder_display,
            days,
        )
        return message

    try:
        start_index = int(offset)
    except (TypeError, ValueError):
        start_index = 0
    if start_index < 0:
        start_index = 0

    total_count = len(emails)
    if start_index >= total_count:
        logger.info(
            "%s: offset %s fuori intervallo (messaggi totali=%s).",
            log_context,
            start_index,
            total_count,
        )
        return (
            f"Nessun messaggio disponibile partendo dall'offset {start_index}. "
            f"La lista corrente contiene {total_count} elementi."
        )

    visible_emails = list(emails)[start_index:start_index + max_results]
    visible_count = len(visible_emails)
    first_position = start_index + 1
    last_position = start_index + visible_count

    if search_term:
        if total_count > visible_count:
            header = (
                f"Trovati {total_count} messaggi che corrispondono a '{search_term}' in {folder_display} "
                f"negli ultimi {days} giorni. Mostro i risultati {first_position}-{last_position}."
            )
        else:
            header = (
                f"Trovati {visible_count} messaggi che corrispondono a '{search_term}' in {folder_display} "
                f"negli ultimi {days} giorni."
            )
    else:
        if total_count > visible_count:
            header = (
                f"Trovati {total_count} messaggi in {folder_display} negli ultimi {days} giorni. "
                f"Mostro i risultati {first_position}-{last_position}."
            )
        else:
            header = f"Trovati {visible_count} messaggi in {folder_display} negli ultimi {days} giorni."

    logger.info(
        "%s: restituiti %s messaggi su %s (termine=%s cartella=%s offset=%s).",
        log_context,
        visible_count,
        total_count,
        search_term,
        folder_display,
        start_index,
    )

    result = header + "\n\n"

    for idx, email in enumerate(visible_emails, 1):
        email_cache[idx] = email

        folder_path = email.get("folder_path") or folder_display
        importance_label = email.get("importance_label") or describe_importance(email.get("importance"))
        trimmed_conv = trim_conversation_id(email.get("conversation_id"))
        attachments_line = None
        if email.get("attachment_names"):
            attachments_line = f"Nomi allegati: {', '.join(email['attachment_names'])}"

        result += f"Messaggio #{idx}\n"
        result += f"Oggetto: {email.get('subject', '(Senza oggetto)')}\n"
        if focus_on_recipients and email.get("to_recipients"):
            result += f"A: {', '.join(email['to_recipients'])}\n"
        result += f"Da: {email.get('sender', 'Sconosciuto')} <{email.get('sender_email', '')}>\n"
        result += f"Ricevuto: {email.get('received_time', 'Sconosciuto')}\n"
        result += f"Cartella: {folder_path}\n"
        result += f"Importanza: {importance_label}\n"
        result += f"Stato lettura: {format_read_status(email.get('unread'))}\n"
        result += f"Allegati: {format_yes_no(email.get('has_attachments'))}\n"
        if attachments_line:
            result += attachments_line + "\n"
        if include_preview and email.get("preview"):
            result += f"Anteprima: {email['preview']}\n"
        if email.get("categories"):
            result += f"Categorie: {email['categories']}\n"
        if trimmed_conv:
            result += f"ID conversazione: {trimmed_conv}\n"
        result += "\n"

    return result.rstrip()


def apply_categories_to_item(mail_item, categories: Sequence[str], overwrite: bool, append: bool) -> List[str]:
    """Apply Outlook categories to a mail item respecting overwrite/append semantics."""
    if overwrite and append:
        raise ValueError("Impossibile usare contemporaneamente overwrite e append.")

    normalized_new = [
        str(category).strip()
        for category in categories
        if str(category).strip()
    ]

    existing_raw = ""
    try:
        existing_raw = getattr(mail_item, "Categories", "") or ""
    except Exception:
        existing_raw = ""

    existing_list = [cat.strip() for cat in existing_raw.split(";") if cat.strip()]

    if overwrite:
        final_categories = list(dict.fromkeys(normalized_new))
    else:
        final_categories = existing_list.copy()
        if final_categories and not append:
            final_categories = []
        for cat in normalized_new:
            if cat not in final_categories:
                final_categories.append(cat)

    final_str = "; ".join(final_categories)
    try:
        mail_item.Categories = final_str
        mail_item.Save()
    except Exception as exc:
        logger.exception("Applicazione categorie fallita.")
        raise ValueError(f"Impossibile applicare le categorie: {exc}") from exc

    return final_categories


# ---------------------------------------------------------------------------
# High-level context helper
# ---------------------------------------------------------------------------
def get_email_context(
    email_number: int,
    include_thread: bool = True,
    thread_limit: int = 5,
    lookback_days: int = 30,
    include_sent: bool = True,
    additional_folders: Optional[Iterable[str]] = None,
) -> str:
    """Provide conversation-aware context for a previously listed email."""
    try:
        if not email_cache:
            logger.warning("get_email_context chiamato ma la cache e vuota.")
            return "Errore: nessun elenco messaggi attivo. Chiedimi prima di mostrare le email e poi ripeti la richiesta."

        if email_number not in email_cache:
            logger.warning("Messaggio numero %s non presente in cache per get_email_context.", email_number)
            return f"Errore: il messaggio #{email_number} non e presente nell'elenco corrente."

        if not isinstance(thread_limit, int) or thread_limit < 1:
            logger.warning("Valore thread_limit non valido per get_email_context: %s", thread_limit)
            return "Errore: 'thread_limit' deve essere un intero positivo."

        if not isinstance(lookback_days, int) or lookback_days < 1 or lookback_days > 180:
            logger.warning("Valore lookback_days non valido per get_email_context: %s", lookback_days)
            return "Errore: 'lookback_days' deve essere un intero tra 1 e 180."

        extra_folders: Optional[List[str]] = None
        if additional_folders is not None:
            if isinstance(additional_folders, str):
                extra_folders = [additional_folders]
            else:
                try:
                    extra_folders = [str(name) for name in additional_folders if name]
                except TypeError:
                    logger.warning("Valore 'additional_folders' non valido: %s", additional_folders)
                    return "Errore: 'additional_folders' deve essere una lista di nomi di cartella."

        include_thread_bool = coerce_bool(include_thread)
        include_sent_bool = coerce_bool(include_sent)
        logger.info(
            "get_email_context chiamato per messaggio #%s include_thread=%s thread_limit=%s lookback_days=%s include_sent=%s cartelle_extra=%s",
            email_number,
            include_thread_bool,
            thread_limit,
            lookback_days,
            include_sent_bool,
            extra_folders,
        )

        email_data = email_cache[email_number]
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_data["id"])
        if not email:
            return f"Errore: il messaggio #{email_number} non puo essere recuperato da Outlook."

        importance_label = email_data.get("importance_label") or describe_importance(email_data.get("importance"))

        attachment_names = list(email_data.get("attachment_names") or [])
        try:
            if hasattr(email, "Attachments") and email.Attachments.Count > 0:
                for i in range(1, email.Attachments.Count + 1):
                    try:
                        file_name = email.Attachments(i).FileName
                        if file_name and file_name not in attachment_names:
                            attachment_names.append(file_name)
                    except Exception:
                        continue
        except Exception:
            pass

        participants = set()
        sender_display = f"{email_data.get('sender', 'Sconosciuto')} <{email_data.get('sender_email', '')}>".strip()
        if sender_display:
            participants.add(sender_display)
        for recipient in email_data.get("recipients", []):
            if recipient:
                participants.add(recipient)

        context_lines = [
            f"Contesto per il messaggio #{email_number}",
            "",
            f"Oggetto: {email_data.get('subject', 'Oggetto sconosciuto')}",
            f"Da: {email_data.get('sender', 'Mittente sconosciuto')} <{email_data.get('sender_email', '')}>",
        ]

        if email_data.get("to_recipients"):
            context_lines.append(f"A: {', '.join(email_data['to_recipients'])}")
        if email_data.get("cc_recipients"):
            context_lines.append(f"Cc: {', '.join(email_data['cc_recipients'])}")
        if email_data.get("bcc_recipients"):
            context_lines.append(f"Ccn: {', '.join(email_data['bcc_recipients'])}")

        context_lines.extend(
            [
                f"Ricevuto: {email_data.get('received_time', 'Sconosciuto')}",
                f"Cartella: {email_data.get('folder_path', 'Sconosciuta')}",
                f"Importanza: {importance_label}",
                f"Stato lettura: {format_read_status(email_data.get('unread'))}",
            ]
        )

        if email_data.get("categories"):
            context_lines.append(f"Categorie: {email_data['categories']}")
        if email_data.get("conversation_id"):
            trimmed_conv = trim_conversation_id(email_data["conversation_id"], max_chars=32)
            conv_line = f"ID conversazione: {trimmed_conv}" if trimmed_conv else "ID conversazione: (Non disponibile)"
            if trimmed_conv and trimmed_conv.endswith("..."):
                conv_line += " (troncato)"
            context_lines.append(conv_line)

        if participants:
            context_lines.append(f"Partecipanti coinvolti: {', '.join(sorted(participants))}")

        if email_data.get("preview"):
            context_lines.append(f"Anteprima corpo: {email_data['preview']}")

        if attachment_names:
            context_lines.append(f"Allegati: {', '.join(attachment_names)}")

        body_content = email_data.get("body", "")
        if body_content and len(body_content) > 4000:
            truncated_body = body_content[:4000].rstrip() + "\n[Corpo troncato per brevita]"
        else:
            truncated_body = body_content or "(Nessun contenuto)"

        context_lines.append("")
        context_lines.append("Corpo del messaggio corrente:")
        context_lines.append(truncated_body)
        context_lines.append("Per leggere il corpo completo chiedimi i dettagli di questo messaggio e li recuperero subito.")

        if include_thread_bool:
            context_lines.append("")
            context_lines.append("Messaggi correlati della conversazione:")
            related_emails = get_related_conversation_emails(
                namespace=namespace,
                mail_item=email,
                max_items=thread_limit,
                lookback_days=lookback_days,
                include_sent=include_sent_bool,
                additional_folders=extra_folders,
            )

            if not related_emails:
                context_lines.append("- Nessun messaggio aggiuntivo trovato nell'intervallo indicato.")
            else:
                for idx, related in enumerate(related_emails, 1):
                    summary_header = (
                        f"{idx}. {related.get('received_time', 'Orario sconosciuto')} | "
                        f"{related.get('sender', 'Mittente sconosciuto')} | "
                        f"{related.get('folder_path', 'Cartella sconosciuta')}"
                    )
                    context_lines.append(summary_header)
                    if related.get("subject") and related["subject"] != email_data.get("subject"):
                        context_lines.append(f"   Oggetto: {related['subject']}")
                    if related.get("preview"):
                        context_lines.append(f"   Anteprima: {related['preview']}")
                    if related.get("has_attachments"):
                        context_lines.append(
                            f"   Allegati: {related.get('attachment_count', 0)} file."
                        )
                        if related.get("attachment_names"):
                            context_lines.append(f"   Nomi allegati: {', '.join(related['attachment_names'])}")

        context_lines.append("")
        context_lines.append(
            "Suggerimento: fammi sapere se vuoi rispondere o iniziare un nuovo thread e preparero la bozza corrispondente."
        )

        return "\n".join(context_lines)

    except Exception as exc:
        logger.exception("Errore nel recupero del contesto per il messaggio #%s.", email_number)
        return f"Errore durante il recupero del contesto del messaggio: {exc}"


# ---------------------------------------------------------------------------
# Backwards-compatible aliases (legacy underscore-prefixed names)
# ---------------------------------------------------------------------------
_resolve_mail_item = resolve_mail_item
_update_cached_email = update_cached_email
_normalize_email_address = normalize_email_address
_extract_email_domain = extract_email_domain
_derive_sender_email = derive_sender_email
_ensure_domain_folder_structure = ensure_domain_folder_structure
_collect_user_addresses = collect_user_addresses
_mail_item_marked_replied = mail_item_marked_replied
_present_email_listing = present_email_listing
_email_has_user_reply = email_has_user_reply
_email_has_user_reply_with_context = email_has_user_reply_with_context
_build_conversation_outline = build_conversation_outline
