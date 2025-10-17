import argparse
import datetime
import os
from typing import Any, Dict, List, Optional, Set, Tuple

from mcp.server.fastmcp import FastMCP, Context
from mcp.server.fastmcp.exceptions import ToolError

from outlook_mcp import (
    ATTACHMENT_NAME_PREVIEW_MAX,
    BODY_PREVIEW_MAX_CHARS,
    CONVERSATION_ID_PREVIEW_MAX,
    DEFAULT_CONVERSATION_SAMPLE_LIMIT,
    DEFAULT_DOMAIN_ROOT_NAME,
    DEFAULT_DOMAIN_SUBFOLDERS,
    DEFAULT_MAX_RESULTS,
    LAST_VERB_REPLY_CODES,
    MAX_CONVERSATION_LOOKBACK_DAYS,
    MAX_DAYS,
    MAX_EMAIL_SCAN_PER_FOLDER,
    MAX_EVENT_LOOKAHEAD_DAYS,
    PENDING_SCAN_MULTIPLIER,
    PR_LAST_VERB_EXECUTED,
    PR_LAST_VERB_EXECUTION_TIME,
    calendar_cache,
    clear_calendar_cache,
    clear_email_cache,
    connect_to_outlook,
    email_cache,
    logger,
)
from outlook_mcp.utils import (
    build_body_preview,
    coerce_bool,
    describe_item_type,
    ensure_int_list,
    ensure_string_list,
    extract_attachment_names,
    extract_recipients,
    normalize_folder_path,
    parse_item_type_hint,
    safe_child_count,
    safe_entry_id,
    safe_filename,
    safe_folder_path,
    safe_folder_size,
    safe_store_id,
    safe_total_count,
    safe_unread_count,
    shorten_identifier,
    to_python_datetime,
    trim_conversation_id,
)
from outlook_mcp import folders as folder_service

try:
    from fastapi import Body, FastAPI, HTTPException
    from pydantic import BaseModel, Field
    import uvicorn
except ImportError:  # Optional dependencies loaded only for HTTP mode
    Body = None  # type: ignore[assignment]
    FastAPI = None  # type: ignore[assignment]
    HTTPException = None  # type: ignore[assignment]
    BaseModel = None  # type: ignore[assignment]
    Field = None  # type: ignore[assignment]
    uvicorn = None  # type: ignore[assignment]


# Initialize FastMCP server
mcp = FastMCP("outlook-assistant", host="0.0.0.0", port=8000)

# Backwards-compatible aliases for shared helpers
get_folder_by_name = folder_service.get_folder_by_name
get_folder_by_path = folder_service.get_folder_by_path
resolve_folder = folder_service.resolve_folder

def _resolve_mail_item(namespace, *, email_number: Optional[int] = None, message_id: Optional[str] = None):
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
    except Exception as exc:
        raise ToolError(f"Impossibile recuperare il messaggio con ID '{message_id}': {exc}") from exc

    if not mail_item:
        raise ToolError("Outlook ha restituito un elemento vuoto per l'ID specificato.")

    return cached_entry, mail_item

def _update_cached_email(email_number: Optional[int], **updates: Any) -> None:
    """Apply in-place updates to the cached representation of an email."""
    if email_number is None:
        return
    if not email_cache or email_number not in email_cache:
        return
    email_cache[email_number].update({key: value for key, value in updates.items() if value is not None})

def _normalize_email_address(value: Optional[str]) -> Optional[str]:
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

def _extract_email_domain(address: Optional[str]) -> Optional[str]:
    """Return email domain portion."""
    normalized = _normalize_email_address(address)
    if not normalized or "@" not in normalized:
        return None
    return normalized.split("@", 1)[1]

def _derive_sender_email(entry: Dict[str, Any]) -> Optional[str]:
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
    except Exception as exc:
        raise Exception(f"Impossibile creare la cartella '{name}': {exc}") from exc

def _ensure_domain_folder_structure(
    namespace,
    domain: str,
    root_folder_name: str = DEFAULT_DOMAIN_ROOT_NAME,
    subfolders: Optional[List[str]] = None,
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

def _parse_datetime_string(value: Optional[str]) -> Optional[datetime.datetime]:
    """Parse ISO-like or display datetime strings into datetime objects."""
    if not value:
        return None
    text = value.strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1]
    try:
        return datetime.datetime.fromisoformat(text)
    except ValueError:
        pass
    formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d",
    ]
    for fmt in formats:
        try:
            return datetime.datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None

def _extract_best_timestamp(entry: Optional[Dict[str, Any]]) -> Optional[datetime.datetime]:
    """Derive the most relevant timestamp from a formatted email dictionary."""
    if not entry:
        return None
    for key in ("received_iso", "sent_iso", "last_modified_iso"):
        dt = _parse_datetime_string(entry.get(key))
        if dt:
            return dt
    for key in ("received_time", "sent_time", "last_modified_time"):
        dt = _parse_datetime_string(entry.get(key))
        if dt:
            return dt
    return None

def _describe_importance(value: Any) -> str:
    """Map Outlook importance levels to descriptive labels."""
    importance_map = {0: "Bassa", 1: "Normale", 2: "Alta"}
    if isinstance(value, int) and value in importance_map:
        return importance_map[value]
    return str(value) if value is not None else "Sconosciuta"

def _yes_no(value: bool) -> str:
    """Return Si/No for boolean flags."""
    return "Si" if bool(value) else "No"

def _read_status(unread: bool) -> str:
    """Return localized read status."""
    return "Non letta" if unread else "Letta"

def _collect_user_addresses(namespace) -> Set[str]:
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
        for normalized in (_normalize_email_address(addr) for addr in addresses)
        if normalized
    }
    if normalized:
        logger.debug("Rilevati indirizzi utente: %s", ", ".join(sorted(normalized)))
    else:
        logger.debug("Nessun indirizzo utente normalizzato rilevato.")
    return normalized

def _mail_item_marked_replied(mail_item, baseline: Optional[datetime.datetime]) -> bool:
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

# Helper functions
def format_email(mail_item) -> Dict[str, Any]:
    """Format an Outlook mail item into a structured dictionary"""
    try:
        # Extract recipients grouped by type
        recipients_by_type = extract_recipients(mail_item)
        all_recipients = (
            recipients_by_type["to"]
            + recipients_by_type["cc"]
            + recipients_by_type["bcc"]
        )

        # Capture body and preview
        body_content = getattr(mail_item, "Body", "") or ""
        preview = build_body_preview(body_content)

        # Prepare received time representations
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

        importance_value = mail_item.Importance if hasattr(mail_item, 'Importance') else None
        importance_label = _describe_importance(importance_value)

        # Format the email data
        email_data = {
            "id": mail_item.EntryID,
            "conversation_id": mail_item.ConversationID if hasattr(mail_item, 'ConversationID') else None,
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
            "unread": mail_item.UnRead if hasattr(mail_item, 'UnRead') else False,
            "importance": importance_value if importance_value is not None else 1,
            "importance_label": importance_label,
            "categories": mail_item.Categories if hasattr(mail_item, 'Categories') else "",
            "folder_path": safe_folder_path(mail_item),
            "message_class": getattr(mail_item, "MessageClass", ""),
        }
        return email_data
    except Exception as e:
        logger.exception("Impossibile formattare il messaggio con EntryID=%s.", getattr(mail_item, "EntryID", "Sconosciuto"))
        raise Exception(f"Impossibile formattare il messaggio: {str(e)}")

def get_emails_from_folder(folder, days: int, search_term: Optional[str] = None):
    """Get emails from a folder with optional search filter"""
    emails_list = []
    
    # Calculate the date threshold
    now = datetime.datetime.now()
    threshold_date = now - datetime.timedelta(days=days)
    
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

        # Set up filtering
        folder_items = folder.Items
        folder_items.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
        logger.info(
            "Raccolta email dalla cartella '%s' con giorni=%s termine=%s.",
            getattr(folder, "Name", str(folder)),
            days,
            search_term,
        )
        
        # If we have a search term, apply it
        if search_term:
            # Handle OR operators in search term
            search_terms = [term.strip() for term in search_term.split(" OR ")]
            
            # Try to create a filter for subject, sender name or body
            try:
                # Build SQL filter with OR conditions for each search term
                sql_conditions = []
                for term in search_terms:
                    sql_conditions.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{term}%'")
                    sql_conditions.append(f"\"urn:schemas:httpmail:fromname\" LIKE '%{term}%'")
                    sql_conditions.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{term}%'")
                
                filter_term = f"@SQL=" + " OR ".join(sql_conditions)
                folder_items = folder_items.Restrict(filter_term)
            except:
                # If filtering fails, we'll do manual filtering later
                pass
        
        # Process emails
        count = 0
        for item in folder_items:
            try:
                if hasattr(item, 'ReceivedTime') and item.ReceivedTime:
                    # Convert to naive datetime for comparison
                    received_time = item.ReceivedTime.replace(tzinfo=None)
                    
                    # Skip emails older than our threshold
                    if received_time < threshold_date:
                        break
                    
                    # Manual search filter if needed
                    if search_term and folder_items == folder.Items:  # If we didn't apply filter earlier
                        # Handle OR operators in search term for manual filtering
                        search_terms = [term.strip().lower() for term in search_term.split(" OR ")]
                        
                        # Check if any of the search terms match
                        found_match = False
                        for term in search_terms:
                            if (term in item.Subject.lower() or 
                                term in item.SenderName.lower() or 
                                term in item.Body.lower()):
                                found_match = True
                                break
                        
                        if not found_match:
                            continue
                    
                    # Format and add the email
                    email_data = format_email(item)
                    emails_list.append(email_data)
                    count += 1
                    if count >= MAX_EMAIL_SCAN_PER_FOLDER:
                        break
            except Exception as e:
                logger.warning("Errore durante l'elaborazione di un messaggio: %s", str(e))
                continue
                
    except Exception as e:
        logger.exception("Errore nel recupero dei messaggi dalla cartella '%s'.", getattr(folder, "Name", str(folder)))
        
    return emails_list

def resolve_additional_folders(namespace, folder_names: Optional[List[str]]) -> List:
    """Resolve extra folder names to Outlook folder objects."""
    resolved = []
    if not folder_names:
        return resolved

    seen_paths = set()
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
    folders: List,
    days: int,
    search_term: Optional[str] = None,
    target_total: Optional[int] = None,
) -> List[Dict[str, Any]]:
    """Aggregate emails from multiple folders into a single newest-first list."""
    aggregated: Dict[str, Dict[str, Any]] = {}
    total_folders = len(folders)
    max_per_folder = max(150 // max(total_folders, 1), 5)
    for folder in folders:
        try:
            folder_emails = get_emails_from_folder(folder, days, search_term)
        except Exception:
            logger.debug("Cartella ignorata durante la raccolta globale: %s", getattr(folder, "FolderPath", folder))
            continue

        limited_emails = folder_emails[:max_per_folder]
        if len(folder_emails) > len(limited_emails):
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

    # Sort by ISO timestamp if available, otherwise fallback to received_time string
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

def get_all_calendar_folders(namespace) -> List:
    """Return every Outlook folder that stores appointments."""
    calendar_folders = []
    visited_paths = set()

    def visit(folder):
        try:
            path = folder.FolderPath
        except Exception:
            path = str(folder)
        if path in visited_paths:
            return
        visited_paths.add(path)

        try:
            default_item_type = folder.DefaultItemType
        except Exception:
            default_item_type = None

        if default_item_type == 1:  # olAppointmentItem
            calendar_folders.append(folder)

        try:
            for sub in folder.Folders:
                visit(sub)
        except Exception:
            return

    try:
        default_calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar
        visit(default_calendar)
    except Exception:
        logger.warning("Impossibile ottenere la cartella Calendario predefinita.")

    try:
        for root in namespace.Folders:
            visit(root)
    except Exception:
        logger.warning("Impossibile enumerare le radici per i calendari.")

    logger.debug("Rilevate %s cartelle calendario totali.", len(calendar_folders))
    return calendar_folders

def get_calendar_folder_by_name(namespace, calendar_name: str):
    """Find a calendar folder by its display name."""
    if not calendar_name:
        return namespace.GetDefaultFolder(9)
    target = calendar_name.lower()
    for folder in get_all_calendar_folders(namespace):
        try:
            if folder.Name.lower() == target:
                return folder
        except Exception:
            continue
    return None

def format_calendar_event(appointment) -> Dict[str, Any]:
    """Generate a structured representation of an Outlook appointment."""
    try:
        start_dt = to_python_datetime(getattr(appointment, "Start", None))
        end_dt = to_python_datetime(getattr(appointment, "End", None))

        def fmt(dt: Optional[datetime.datetime]) -> Optional[str]:
            if not dt:
                return None
            return dt.strftime("%Y-%m-%d %H:%M")

        start_iso = start_dt.strftime("%Y-%m-%dT%H:%M") if start_dt else None
        end_iso = end_dt.strftime("%Y-%m-%dT%H:%M") if end_dt else None

        required = getattr(appointment, "RequiredAttendees", "") or ""
        optional = getattr(appointment, "OptionalAttendees", "") or ""
        body = getattr(appointment, "Body", "") or ""
        preview = build_body_preview(body, max_chars=320)

        event_data = {
            "id": getattr(appointment, "EntryID", None),
            "subject": getattr(appointment, "Subject", ""),
            "location": getattr(appointment, "Location", ""),
            "start_time": fmt(start_dt),
            "end_time": fmt(end_dt),
            "start_iso": start_iso,
            "end_iso": end_iso,
            "organizer": getattr(appointment, "Organizer", ""),
            "required_attendees": required,
            "optional_attendees": optional,
            "all_day": getattr(appointment, "AllDayEvent", False),
            "is_recurring": getattr(appointment, "IsRecurring", False),
            "body": body,
            "preview": preview,
            "folder_path": safe_folder_path(appointment),
            "categories": getattr(appointment, "Categories", ""),
            "duration_minutes": int((end_dt - start_dt).total_seconds() / 60) if start_dt and end_dt else None,
        }
        return event_data
    except Exception as exc:
        raise Exception(f"Impossibile formattare l'evento di calendario: {exc}")

def get_events_from_folder(folder, days: int, search_term: Optional[str] = None) -> List[Dict[str, Any]]:
    """Retrieve upcoming events from a calendar folder."""
    now = datetime.datetime.now()
    horizon = now + datetime.timedelta(days=days)
    events: List[Dict[str, Any]] = []

    try:
        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
    except Exception:
        logger.warning("Impossibile accedere o ordinare gli eventi della cartella calendario '%s'.", getattr(folder, "Name", folder))
        return events

    search_terms = []
    if search_term:
        search_terms = [term.strip().lower() for term in search_term.split(" OR ") if term.strip()]

    scanned = 0
    max_scan = 500
    for appointment in items:
        scanned += 1
        if scanned > max_scan:
            break

        try:
            start_dt = to_python_datetime(getattr(appointment, "Start", None))
            end_dt = to_python_datetime(getattr(appointment, "End", None))

            if not start_dt:
                continue
            if end_dt and end_dt < now:
                continue
            if start_dt > horizon:
                break

            if search_terms:
                haystack = " ".join(
                    filter(
                        None,
                        [
                            getattr(appointment, "Subject", ""),
                            getattr(appointment, "Location", ""),
                            getattr(appointment, "Organizer", ""),
                            getattr(appointment, "Body", ""),
                            getattr(appointment, "RequiredAttendees", "") or "",
                            getattr(appointment, "OptionalAttendees", "") or "",
                        ],
                    )
                ).lower()
                if not any(term in haystack for term in search_terms):
                    continue

            event_data = format_calendar_event(appointment)
            events.append(event_data)
        except Exception:
            logger.debug("Evento calendario ignorato per errore di elaborazione.", exc_info=True)
            continue

    logger.info(
        "Recuperati %s eventi dalla cartella calendario '%s'.",
        len(events),
        getattr(folder, "Name", folder),
    )
    return events

def collect_events_across_calendars(
    folders: List,
    days: int,
    search_term: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Aggregate calendar events across multiple folders."""
    aggregated: Dict[str, Dict[str, Any]] = {}
    for folder in folders:
        folder_events = get_events_from_folder(folder, days, search_term)
        for event in folder_events:
            event_id = event.get("id")
            if not event_id:
                continue
            if event_id not in aggregated:
                aggregated[event_id] = event

    sorted_events = sorted(
        aggregated.values(),
        key=lambda evt: evt.get("start_iso") or "",
    )
    logger.info(
        "Aggregati %s eventi provenienti da %s cartelle calendario.",
        len(sorted_events),
        len(folders),
    )
    return sorted_events

def get_related_conversation_emails(
    namespace,
    mail_item,
    max_items: int = 5,
    lookback_days: int = 30,
    include_sent: bool = True,
    additional_folders: Optional[List[str]] = None,
):
    """Collect other emails from the same conversation to build context."""
    conversation_id = getattr(mail_item, "ConversationID", None)
    if not conversation_id:
        logger.debug("Nessun ID conversazione disponibile: ricerca conversazione ignorata.")
        return []

    now = datetime.datetime.now()
    threshold_date = now - datetime.timedelta(days=lookback_days)
    seen_ids = {mail_item.EntryID}
    related_entries = []

    potential_folders = []
    parent_folder = getattr(mail_item, "Parent", None)
    if parent_folder:
        potential_folders.append(parent_folder)

    # Add common folders that usually contain conversation items
    default_folder_ids = [6]  # Inbox
    if include_sent:
        default_folder_ids.append(5)  # Sent Items
    for folder_id in default_folder_ids:
        try:
            folder = namespace.GetDefaultFolder(folder_id)
            potential_folders.append(folder)
        except Exception:
            continue

    # Add user requested folders
    for extra_folder in resolve_additional_folders(namespace, additional_folders):
        potential_folders.append(extra_folder)

    folders_to_scan = []
    seen_paths = set()
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

    # Sort newest first
    related_entries.sort(
        key=lambda entry: entry[0] if entry[0] else datetime.datetime.min,
        reverse=True,
    )
    return [entry[1] for entry in related_entries]

def _email_has_user_reply(
    namespace,
    email_data: Dict[str, Any],
    user_addresses: Set[str],
    conversation_limit: int,
    lookback_days: int,
) -> bool:
    """Determine whether the user has already replied within a conversation."""
    result, _, _ = _email_has_user_reply_with_context(
        namespace=namespace,
        email_data=email_data,
        user_addresses=user_addresses,
        conversation_limit=conversation_limit,
        lookback_days=lookback_days,
        collect_related=False,
    )
    return result

def _email_has_user_reply_with_context(
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
            for addr in (_normalize_email_address(addr) for addr in user_addresses)
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
            sender_email = _normalize_email_address(related.get("sender_email")) or _normalize_email_address(
                related.get("sender")
            )
            if not sender_email or sender_email not in normalized_user_addresses:
                continue
            related_dt = _extract_best_timestamp(related)
            if not baseline_dt or not related_dt or related_dt >= baseline_dt:
                return True, None, None

    if mail_item and _mail_item_marked_replied(mail_item, baseline_dt):
        return True, None, None

    return False, captured_related, mail_item

def _build_conversation_outline(
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

    timeline: List[tuple[Optional[datetime.datetime], Dict[str, Any], bool]] = []
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

        lines.append(f"{prefix} {timestamp} · {sender}: {subject}{preview_line}")

    return "\n".join(lines)

@mcp.tool()
def params(
    protocolVersion: Optional[str] = None,  # type: ignore[non-literal-used]
    capabilities: Optional[Dict[str, Any]] = None,
    clientInfo: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Provide handshake metadata compatible with MCP-aware HTTP clients (e.g. n8n).
    """
    requested_version = protocolVersion or "2025-03-26"
    logger.info(
        "params tool invocato (protocolVersion=%s, clientInfo=%s)",
        requested_version,
        clientInfo,
    )

    tool_summaries: Dict[str, Dict[str, Any]] = {}
    for tool in mcp._tool_manager.list_tools():  # type: ignore[attr-defined]
        tool_summaries[tool.name] = {
            "description": getattr(tool, "description", None),
            "inputSchema": getattr(tool, "input_schema", None),
            "outputSchema": getattr(tool, "output_schema", None),
            "annotations": getattr(tool, "annotations", None),
        }

    default_capabilities = {"tools": {"list": True, "call": True}}
    response_capabilities: Dict[str, Any] = default_capabilities
    if capabilities:
        # Merge nested dicts without mutating caller-provided payload
        response_capabilities = {**capabilities}
        tools_caps = dict(default_capabilities.get("tools", {}))
        tools_caps.update(capabilities.get("tools", {}))
        response_capabilities["tools"] = tools_caps

    return {
        "protocolVersion": requested_version,
        "serverInfo": {
            "name": "outlook-assistant",
            "version": "1.0.0-http",
            "description": (
                "Bridge MCP per Outlook. Gli strumenti abilitano ricerche email, "
                "risposte rapide e consultazione calendario tramite HTTP."
            ),
        },
        "capabilities": response_capabilities,
        "tools": tool_summaries,
        "httpBridge": {
            "health": "GET /health",
            "listTools": "GET /tools",
            "invokeTool": "POST /tools/{tool_name}",
            "invokeToolRoot": "POST /",
        },
    }

# MCP Tools
@mcp.tool()
def get_current_datetime(include_utc: bool = True) -> str:
    """
    Restituisce la data e ora correnti formattate.

    Args:
        include_utc: Includere o meno il riferimento UTC nella risposta
    """
    include_utc_bool = coerce_bool(include_utc)
    logger.info("get_current_datetime chiamato con include_utc=%s", include_utc_bool)
    try:
        local_dt = datetime.datetime.now()
        lines = [
            "Data e ora correnti:",
            f"- Locale: {local_dt.strftime('%Y-%m-%d %H:%M:%S')}",
            f"- Locale ISO: {local_dt.isoformat()}",
        ]
        if include_utc_bool:
            utc_dt = datetime.datetime.now(datetime.UTC).replace(tzinfo=datetime.timezone.utc)
            lines.append(f"- UTC: {utc_dt.strftime('%Y-%m-%d %H:%M:%S')}")
            lines.append(f"- UTC ISO: {utc_dt.isoformat()}")
        return "\n".join(lines)
    except Exception as exc:
        logger.exception("Errore durante get_current_datetime.")
        return f"Errore durante il calcolo della data/ora corrente: {exc}"


@mcp.tool()
def list_folders(
    root_folder_id: Optional[str] = None,
    root_folder_path: Optional[str] = None,
    root_folder_name: Optional[str] = None,
    max_depth: int = 2,
    include_counts: bool = True,
    include_ids: bool = False,
    include_store: bool = False,
    include_paths: bool = True,
) -> str:
    """
    Enumerate Outlook folders starting from the mailbox root (or a custom root).
    """
    if not isinstance(max_depth, int) or max_depth < 0 or max_depth > 10:
        logger.warning("Valore 'max_depth' non valido passato a list_folders: %s", max_depth)
        return "Errore: 'max_depth' deve essere un intero compreso tra 0 e 10."

    include_counts_flag = coerce_bool(include_counts)
    include_ids_flag = coerce_bool(include_ids)
    include_store_flag = coerce_bool(include_store)
    include_paths_flag = coerce_bool(include_paths)

    logger.info(
        (
            "list_folders chiamato (root_id=%s root_path=%s root_name=%s profondita=%s "
            "contatori=%s ids=%s store=%s paths=%s)."
        ),
        root_folder_id,
        root_folder_path,
        root_folder_name,
        max_depth,
        include_counts_flag,
        include_ids_flag,
        include_store_flag,
        include_paths_flag,
    )

    try:
        _, namespace = connect_to_outlook()
        return folder_service.list_folders(
            namespace,
            root_folder_id=root_folder_id,
            root_folder_path=root_folder_path,
            root_folder_name=root_folder_name,
            max_depth=max_depth,
            include_counts=include_counts,
            include_ids=include_ids,
            include_store=include_store,
            include_paths=include_paths,
        )
    except Exception as exc:
        logger.exception("Errore durante l'elenco delle cartelle di Outlook.")
        return f"Errore durante l'elenco delle cartelle: {exc}"

@mcp.tool()
def get_folder_metadata(
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
    include_children: bool = False,
    max_children: int = 20,
    include_counts: bool = True,
) -> str:
    """
    Retrieve detailed metadata for a specific Outlook folder.
    """
    if not isinstance(max_children, int) or max_children < 0:
        return "Errore: 'max_children' deve essere un intero non negativo."

    include_children_flag = coerce_bool(include_children)
    include_counts_flag = coerce_bool(include_counts)
    logger.info(
        "get_folder_metadata chiamato (id=%s path=%s nome=%s figli=%s max=%s contatori=%s).",
        folder_id,
        folder_path,
        folder_name,
        include_children_flag,
        max_children,
        include_counts_flag,
    )

    try:
        _, namespace = connect_to_outlook()
        return folder_service.folder_metadata(
            namespace,
            folder_id=folder_id,
            folder_path=folder_path,
            folder_name=folder_name,
            include_children=include_children,
            max_children=max_children,
            include_counts=include_counts,
        )
    except Exception as exc:
        logger.exception("Errore durante get_folder_metadata.")
        return f"Errore durante il recupero dei metadati della cartella: {exc}"

@mcp.tool()
def create_folder(
    new_folder_name: str,
    parent_folder_id: Optional[str] = None,
    parent_folder_path: Optional[str] = None,
    parent_folder_name: Optional[str] = None,
    item_type: Optional[Any] = None,
    allow_existing: bool = False,
) -> str:
    """
    Create a new Outlook subfolder under the specified parent.
    """
    if not new_folder_name or not new_folder_name.strip():
        return "Errore: specifica un nome valido per la nuova cartella."

    allow_existing_bool = coerce_bool(allow_existing)
    logger.info(
        "create_folder chiamato (nome=%s parent_id=%s parent_path=%s parent_name=%s tipo=%s allow_existing=%s).",
        new_folder_name,
        parent_folder_id,
        parent_folder_path,
        parent_folder_name,
        item_type,
        allow_existing_bool,
    )

    try:
        _, namespace = connect_to_outlook()
        parent, attempts = folder_service.resolve_folder(
            namespace,
            folder_id=parent_folder_id,
            folder_path=parent_folder_path,
            folder_name=parent_folder_name,
        )
        if not parent:
            detail = "; ".join(attempts) if attempts else "cartella padre non trovata."
            return f"Errore: impossibile individuare la cartella padre ({detail})."

        try:
            _, message = folder_service.create_folder(
                parent,
                new_folder_name=new_folder_name,
                item_type=item_type,
                allow_existing=allow_existing_bool,
            )
            return message
        except ValueError as exc:
            return f"Errore: {exc}"
        except RuntimeError as exc:
            return f"Errore: {exc}"
    except Exception as exc:
        logger.exception("Errore durante create_folder.")
        return f"Errore durante la creazione della cartella: {exc}"

@mcp.tool()
def rename_folder(
    folder_id: Optional[str] = None,
    new_name: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
) -> str:
    """
    Rename an existing Outlook folder.
    """
    if not new_name or not new_name.strip():
        return "Errore: specifica un nuovo nome valido per la cartella."

    logger.info(
        "rename_folder chiamato (id=%s path=%s nome=%s nuovo_nome=%s).",
        folder_id,
        folder_path,
        folder_name,
        new_name,
    )

    try:
        _, namespace = connect_to_outlook()
        target, attempts = folder_service.resolve_folder(
            namespace,
            folder_id=folder_id,
            folder_path=folder_path,
            folder_name=folder_name,
        )
        if not target:
            detail = "; ".join(attempts) if attempts else "cartella non trovata."
            return f"Errore: impossibile individuare la cartella da rinominare ({detail})."

        try:
            folder_service.rename_folder(target, new_name)
        except ValueError as exc:
            return f"Errore: {exc}"
        except RuntimeError as exc:
            return f"Errore: {exc}"

        path_display = safe_folder_path(target) or new_name.strip()
        return f"Cartella rinominata in '{new_name.strip()}' (percorso attuale: {path_display})."
    except Exception as exc:
        logger.exception("Errore durante rename_folder.")
        return f"Errore durante la rinomina della cartella: {exc}"

@mcp.tool()
def delete_folder(
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
    confirm: bool = False,
) -> str:
    """
    Delete an Outlook folder (moves it to Deleted Items unless prevented by Outlook).
    """
    if not coerce_bool(confirm):
        return "Conferma mancante: imposta confirm=True per procedere con l'eliminazione della cartella."

    logger.info(
        "delete_folder chiamato (id=%s path=%s nome=%s).",
        folder_id,
        folder_path,
        folder_name,
    )

    try:
        _, namespace = connect_to_outlook()
        target, attempts = folder_service.resolve_folder(
            namespace,
            folder_id=folder_id,
            folder_path=folder_path,
            folder_name=folder_name,
        )
        if not target:
            detail = "; ".join(attempts) if attempts else "cartella non trovata."
            return f"Errore: impossibile individuare la cartella da eliminare ({detail})."

        path_display = safe_folder_path(target) or getattr(target, "Name", "(sconosciuta)")
        try:
            folder_service.delete_folder(target)
        except RuntimeError as exc:
            return f"Errore: {exc}"

        return (
            f"Cartella eliminata: {path_display}. (Se previsto, Outlook l'ha spostata in Posta eliminata.)"
        )
    except Exception as exc:
        logger.exception("Errore durante delete_folder.")
        return f"Errore durante l'eliminazione della cartella: {exc}"


def _present_email_listing(
    emails: List[Dict[str, Any]],
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

    visible_emails = emails[start_index:start_index + max_results]
    visible_count = len(visible_emails)
    total_count = len(emails)
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
        importance_label = email.get("importance_label") or _describe_importance(email.get("importance"))
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
        result += f"Stato lettura: {_read_status(email.get('unread'))}\n"
        result += f"Allegati: {_yes_no(email.get('has_attachments'))}\n"
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

def _present_event_listing(
    events: List[Dict[str, Any]],
    calendar_display: str,
    days: int,
    max_results: int,
    include_description: bool,
    log_context: str,
) -> str:
    """Common presenter for calendar events."""
    clear_calendar_cache()

    if not events:
        logger.info(
            "%s: nessun evento trovato (calendario=%s giorni=%s).",
            log_context,
            calendar_display,
            days,
        )
        return f"Nessun evento in {calendar_display} nei prossimi {days} giorni."

    visible_events = events[:max_results]
    visible_count = len(visible_events)
    total_count = len(events)

    if total_count > visible_count:
        header = (
            f"Trovati {total_count} eventi in {calendar_display} nei prossimi {days} giorni. "
            f"Mostro i primi {visible_count} risultati."
        )
    else:
        header = f"Trovati {visible_count} eventi in {calendar_display} nei prossimi {days} giorni."

    logger.info(
        "%s: restituiti %s eventi su %s (calendario=%s).",
        log_context,
        visible_count,
        total_count,
        calendar_display,
    )

    result = header + "\n\n"

    for idx, event in enumerate(visible_events, 1):
        calendar_cache[idx] = event

        result += f"Evento #{idx}\n"
        result += f"Oggetto: {event.get('subject', '(Senza oggetto)')}\n"
        result += f"Inizio: {event.get('start_time', 'Sconosciuto')}\n"
        result += f"Fine: {event.get('end_time', 'Sconosciuto')}\n"
        result += f"Calendario: {event.get('folder_path') or calendar_display}\n"
        result += f"Luogo: {event.get('location', '') or 'Non specificato'}\n"
        result += f"Organizzatore: {event.get('organizer', 'Non disponibile')}\n"
        result += f"Giornata intera: {_yes_no(event.get('all_day'))}\n"
        if event.get("required_attendees"):
            result += f"Partecipanti obbligatori: {event['required_attendees']}\n"
        if event.get("optional_attendees"):
            result += f"Partecipanti facoltativi: {event['optional_attendees']}\n"
        if event.get("categories"):
            result += f"Categorie: {event['categories']}\n"
        if include_description and event.get("preview"):
            result += f"Anteprima: {event['preview']}\n"
        result += "\n"

    return result.rstrip()


@mcp.tool()
def list_recent_emails(
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
    folder_ids: Optional[Any] = None,
    folder_paths: Optional[Any] = None,
    offset: int = 0,
    unread_only: bool = False,
) -> str:
    """List email titles from the specified number of days with pagination support."""
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        logger.warning("Valore 'days' non valido passato a list_recent_emails: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a list_recent_emails: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"
    try:
        offset_value = int(offset)
        if offset_value < 0:
            raise ValueError
    except (TypeError, ValueError):
        logger.warning("Valore 'offset' non valido passato a list_recent_emails: %s", offset)
        return "Errore: 'offset' deve essere un intero maggiore o uguale a zero."

    include_preview_bool = coerce_bool(include_preview)
    include_all_bool = coerce_bool(include_all_folders)
    unread_only_bool = coerce_bool(unread_only)
    folder_id_list = ensure_string_list(folder_ids)
    folder_path_list = ensure_string_list(folder_paths)
    logger.info(
        (
            "list_recent_emails chiamato con giorni=%s cartella=%s max_risultati=%s "
            "anteprima=%s tutte_le_cartelle=%s offset=%s unread_only=%s folder_ids=%s folder_paths=%s"
        ),
        days,
        folder_name,
        max_results,
        include_preview_bool,
        include_all_bool,
        offset_value,
        unread_only_bool,
        folder_id_list,
        folder_path_list,
    )

    try:
        _, namespace = connect_to_outlook()

        emails: List[Dict[str, Any]]
        folder_display: str
        if folder_id_list or folder_path_list:
            selected_folders: List = []
            failures: List[str] = []
            for entry_id in folder_id_list:
                folder, attempts = folder_service.resolve_folder(namespace, folder_id=entry_id)
                if folder:
                    selected_folders.append(folder)
                else:
                    detail = ", ".join(attempts) if attempts else "non trovato"
                    failures.append(f"ID {entry_id}: {detail}")
            for folder_path in folder_path_list:
                folder, attempts = folder_service.resolve_folder(namespace, folder_path=folder_path)
                if folder:
                    selected_folders.append(folder)
                else:
                    detail = ", ".join(attempts) if attempts else "non trovato"
                    failures.append(f"Percorso {folder_path}: {detail}")
            if not selected_folders:
                detail = "; ".join(failures) if failures else "cartelle non trovate."
                return f"Errore: impossibile individuare le cartelle richieste ({detail})."
            emails = collect_emails_across_folders(selected_folders, days)
            folder_display = "Cartelle selezionate"
        elif include_all_bool:
            if folder_name:
                logger.info("Parametro folder_name ignorato perche include_all_folders=True.")
            folders = get_all_mail_folders(namespace)
            emails = collect_emails_across_folders(folders, days)
            folder_display = "Tutte le cartelle"
        else:
            if folder_name:
                folder = folder_service.get_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella '{folder_name}' non trovata"
            else:
                folder = namespace.GetDefaultFolder(6)
            folder_display = f"'{folder_name}'" if folder_name else "Posta in arrivo"
            emails = get_emails_from_folder(folder, days)

        if unread_only_bool:
            emails = [email for email in emails if email.get("unread")]

        return _present_email_listing(
            emails=emails,
            folder_display=folder_display,
            days=days,
            max_results=max_results,
            include_preview=include_preview_bool,
            log_context="list_recent_emails",
            offset=offset_value,
        )
    except Exception as exc:
        logger.exception("Errore nel recupero dei messaggi per la cartella '%s'.", folder_name or "Posta in arrivo")
        return f"Errore durante il recupero dei messaggi: {exc}"

@mcp.tool()
def list_sent_emails(
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    offset: int = 0,
) -> str:
    """
    Elenca i messaggi inviati di recente per supportare il recupero del contesto.
    
    Args:
        days: Giorni da considerare a ritroso (max 30)
        folder_name: Cartella specifica da controllare (default: Posta inviata)
        max_results: Numero massimo di messaggi da mostrare (1-200)
        include_preview: Includere una breve anteprima del corpo
    
    Returns:
        Elenco numerato dei messaggi inviati
    """
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        logger.warning("Valore 'days' non valido passato a list_sent_emails: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a list_sent_emails: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"
    try:
        offset_value = int(offset)
        if offset_value < 0:
            raise ValueError
    except (TypeError, ValueError):
        logger.warning("Valore 'offset' non valido passato a list_sent_emails: %s", offset)
        return "Errore: 'offset' deve essere un intero maggiore o uguale a zero."

    include_preview = coerce_bool(include_preview)
    logger.info(
        "list_sent_emails chiamato con giorni=%s cartella=%s max_risultati=%s anteprima=%s offset=%s",
        days,
        folder_name,
        max_results,
        include_preview,
        offset_value,
    )

    try:
        _, namespace = connect_to_outlook()

        if folder_name:
            folder = folder_service.get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Errore: cartella '{folder_name}' non trovata"
            folder_display = f"'{folder_name}'"
        else:
            folder = namespace.GetDefaultFolder(5)  # Sent Items
            folder_display = "Posta inviata"

        emails = get_emails_from_folder(folder, days)
        return _present_email_listing(
            emails=emails,
            folder_display=folder_display,
            days=days,
            max_results=max_results,
            include_preview=include_preview,
            log_context="list_sent_emails",
            search_term=None,
            focus_on_recipients=True,
            offset=offset_value,
        )
    except Exception as e:
        logger.exception("Errore nel recupero dei messaggi inviati per la cartella '%s'.", folder_name or "Posta inviata")
        return f"Errore durante il recupero dei messaggi inviati: {str(e)}"


@mcp.tool()
def search_emails(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
    folder_ids: Optional[Any] = None,
    folder_paths: Optional[Any] = None,
    offset: int = 0,
    unread_only: bool = False,
) -> str:
    """Search emails by contact name or keyword within a time period."""
    if not search_term:
        logger.warning("search_emails chiamato senza termine di ricerca.")
        return "Errore: inserisci un termine di ricerca"

    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        logger.warning("Valore 'days' non valido passato a search_emails: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a search_emails: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"
    try:
        offset_value = int(offset)
        if offset_value < 0:
            raise ValueError
    except (TypeError, ValueError):
        logger.warning("Valore 'offset' non valido passato a search_emails: %s", offset)
        return "Errore: 'offset' deve essere un intero maggiore o uguale a zero."

    include_preview_bool = coerce_bool(include_preview)
    include_all_bool = coerce_bool(include_all_folders)
    unread_only_bool = coerce_bool(unread_only)
    folder_id_list = ensure_string_list(folder_ids)
    folder_path_list = ensure_string_list(folder_paths)
    logger.info(
        (
            "search_emails chiamato con termine='%s' giorni=%s cartella=%s max_risultati=%s "
            "anteprima=%s tutte_le_cartelle=%s offset=%s unread_only=%s folder_ids=%s folder_paths=%s"
        ),
        search_term,
        days,
        folder_name,
        max_results,
        include_preview_bool,
        include_all_bool,
        offset_value,
        unread_only_bool,
        folder_id_list,
        folder_path_list,
    )

    try:
        _, namespace = connect_to_outlook()

        emails: List[Dict[str, Any]]
        folder_display: str
        if folder_id_list or folder_path_list:
            selected_folders: List = []
            failures: List[str] = []
            for entry_id in folder_id_list:
                folder, attempts = folder_service.resolve_folder(namespace, folder_id=entry_id)
                if folder:
                    selected_folders.append(folder)
                else:
                    detail = ", ".join(attempts) if attempts else "non trovato"
                    failures.append(f"ID {entry_id}: {detail}")
            for folder_path in folder_path_list:
                folder, attempts = folder_service.resolve_folder(namespace, folder_path=folder_path)
                if folder:
                    selected_folders.append(folder)
                else:
                    detail = ", ".join(attempts) if attempts else "non trovato"
                    failures.append(f"Percorso {folder_path}: {detail}")
            if not selected_folders:
                detail = "; ".join(failures) if failures else "cartelle non trovate."
                return f"Errore: impossibile individuare le cartelle richieste ({detail})."
            emails = collect_emails_across_folders(selected_folders, days, search_term)
            folder_display = "Cartelle selezionate"
        elif include_all_bool:
            if folder_name:
                logger.info("Parametro folder_name ignorato perche include_all_folders=True.")
            folders = get_all_mail_folders(namespace)
            emails = collect_emails_across_folders(folders, days, search_term)
            folder_display = "Tutte le cartelle"
        else:
            if folder_name:
                folder = folder_service.get_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella '{folder_name}' non trovata"
            else:
                folder = namespace.GetDefaultFolder(6)
            folder_display = f"'{folder_name}'" if folder_name else "Posta in arrivo"
            emails = get_emails_from_folder(folder, days, search_term)

        if unread_only_bool:
            emails = [email for email in emails if email.get("unread")]

        return _present_email_listing(
            emails=emails,
            folder_display=folder_display,
            days=days,
            max_results=max_results,
            include_preview=include_preview_bool,
            log_context="search_emails",
            search_term=search_term,
            offset=offset_value,
        )
    except Exception as exc:
        logger.exception(
            "Errore durante la ricerca dei messaggi con termine '%s' nella cartella '%s'.",
            search_term,
            folder_name or "Posta in arrivo",
        )
        return f"Errore nella ricerca dei messaggi: {exc}"

@mcp.tool()
def list_pending_replies(
    days: int = 14,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
    include_unread_only: bool = False,
    conversation_lookback_days: Optional[int] = None,
) -> str:
    """
    Elenca i messaggi piu recenti che non risultano ancora gestiti con una risposta inviata.
    """
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        logger.warning("Valore 'days' non valido passato a list_pending_replies: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a list_pending_replies: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"

    include_preview_bool = coerce_bool(include_preview)
    include_all_bool = coerce_bool(include_all_folders)
    unread_only_bool = coerce_bool(include_unread_only)

    if conversation_lookback_days is None:
        lookback_days = max(days * 2, 14)
    elif isinstance(conversation_lookback_days, int) and 1 <= conversation_lookback_days <= MAX_CONVERSATION_LOOKBACK_DAYS:
        lookback_days = conversation_lookback_days
    else:
        logger.warning(
            "Valore 'conversation_lookback_days' non valido passato a list_pending_replies: %s",
            conversation_lookback_days,
        )
        return f"Errore: 'conversation_lookback_days' deve essere un intero tra 1 e {MAX_CONVERSATION_LOOKBACK_DAYS}"

    lookback_days = min(lookback_days, MAX_CONVERSATION_LOOKBACK_DAYS)
    max_processed_before_break = max(max_results * PENDING_SCAN_MULTIPLIER, max_results + 25)
    logger.info(
        (
            "list_pending_replies chiamato con giorni=%s cartella=%s max_risultati=%s "
            "anteprima=%s tutte_le_cartelle=%s solo_non_letti=%s lookback_conv=%s"
        ),
        days,
        folder_name,
        max_results,
        include_preview_bool,
        include_all_bool,
        unread_only_bool,
        lookback_days,
    )

    try:
        _, namespace = connect_to_outlook()
        user_addresses = _collect_user_addresses(namespace)
        normalized_user_addresses = {
            addr for addr in (_normalize_email_address(addr) for addr in user_addresses) if addr
        }

        if include_all_bool:
            if folder_name:
                logger.info("Parametro folder_name ignorato poiche include_all_folders=True.")
            folders = get_all_mail_folders(namespace)
            candidate_emails = collect_emails_across_folders(
                folders,
                days,
                target_total=max_processed_before_break,
            )
            folder_display = "Tutte le cartelle (senza risposta)"
        else:
            if folder_name:
                folder = folder_service.get_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella '{folder_name}' non trovata"
                candidate_emails = get_emails_from_folder(folder, days)
                folder_display = f"'{folder_name}' (senza risposta)"
            else:
                folder = namespace.GetDefaultFolder(6)
                candidate_emails = get_emails_from_folder(folder, days)
                folder_display = "Posta in arrivo (senza risposta)"
            if len(candidate_emails) > max_processed_before_break:
                candidate_emails = candidate_emails[:max_processed_before_break]

        pending_emails: List[Dict[str, Any]] = []
        processed = 0
        truncated_scan = False

        for email in candidate_emails:
            processed += 1

            if unread_only_bool and not email.get("unread"):
                if processed >= max_processed_before_break and len(pending_emails) >= max_results:
                    truncated_scan = processed < len(candidate_emails)
                    break
                continue

            sender_email = _normalize_email_address(email.get("sender_email")) or _normalize_email_address(
                email.get("sender")
            )
            if sender_email and sender_email in normalized_user_addresses:
                if processed >= max_processed_before_break and len(pending_emails) >= max_results:
                    truncated_scan = processed < len(candidate_emails)
                    break
                continue

            already_replied = False
            related_entries: Optional[List[Dict[str, Any]]] = None
            mail_item_ref = None
            try:
                already_replied, related_entries, mail_item_ref = _email_has_user_reply_with_context(
                    namespace=namespace,
                    email_data=email,
                    user_addresses=user_addresses,
                    conversation_limit=DEFAULT_CONVERSATION_SAMPLE_LIMIT,
                    lookback_days=lookback_days,
                    collect_related=True,
                )
            except Exception:
                logger.debug(
                    "Controllo risposta fallito per il messaggio %s.",
                    email.get("id"),
                    exc_info=True,
                )
                already_replied = False
                related_entries = None
                mail_item_ref = None

            if not already_replied:
                outline = _build_conversation_outline(
                    namespace=namespace,
                    email_data=email,
                    lookback_days=lookback_days,
                    max_items=4,
                    preloaded_entries=related_entries,
                    mail_item=mail_item_ref,
                )
                if outline:
                    email["_conversation_outline"] = outline
                pending_emails.append(email)

            if len(pending_emails) >= max_results:
                truncated_scan = processed < len(candidate_emails)
                break

        logger.info(
            "list_pending_replies ha trovato %s messaggi da gestire su %s analizzati.",
            len(pending_emails),
            processed,
        )

        presentation = _present_email_listing(
            emails=pending_emails,
            folder_display=folder_display,
            days=days,
            max_results=max_results,
            include_preview=include_preview_bool,
            log_context="list_pending_replies",
        )

        if truncated_scan:
            presentation += (
                "\nNota: per limitare i tempi di ricerca sono stati esaminati i messaggi piu recenti. "
                "Aumenta 'days' o 'max_results' per ampliare il controllo."
            )
        elif not user_addresses:
            presentation += (
                "\nAvvertenza: impossibile determinare l'indirizzo dell'account Outlook, i risultati potrebbero essere incompleti."
            )

        return presentation

    except Exception as exc:
        logger.exception("Errore durante list_pending_replies per la cartella '%s'.", folder_name or "Posta in arrivo")
        return f"Errore durante il calcolo delle risposte mancanti: {str(exc)}"

@mcp.tool()
def list_upcoming_events(
    days: int = 7,
    calendar_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_description: bool = False,
    include_all_calendars: bool = False,
) -> str:
    """
    Elenca gli eventi in arrivo dal calendario selezionato.
    
    Args:
        days: Orizzonte temporale in giorni (max 90)
        calendar_name: Nome del calendario da esaminare (facoltativo)
        max_results: Numero massimo di eventi da mostrare (1-200)
        include_description: Includere un'anteprima della descrizione dell'evento
        include_all_calendars: Ignora calendar_name e scansiona tutti i calendari disponibili
    """
    if not isinstance(days, int) or days < 1 or days > MAX_EVENT_LOOKAHEAD_DAYS:
        logger.warning("Valore 'days' non valido passato a list_upcoming_events: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_EVENT_LOOKAHEAD_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a list_upcoming_events: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"

    include_desc = coerce_bool(include_description)
    include_all = coerce_bool(include_all_calendars)
    logger.info(
        "list_upcoming_events chiamato con giorni=%s calendario=%s max_risultati=%s descrizione=%s tutti_i_calendari=%s",
        days,
        calendar_name,
        max_results,
        include_desc,
        include_all,
    )

    try:
        _, namespace = connect_to_outlook()

        if include_all:
            calendars = get_all_calendar_folders(namespace)
            events = collect_events_across_calendars(calendars, days)
            calendar_display = "Tutti i calendari"
        else:
            calendar_folder = get_calendar_folder_by_name(namespace, calendar_name) if calendar_name else namespace.GetDefaultFolder(9)
            if not calendar_folder:
                return f"Errore: calendario '{calendar_name}' non trovato"
            calendar_display = calendar_folder.Name if calendar_name else "Calendario"
            events = get_events_from_folder(calendar_folder, days)

        return _present_event_listing(
            events=events,
            calendar_display=calendar_display,
            days=days,
            max_results=max_results,
            include_description=include_desc,
            log_context="list_upcoming_events",
        )
    except Exception as e:
        logger.exception("Errore durante il recupero degli eventi per il calendario '%s'.", calendar_name or "Calendario")
        return f"Errore durante il recupero degli eventi: {str(e)}"

@mcp.tool()
def search_calendar_events(
    search_term: str,
    days: int = 30,
    calendar_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_description: bool = False,
    include_all_calendars: bool = False,
) -> str:
    """
    Cerca eventi nel calendario in base a una parola chiave.
    """
    if not search_term:
        logger.warning("search_calendar_events chiamato senza termine di ricerca.")
        return "Errore: inserisci un termine di ricerca per il calendario"

    if not isinstance(days, int) or days < 1 or days > MAX_EVENT_LOOKAHEAD_DAYS:
        logger.warning("Valore 'days' non valido passato a search_calendar_events: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_EVENT_LOOKAHEAD_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a search_calendar_events: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"

    include_desc = coerce_bool(include_description)
    include_all = coerce_bool(include_all_calendars)
    logger.info(
        "search_calendar_events chiamato con termine='%s' giorni=%s calendario=%s max_risultati=%s descrizione=%s tutti_i_calendari=%s",
        search_term,
        days,
        calendar_name,
        max_results,
        include_desc,
        include_all,
    )

    try:
        _, namespace = connect_to_outlook()

        if include_all:
            calendars = get_all_calendar_folders(namespace)
            events = collect_events_across_calendars(calendars, days, search_term)
            calendar_display = "Tutti i calendari"
        else:
            calendar_folder = get_calendar_folder_by_name(namespace, calendar_name) if calendar_name else namespace.GetDefaultFolder(9)
            if not calendar_folder:
                return f"Errore: calendario '{calendar_name}' non trovato"
            calendar_display = calendar_folder.Name if calendar_name else "Calendario"
            events = get_events_from_folder(calendar_folder, days, search_term)

        return _present_event_listing(
            events=events,
            calendar_display=calendar_display,
            days=days,
            max_results=max_results,
            include_description=include_desc,
            log_context="search_calendar_events",
        )
    except Exception as e:
        logger.exception(
            "Errore durante la ricerca di eventi con termine '%s' nel calendario '%s'.",
            search_term,
            calendar_name or "Calendario",
        )
        return f"Errore durante la ricerca degli eventi: {str(e)}"


@mcp.tool()
def create_calendar_event(
    subject: str,
    start_time: str,
    duration_minutes: Optional[int] = 60,
    location: Optional[str] = None,
    body: Optional[str] = None,
    attendees: Optional[Any] = None,
    reminder_minutes: Optional[int] = 15,
    calendar_name: Optional[str] = None,
    all_day: bool = False,
    send_invitations: bool = True,
) -> str:
    """Create a new Outlook calendar event (meeting or appointment)."""
    if not subject or not subject.strip():
        return "Errore: specifica un oggetto ('subject') per l'evento."

    start_dt = _parse_datetime_string(start_time)
    if not start_dt:
        return "Errore: 'start_time' deve essere una data valida (es. '2025-10-20 10:30')."

    all_day_bool = coerce_bool(all_day)
    send_bool = coerce_bool(send_invitations)

    if not all_day_bool:
        if duration_minutes is None:
            duration_value = 60
        else:
            try:
                duration_value = int(duration_minutes)
            except (TypeError, ValueError):
                return "Errore: 'duration_minutes' deve essere un intero positivo."
            if duration_value <= 0:
                return "Errore: 'duration_minutes' deve essere un intero positivo."
    else:
        duration_value = None

    attendee_list = ensure_string_list(attendees)

    reminder_set = False
    reminder_value: Optional[int] = None
    if reminder_minutes is not None:
        try:
            reminder_value = int(reminder_minutes)
        except (TypeError, ValueError):
            return "Errore: 'reminder_minutes' deve essere un intero (minuti)."
        if reminder_value >= 0:
            reminder_set = True
        else:
            reminder_value = None

    logger.info(
        "create_calendar_event chiamato (subject=%s start=%s durata=%s luogo=%s partecipanti=%s all_day=%s invia=%s calendario=%s).",
        subject,
        start_time,
        duration_value,
        location,
        attendee_list,
        all_day_bool,
        send_bool,
        calendar_name,
    )

    try:
        outlook, namespace = connect_to_outlook()
        if calendar_name:
            target_calendar = get_calendar_folder_by_name(namespace, calendar_name)
            if not target_calendar:
                return f"Errore: calendario '{calendar_name}' non trovato."
        else:
            target_calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar

        calendar_display = getattr(target_calendar, "Name", "Calendario")
        appointment = outlook.CreateItem(1)  # olAppointmentItem

        subject_clean = subject.strip()
        appointment.Subject = subject_clean

        if all_day_bool:
            normalized_start = start_dt.replace(hour=0, minute=0, second=0, microsecond=0)
            appointment.Start = normalized_start
            appointment.End = normalized_start + datetime.timedelta(days=1)
            appointment.AllDayEvent = True
        else:
            appointment.Start = start_dt
            appointment.AllDayEvent = False
            appointment.Duration = duration_value or 60
            appointment.End = start_dt + datetime.timedelta(minutes=appointment.Duration)

        if location:
            appointment.Location = location
        if body:
            appointment.Body = body

        if attendee_list:
            appointment.MeetingStatus = 1  # olMeeting
            for email in attendee_list:
                if not email:
                    continue
                try:
                    recipient = appointment.Recipients.Add(email)
                    if hasattr(recipient, "Type"):
                        recipient.Type = 1  # Required attendee
                except Exception as exc:
                    logger.warning("Impossibile aggiungere il destinatario '%s': %s", email, exc)

        appointment.ReminderSet = reminder_set
        if reminder_set and reminder_value is not None:
            appointment.ReminderMinutesBeforeStart = reminder_value

        try:
            appointment.Save()
        except Exception as exc:
            logger.exception("Salvataggio dell'appuntamento fallito.")
            return f"Errore: impossibile salvare l'evento ({exc})."

        if calendar_name:
            try:
                moved = appointment.Move(target_calendar)
                if moved:
                    appointment = moved
            except Exception as exc:
                logger.exception("Impossibile spostare l'evento nel calendario '%s'.", calendar_display)
                return (
                    "Errore: salvataggio completato ma non è stato possibile spostare l'evento nel "
                    f"calendario '{calendar_display}' ({exc})."
                )

        if send_bool and attendee_list:
            try:
                appointment.Send()
            except Exception as exc:
                logger.exception("Invio degli inviti fallito.")
                return f"Errore: evento creato ma invio degli inviti fallito ({exc})."

        clear_calendar_cache()

        entry_id = safe_entry_id(appointment) or "N/D"
        start_display = start_dt.strftime("%Y-%m-%d") if all_day_bool else start_dt.strftime("%Y-%m-%d %H:%M")

        lines = [
            f"Evento '{subject_clean}' creato per {start_display} nel calendario '{calendar_display}'.",
            f"Message ID: {entry_id}",
        ]
        if attendee_list:
            lines.append(f"Partecipanti: {', '.join(attendee_list)}")
        if send_bool and attendee_list:
            lines.append("Inviti inviati ai partecipanti.")
        elif attendee_list:
            lines.append("Inviti non inviati (send_invitations impostato a False).")

        return "\n".join(lines)
    except Exception as exc:
        logger.exception("Errore durante create_calendar_event.")
        return f"Errore durante la creazione dell'evento: {exc}"

@mcp.tool()
def get_email_by_number(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    index: Optional[int] = None,
    include_body: bool = True,
) -> str:
    """
    Retrieve the detailed content of a specific email.
    
    Args:
        email_number: Number previously assigned during an email listing.
        message_id: Outlook EntryID of the message (alternative to email_number).
        folder_id: EntryID of the folder that contains the message (used with index).
        folder_path: FolderPath string used to locate the folder (used with index).
        index: 1-based position inside the folder (sorted by ReceivedTime desc).
        include_body: Include the full message body in the textual response.
    """
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
        importance_label = email_data.get("importance_label") or _describe_importance(email_data.get("importance"))
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
                f"Stato lettura: {_read_status(email_data.get('unread'))}",
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

        result_lines.append(f"Allegati: {_yes_no(email_data.get('has_attachments'))}")
        if attachment_names_preview:
            result_lines.append(f"Nomi allegati: {', '.join(attachment_names_preview)}")

        attachment_lines = []
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
def get_event_by_number(event_number: int) -> str:
    """
    Recupera i dettagli completi di un evento dall'ultimo elenco.
    """
    try:
        if not calendar_cache:
            logger.warning("get_event_by_number chiamato ma la cache eventi e vuota.")
            return "Errore: nessun elenco eventi attivo. Chiedimi di aggiornare gli eventi e poi ripeti la richiesta."

        if event_number not in calendar_cache:
            logger.warning("Evento numero %s non presente in cache per get_event_by_number.", event_number)
            return f"Errore: l'evento #{event_number} non e presente nell'elenco corrente."

        event = calendar_cache[event_number]
        logger.info("Recupero dettagli completi per l'evento #%s.", event_number)

        lines = [
            f"Dettagli evento #{event_number}:",
            "",
            f"Oggetto: {event.get('subject', '(Senza oggetto)')}",
            f"Inizio: {event.get('start_time', 'Sconosciuto')}",
            f"Fine: {event.get('end_time', 'Sconosciuto')}",
            f"Calendario: {event.get('folder_path', 'Non disponibile')}",
            f"Luogo: {event.get('location', '') or 'Non specificato'}",
            f"Organizzatore: {event.get('organizer', 'Non disponibile')}",
            f"Giornata intera: {_yes_no(event.get('all_day'))}",
            f"Evento ricorrente: {_yes_no(event.get('is_recurring'))}",
        ]

        if event.get("duration_minutes"):
            lines.append(f"Durata (minuti): {event['duration_minutes']}")
        if event.get("required_attendees"):
            lines.append(f"Partecipanti obbligatori: {event['required_attendees']}")
        if event.get("optional_attendees"):
            lines.append(f"Partecipanti facoltativi: {event['optional_attendees']}")
        if event.get("categories"):
            lines.append(f"Categorie: {event['categories']}")
        if event.get("preview"):
            lines.append(f"Anteprima descrizione: {event['preview']}")

        body_content = event.get("body", "")
        if body_content and len(body_content) > 4000:
            body_content = body_content[:4000].rstrip() + "\n[Descrizione troncata per brevita]"

        lines.append("")
        lines.append("Descrizione completa:")
        lines.append(body_content or "(Nessuna descrizione)")

        return "\n".join(lines)
    except Exception as e:
        logger.exception("Errore nel recupero dei dettagli per l'evento #%s.", event_number)
        return f"Errore durante il recupero dei dettagli dell'evento: {str(e)}"


@mcp.tool()
def ensure_domain_folder(
    email_number: Optional[int] = None,
    sender_email: Optional[str] = None,
    root_folder_name: Optional[str] = None,
    subfolders: Optional[str] = None,
) -> str:
    """Crea (se manca) la struttura di cartelle per il dominio del mittente."""
    try:
        target_email = sender_email
        email_entry: Optional[Dict[str, Any]] = None
        if email_number:
            if not email_cache:
                return "Errore: nessun elenco messaggi attivo. Mostra prima le email e riprova."
            email_entry = email_cache.get(email_number)
            if not email_entry:
                return f"Errore: il messaggio #{email_number} non e presente nella cache corrente."
            target_email = target_email or _derive_sender_email(email_entry)
        if not target_email:
            return "Errore: specifica un mittente (sender_email) oppure il numero di un messaggio gia elencato."

        domain = _extract_email_domain(target_email)
        if not domain:
            return f"Errore: impossibile determinare il dominio dal mittente '{target_email}'."

        custom_subfolders: Optional[List[str]] = None
        if subfolders:
            custom_subfolders = [folder.strip() for folder in subfolders.split("|") if folder.strip()]

        _, namespace = connect_to_outlook()
        domain_folder, domain_created, created_subfolders = _ensure_domain_folder_structure(
            namespace=namespace,
            domain=domain,
            root_folder_name=root_folder_name or DEFAULT_DOMAIN_ROOT_NAME,
            subfolders=custom_subfolders or DEFAULT_DOMAIN_SUBFOLDERS,
        )
        folder_path = getattr(domain_folder, "FolderPath", f"{domain_folder}")
        summary_parts = [f"Cartella dominio '{domain}' pronta: {folder_path}"]
        if domain_created:
            summary_parts.append("Cartella dominio creata ex novo.")
        if created_subfolders:
            summary_parts.append(
                "Sottocartelle create: " + ", ".join(created_subfolders)
            )
        return " ".join(summary_parts)
    except Exception as exc:
        logger.exception("Errore durante ensure_domain_folder (email_number=%s).", email_number)
        return f"Errore durante la verifica/creazione della cartella dominio: {exc}"


@mcp.tool()
def move_email_to_domain_folder(
    email_number: int,
    root_folder_name: Optional[str] = None,
    create_if_missing: bool = True,
    subfolders: Optional[str] = None,
) -> str:
    """Sposta un messaggio nella cartella del dominio mittente creando la struttura se necessario."""
    try:
        if not email_cache or email_number not in email_cache:
            return "Errore: nessun elenco messaggi attivo o numero non valido."

        email_entry = email_cache[email_number]
        sender = _derive_sender_email(email_entry)
        if not sender:
            return "Errore: il messaggio non contiene un mittente valido."
        domain = _extract_email_domain(sender)
        if not domain:
            return f"Errore: impossibile determinare il dominio dal mittente '{sender}'."

        _, namespace = connect_to_outlook()

        if create_if_missing:
            custom_subfolders = [seg.strip() for seg in subfolders.split("|") if seg.strip()] if subfolders else None
            domain_folder, _, _ = _ensure_domain_folder_structure(
                namespace=namespace,
                domain=domain,
                root_folder_name=root_folder_name or DEFAULT_DOMAIN_ROOT_NAME,
                subfolders=custom_subfolders or DEFAULT_DOMAIN_SUBFOLDERS,
            )
        else:
            inbox = namespace.GetDefaultFolder(6)
            root_folder = None
            try:
                for sub in inbox.Folders:
                    if sub.Name.lower() == (root_folder_name or DEFAULT_DOMAIN_ROOT_NAME).lower():
                        root_folder = sub
                        break
            except Exception:
                root_folder = None
            domain_folder = None
            if root_folder:
                try:
                    for sub in root_folder.Folders:
                        if sub.Name.lower() == domain.lower():
                            domain_folder = sub
                            break
                except Exception:
                    domain_folder = None
            if not domain_folder:
                return "Cartella dominio non trovata e creazione disabilitata."

        mail_item = namespace.GetItemFromID(email_entry["id"])
        if not mail_item:
            return f"Errore: impossibile recuperare il messaggio #{email_number} da Outlook."

        mail_item.Move(domain_folder)
        folder_path = getattr(domain_folder, "FolderPath", f"{domain_folder}")
        return (
            f"Messaggio #{email_number} spostato nella cartella dominio '{domain}' "
            f"({folder_path})."
        )
    except Exception as exc:
        logger.exception("Errore durante move_email_to_domain_folder per messaggio #%s.", email_number)
        return f"Errore durante lo spostamento nella cartella dominio: {exc}"


@mcp.tool()
def set_email_category(
    email_number: int,
    category: str,
    overwrite: bool = False,
) -> str:
    """Compatibilita retro: delega a apply_category."""
    overwrite_bool = coerce_bool(overwrite)
    append_flag = not overwrite_bool
    return apply_category(
        categories=[category],
        email_number=email_number,
        overwrite=overwrite_bool,
        append=append_flag,
    )

@mcp.tool()
def get_email_context(
    email_number: int,
    include_thread: bool = True,
    thread_limit: int = 5,
    lookback_days: int = 30,
    include_sent: bool = True,
    additional_folders: Optional[List[str]] = None,
) -> str:
    """
    Provide conversation-aware context for a previously listed email.
    
    Args:
        email_number: The number of the email from the last list/search result
        include_thread: Whether to include other emails from the same conversation
        thread_limit: Maximum number of related conversation emails to include
        lookback_days: How far back to look for related messages
        include_sent: Include messaggi della Posta inviata nella ricerca del thread
        additional_folders: Elenco di cartelle extra da scandire (nomi Outlook)
    
    Returns:
        Detailed context summary for the specified email
    """
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
        
        extra_folders = None
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

        importance_label = email_data.get("importance_label") or _describe_importance(email_data.get("importance"))

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
                f"Stato lettura: {_read_status(email_data.get('unread'))}",
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

        # Always leave a blank line before additional sections
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

    except Exception as e:
        logger.exception("Errore nel recupero del contesto per il messaggio #%s.", email_number)
        return f"Errore durante il recupero del contesto del messaggio: {str(e)}"

@mcp.tool()
def move_email_to_folder(
    target_folder_id: Optional[str] = None,
    target_folder_path: Optional[str] = None,
    target_folder_name: Optional[str] = None,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    create_if_missing: bool = False,
) -> str:
    """Move a message to a specific Outlook folder."""
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

        _, namespace = connect_to_outlook()
        try:
            cached_entry, mail_item = _resolve_mail_item(
                namespace, email_number=email_number, message_id=message_id
            )
        except ToolError as exc:
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
            _update_cached_email(
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

@mcp.tool()
def mark_email_read_unread(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    unread: Optional[bool] = None,
    flag: Optional[str] = None,
) -> str:
    """Toggle the unread flag of a message."""
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

        _, namespace = connect_to_outlook()
        try:
            _, mail_item = _resolve_mail_item(namespace, email_number=email_number, message_id=message_id)
        except ToolError as exc:
            return f"Errore: {exc}"

        try:
            mail_item.UnRead = target_unread
            mail_item.Save()
        except Exception as exc:
            logger.exception("Outlook ha rifiutato l'aggiornamento dello stato lettura.")
            return f"Errore: impossibile aggiornare lo stato lettura ({exc})."

        _update_cached_email(email_number, unread=target_unread)
        status_label = "Non letta" if target_unread else "Letta"
        reference = f"#{email_number}" if email_number is not None else (message_id or safe_entry_id(mail_item) or "messaggio")
        return f"Messaggio {reference} contrassegnato come {status_label}."
    except Exception as exc:
        logger.exception("Errore durante mark_email_read_unread.")
        return f"Errore durante l'aggiornamento dello stato di lettura: {exc}"

def _apply_categories_to_item(mail_item, categories: List[str], overwrite: bool, append: bool) -> List[str]:
    """Apply categories to the given mail item and return the final set."""
    normalized = [cat.strip() for cat in categories if cat and cat.strip()]
    if not normalized:
        raise ValueError("Nessuna categoria valida fornita.")

    existing_raw = getattr(mail_item, "Categories", "") or ""
    existing = [segment.strip() for segment in existing_raw.split(";") if segment.strip()]
    existing_set = {cat for cat in existing if cat}
    new_set = {cat for cat in normalized if cat}

    if overwrite:
        final = sorted(new_set)
    else:
        if not existing_set and not append:
            final = sorted(new_set)
        else:
            final = sorted(existing_set.union(new_set))

    mail_item.Categories = "; ".join(final)
    mail_item.Save()
    return final

@mcp.tool()
def apply_category(
    categories: Optional[Any] = None,
    category: Optional[str] = None,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    overwrite: bool = False,
    append: bool = False,
) -> str:
    """Apply one or more Outlook categories to a message."""
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

        _, namespace = connect_to_outlook()
        try:
            _, mail_item = _resolve_mail_item(namespace, email_number=email_number, message_id=message_id)
        except ToolError as exc:
            return f"Errore: {exc}"

        try:
            final_categories = _apply_categories_to_item(mail_item, category_list, overwrite_bool, append_bool)
        except ValueError as exc:
            return f"Errore: {exc}"

        if email_number is not None:
            _update_cached_email(email_number, categories="; ".join(final_categories))

        reference = f"#{email_number}" if email_number is not None else (message_id or safe_entry_id(mail_item) or "messaggio")
        return f"Categorie applicate al messaggio {reference}: {', '.join(final_categories) if final_categories else '(nessuna)'}."
    except Exception as exc:
        logger.exception("Errore durante apply_category.")
        return f"Errore durante l'aggiornamento delle categorie: {exc}"

@mcp.tool()
def reply_to_email_by_number(
    email_number: Optional[int] = None,
    reply_text: str = "",
    message_id: Optional[str] = None,
    reply_all: bool = False,
    send: bool = True,
    attachments: Optional[Any] = None,
    use_html: bool = False,
) -> str:
    """Reply to an Outlook message using cached numbering or EntryID."""
    try:
        if not reply_text.strip():
            return "Errore: specifica il testo della risposta."

        reply_all_bool = coerce_bool(reply_all)
        send_bool = coerce_bool(send)
        use_html_bool = coerce_bool(use_html)
        attachment_paths = ensure_string_list(attachments)
        logger.info(
            "reply_to_email_by_number chiamato (numero=%s id=%s reply_all=%s invia=%s allegati=%s html=%s).",
            email_number,
            message_id,
            reply_all_bool,
            send_bool,
            attachment_paths,
            use_html_bool,
        )

        _, namespace = connect_to_outlook()
        try:
            _, mail_item = _resolve_mail_item(namespace, email_number=email_number, message_id=message_id)
        except ToolError as exc:
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
        logger.exception("Errore durante reply_to_email_by_number (numero=%s id=%s).", email_number, message_id)
        return f"Errore durante l'invio della risposta: {exc}"

@mcp.tool()
def compose_email(
    recipient_email: str,
    subject: str,
    body: str,
    cc_email: Optional[str] = None,
    bcc_email: Optional[str] = None,
    attachments: Optional[Any] = None,
    send: bool = True,
    use_html: bool = False,
) -> str:
    """Compose a new Outlook email with optional attachments."""
    try:
        if not recipient_email.strip():
            return "Errore: specifica almeno un destinatario."

        send_bool = coerce_bool(send)
        use_html_bool = coerce_bool(use_html)
        attachment_paths = ensure_string_list(attachments)
        logger.info(
            "compose_email chiamato (destinatario=%s cc=%s bcc=%s oggetto='%s' invia=%s allegati=%s html=%s).",
            recipient_email,
            cc_email,
            bcc_email,
            subject,
            send_bool,
            attachment_paths,
            use_html_bool,
        )

        outlook, _ = connect_to_outlook()
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

        for path_value in attachment_paths:
            absolute = os.path.abspath(path_value)
            if not os.path.exists(absolute):
                return f"Errore: file '{absolute}' non trovato."
            try:
                mail.Attachments.Add(absolute)
            except Exception as exc:
                logger.exception("Impossibile allegare il file %s.", absolute)
                return f"Errore: impossibile allegare '{absolute}' ({exc})."

        if send_bool:
            mail.Send()
            return f"Email inviata a: {recipient_email}"

        mail.Save()
        entry_id = safe_entry_id(mail)
        return f"Bozza salvata (message_id={entry_id or 'N/D'})."
    except Exception as exc:
        logger.exception("Errore durante compose_email per destinatario %s.", recipient_email)
        return f"Errore durante la composizione dell'email: {exc}"

@mcp.tool()
def get_attachments(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    save_to: Optional[str] = None,
    download: bool = False,
    limit: Optional[int] = None,
) -> str:
    """List (and optionally download) attachments from a message."""
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
def search_contacts(
    search_term: Optional[str] = None,
    max_results: int = 50,
) -> str:
    """Search Outlook contacts optionally filtering by a search term."""
    try:
        if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
            return "Errore: 'max_results' deve essere un intero tra 1 e 200."

        if search_term is not None:
            search_display = str(search_term).strip()
            normalized_term = search_display.lower()
        else:
            search_display = None
            normalized_term = ""

        logger.info(
            "search_contacts chiamato (termine=%s max_results=%s).",
            search_display,
            max_results,
        )

        _, namespace = connect_to_outlook()
        try:
            contacts_folder = namespace.GetDefaultFolder(10)  # Contacts default folder
        except Exception as exc:
            logger.exception("Impossibile accedere alla cartella Contatti.")
            return f"Errore: impossibile accedere alla cartella dei contatti ({exc})."

        items = getattr(contacts_folder, "Items", None)
        if not items:
            return "Nessun contatto disponibile."

        try:
            total_count = getattr(items, "Count", None)
        except Exception:
            total_count = None

        def iter_contacts():
            """Yield contact entries from the COM collection."""
            try:
                count = getattr(items, "Count", None)
            except Exception:
                count = None
            if isinstance(count, int) and count > 0 and callable(getattr(items, "__call__", None)):
                for index in range(1, count + 1):
                    try:
                        yield items(index)
                    except Exception:
                        continue
                return
            try:
                iterator = iter(items)
            except TypeError:
                backing = getattr(items, "_items", None)
                if isinstance(backing, (list, tuple)):
                    for entry in backing:
                        yield entry
                    return
                try:
                    current = items.GetFirst()
                except Exception:
                    current = None
                if current is not None:
                    while current:
                        yield current
                        try:
                            current = items.GetNext()
                        except Exception:
                            break
                    return
                try:
                    materialized = list(items)  # type: ignore[arg-type]
                except Exception:
                    return
                for entry in materialized:
                    yield entry
                return
            else:
                for entry in iterator:
                    yield entry

        matches: List[Dict[str, str]] = []

        for contact in iter_contacts():
            if not contact:
                continue

            name_candidates = [
                getattr(contact, "FullName", None),
                getattr(contact, "FileAs", None),
                getattr(contact, "CompanyName", None),
            ]
            display_name = next((value for value in name_candidates if value), "Senza nome")

            email_candidates = [
                getattr(contact, "Email1Address", None),
                getattr(contact, "Email2Address", None),
                getattr(contact, "Email3Address", None),
            ]
            primary_email = next((value for value in email_candidates if value), "")

            phone_candidates = [
                getattr(contact, "MobileTelephoneNumber", None),
                getattr(contact, "BusinessTelephoneNumber", None),
                getattr(contact, "HomeTelephoneNumber", None),
                getattr(contact, "PrimaryTelephoneNumber", None),
            ]
            phone_number = next((value for value in phone_candidates if value), "")

            company = getattr(contact, "CompanyName", "") or ""
            categories = getattr(contact, "Categories", "") or ""

            if normalized_term:
                haystack_parts = [
                    str(display_name),
                    str(primary_email or ""),
                    company,
                    str(phone_number or ""),
                    categories,
                ]
                haystack = " ".join(part.lower() for part in haystack_parts if part)
                if normalized_term not in haystack:
                    continue

            matches.append(
                {
                    "name": str(display_name),
                    "email": str(primary_email).strip() if primary_email else "",
                    "company": company.strip(),
                    "phone": str(phone_number).strip() if phone_number else "",
                }
            )

            if len(matches) >= max_results:
                break

        if not matches:
            return "Nessun contatto corrisponde ai criteri richiesti."

        header_suffix = ""
        if total_count is not None:
            header_suffix = f" su {total_count}"
        lines = [f"Trovati {len(matches)} contatti{header_suffix}.", ""]
        for index, info in enumerate(matches, 1):
            parts = [f"{index}. {info['name']}"]
            if info["email"]:
                parts.append(f"<{info['email']}>")
            details: List[str] = []
            if info["company"]:
                details.append(info["company"])
            if info["phone"]:
                details.append(info["phone"])
            if details:
                parts.append(f"({'; '.join(details)})")
            lines.append(" ".join(parts))

        if normalized_term:
            lines.append("")
            lines.append(f"Filtro applicato: '{search_display}'.")

        return "\n".join(lines)
    except Exception as exc:
        logger.exception("Errore durante search_contacts.")
        return f"Errore durante la ricerca dei contatti: {exc}"

@mcp.tool()
def attach_to_email(
    attachments: Any,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    send: bool = False,
) -> str:
    """Attach local files to an Outlook message (draft or reply)."""
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

@mcp.tool()
def batch_manage_emails(
    email_numbers: Optional[Any] = None,
    message_ids: Optional[Any] = None,
    move_to_folder_id: Optional[str] = None,
    move_to_folder_path: Optional[str] = None,
    move_to_folder_name: Optional[str] = None,
    mark_as: Optional[str] = None,
    delete: bool = False,
) -> str:
    """Apply move/mark/delete actions to multiple messages in a single call."""
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
        logger.info(
            "batch_manage_emails chiamato (numeri=%s ids=%s move=%s mark=%s delete=%s).",
            numbers,
            ids,
            move_requested,
            mark_target,
            delete_bool,
        )

        _, namespace = connect_to_outlook()
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
                _, mail_item = _resolve_mail_item(namespace, email_number=number, message_id=entry_id)
            except ToolError as exc:
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
                if delete_bool and email_cache and number in email_cache:
                    email_cache.pop(number, None)
                else:
                    updates: Dict[str, Any] = {}
                    if move_requested and target_folder:
                        updates["folder_path"] = safe_folder_path(target_folder)
                        updates["id"] = safe_entry_id(mail_item) or reference_id
                    if mark_target is not None:
                        updates["unread"] = mark_target
                    if updates:
                        _update_cached_email(number, **updates)

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

# ---------------------------------------------------------------------------
# Application entrypoints
# ---------------------------------------------------------------------------

def _serialize_tool_metadata(tool: Any) -> Dict[str, Any]:
    """Convert FastMCP tool metadata into plain dicts for JSON responses."""
    return {
        "name": getattr(tool, "name", None),
        "description": getattr(tool, "description", None),
        "input_schema": getattr(tool, "inputSchema", None),
        "output_schema": getattr(tool, "outputSchema", None),
        "annotations": getattr(tool, "annotations", None),
    }


def _serialize_contents(contents: List[Any]) -> List[Dict[str, Any]]:
    """Convert FastMCP content payloads into JSON serializable dictionaries."""
    serialized: List[Dict[str, Any]] = []
    for item in contents:
        if hasattr(item, "__dict__"):
            # Copy to avoid leaking internal references
            serialized.append(dict(item.__dict__))
        else:
            serialized.append({"type": "text", "text": str(item)})
    return serialized


def _verify_outlook_connection() -> None:
    """Validate Outlook availability before starting any server mode."""
    print("Connessione a Outlook...")
    logger.info("Verifica della connessione a Outlook in corso.")
    outlook, namespace = connect_to_outlook()
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
    inbox_items = getattr(getattr(inbox, "Items", None), "Count", "sconosciuto")
    print(f"Connessione a Outlook riuscita. La Posta in arrivo contiene {inbox_items} elementi.")
    logger.info(
        "Connessione a Outlook verificata. La Posta in arrivo contiene %s elementi.",
        inbox_items,
    )
    # Release COM references to avoid locking Outlook instances unnecessarily
    del inbox
    del namespace
    del outlook


def _create_http_app() -> Any:
    """Instantiate the optional FastAPI bridge for HTTP integrations."""
    if (
        FastAPI is None
        or BaseModel is None
        or Field is None
        or HTTPException is None
        or Body is None
        or uvicorn is None
    ):
        raise RuntimeError(
            "La modalita HTTP richiede fastapi, uvicorn e pydantic. "
            "Installa le dipendenze: pip install fastapi uvicorn"
        )

    class ToolCallRequest(BaseModel):
        """Pydantic model describing an HTTP tool invocation."""

        arguments: Dict[str, Any] = Field(default_factory=dict)

    app = FastAPI(
        title="Outlook MCP HTTP Bridge",
        description="Bridge HTTP leggero per richiamare gli strumenti MCP di Outlook da automazioni esterne (es. n8n).",
        version="1.0.0",
    )

    async def _handle_tool_invocation(tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """Shared execution helper for HTTP-triggered tool calls."""
        logger.info("Invocazione HTTP del tool %s con argomenti=%s", tool_name, arguments)
        try:
            contents, output = await mcp.call_tool(tool_name, arguments)
            return {
                "tool": tool_name,
                "content": _serialize_contents(contents),
                "result": output,
            }
        except ToolError as exc:
            logger.warning("Tool %s non trovato per invocazione HTTP: %s", tool_name, exc)
            raise HTTPException(status_code=404, detail=str(exc))  # type: ignore[misc]
        except Exception as exc:  # pylint: disable=broad-except
            logger.exception("Errore durante l'esecuzione HTTP del tool %s.", tool_name)
            raise HTTPException(status_code=500, detail=str(exc))  # type: ignore[misc]

    def _resolve_root_payload(payload: Dict[str, Any]) -> Dict[str, Any]:
        """Normalize POST / payloads into {'tool': str, 'arguments': dict}."""
        preferred_keys = ("tool", "tool_name", "toolName", "name")
        tool_name: Optional[str] = None
        arguments: Any = payload.get("arguments", {})

        for key in preferred_keys:
            value = payload.get(key)
            if isinstance(value, str) and value.strip():
                tool_name = value.strip()
                break

        if tool_name is None:
            candidate_pairs = [
                (key, value)
                for key, value in payload.items()
                if key not in {"arguments"}
                and isinstance(key, str)
                and isinstance(value, dict)
            ]
            if len(candidate_pairs) == 1:
                tool_name, arguments = candidate_pairs[0]

        if tool_name is None:
            raise HTTPException(
                status_code=400,
                detail=(
                    "Specifica il tool da eseguire usando il campo 'tool' (stringa) "
                    "oppure struttura il payload come {\"nome_tool\": {...}}."
                ),
            )

        if not isinstance(arguments, dict):
            raise HTTPException(
                status_code=400,
                detail="Il campo 'arguments' deve essere un oggetto JSON (dizionario).",
            )

        return {"tool": tool_name, "arguments": arguments}

    @app.get("/")
    async def root() -> Dict[str, Any]:
        """Provide a quick-start payload when browsing the root endpoint."""
        tools = await mcp.list_tools()
        return {
            "message": "Outlook MCP HTTP Bridge attivo. Usa POST /tools/{tool_name} oppure POST / con {\"tool\": \"nome\", \"arguments\": {...}}.",
            "available_tools": [tool.name for tool in tools],
        }

    @app.post("/")
    async def invoke_tool_root(payload: Dict[str, Any] = Body(...)) -> Dict[str, Any]:
        """Allow POST / to execute a tool with flexible payload aliases."""
        normalized = _resolve_root_payload(payload)
        return await _handle_tool_invocation(
            normalized["tool"], normalized.get("arguments", {})
        )

    @app.get("/health")
    async def health_check() -> Dict[str, str]:
        """Simple readiness probe for container orchestrators."""
        return {"status": "ok"}

    @app.get("/tools")
    async def list_tools() -> Dict[str, Any]:
        """Return metadata for the registered MCP tools."""
        tools = await mcp.list_tools()
        return {"tools": [_serialize_tool_metadata(tool) for tool in tools]}

    @app.post("/tools/{tool_name}")
    async def invoke_tool(tool_name: str, request: ToolCallRequest) -> Dict[str, Any]:
        """Execute an MCP tool and return the serialized content/result."""
        arguments = request.arguments or {}
        return await _handle_tool_invocation(tool_name, arguments)

    return app


def _start_http_bridge(host: str, port: int) -> None:
    """Run the FastAPI HTTP server for MCP bridge mode."""
    app = _create_http_app()
    logger.info("Avvio del bridge HTTP su http://%s:%s", host, port)
    uvicorn.run(app, host=host, port=port, log_level="info")  # type: ignore[arg-type]


def _build_arg_parser() -> argparse.ArgumentParser:
    """Create CLI parser supporting both MCP and HTTP bridge modes."""
    parser = argparse.ArgumentParser(
        description="Outlook MCP Server - accesso diretto o bridge HTTP."
    )
    parser.add_argument(
        "--mode",
        choices=("mcp", "http"),
        default="mcp",
        help="Modalita di esecuzione: 'mcp' (default) per FastMCP oppure 'http' per il bridge REST.",
    )
    parser.add_argument(
        "--host",
        default="0.0.0.0",
        help="Indirizzo di bind per i server di rete (streamable-http, sse o bridge REST).",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Porta di ascolto per i server di rete (streamable-http, sse o bridge REST).",
    )
    parser.add_argument(
        "--transport",
        choices=("stdio", "sse", "streamable-http"),
        default="stdio",
        help="Trasporto FastMCP quando --mode=mcp. Usa 'streamable-http' per n8n o altri client MCP HTTP.",
    )
    parser.add_argument(
        "--stream-path",
        default="/mcp",
        help="Percorso base per il trasporto streamable-http (solo quando --transport=streamable-http).",
    )
    parser.add_argument(
        "--mount-path",
        default="/",
        help="Percorso di montaggio SSE (solo quando --transport=sse).",
    )
    parser.add_argument(
        "--sse-path",
        default="/sse",
        help="Endpoint SSE relativo (solo quando --transport=sse).",
    )
    parser.add_argument(
        "--skip-outlook-check",
        action="store_true",
        help="Salta il controllo iniziale della connessione a Outlook (sconsigliato).",
    )
    return parser


def main() -> None:
    """Entrypoint principale per il server MCP/bridge HTTP."""
    parser = _build_arg_parser()
    args = parser.parse_args()

    print("Avvio di Outlook MCP Server...")
    logger.info("Outlook MCP Server avviato in modalita %s.", args.mode)

    if not args.skip_outlook_check:
        _verify_outlook_connection()
    else:
        logger.warning("Controllo iniziale di Outlook disabilitato per richiesta esplicita.")

    if args.mode == "mcp":
        transport = args.transport
        if transport in {"sse", "streamable-http"}:
            mcp.settings.host = args.host
            mcp.settings.port = args.port
        if transport == "sse":
            mcp.settings.mount_path = args.mount_path
            mcp.settings.sse_path = args.sse_path
            print(
                f"Avvio del server MCP (SSE) su http://{args.host}:{args.port}{args.sse_path} "
                "(Ctrl+C per interrompere)."
            )
            logger.info(
                "Server MCP avviato in modalita SSE su http://%s:%s%s (mount=%s).",
                args.host,
                args.port,
                args.sse_path,
                args.mount_path,
            )
            mcp.run(transport="sse", mount_path=args.mount_path)
        elif transport == "streamable-http":
            mcp.settings.streamable_http_path = args.stream_path
            print(
                f"Avvio del server MCP (streamable-http) su "
                f"http://{args.host}:{args.port}{args.stream_path} (Ctrl+C per interrompere)."
            )
            logger.info(
                "Server MCP avviato in modalita streamable-http su http://%s:%s%s.",
                args.host,
                args.port,
                args.stream_path,
            )
            mcp.run(transport="streamable-http")
        else:
            print("Avvio del server MCP (stdio). Premi Ctrl+C per interrompere.")
            logger.info("Server MCP avviato su stdio.")
            mcp.run()
    else:
        print(f"Avvio del bridge HTTP su http://{args.host}:{args.port} (Ctrl+C per interrompere).")
        _start_http_bridge(args.host, args.port)


# Run the server
if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # pylint: disable=broad-except
        print(f"Errore durante l'avvio del server: {str(exc)}")
        logger.exception("Errore durante l'avvio di Outlook MCP Server.")
