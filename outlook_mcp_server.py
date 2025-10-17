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
from outlook_mcp.features import feature_gate, is_tool_enabled, get_tool_group

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

        lines.append(f"{prefix} {timestamp} Â· {sender}: {subject}{preview_line}")

    # Normalize any stray non-ASCII separator that may appear in some consoles
    output = "\n".join(lines)
    try:
        output = output.replace(" �� ", " -> ")
    except Exception:
        pass
    return output

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
        tool_group = get_tool_group(getattr(tool, "name", ""))
        if not is_tool_enabled(getattr(tool, "name", ""), tool_group):
            continue
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
    from outlook_mcp.tools.folders import list_folders as _impl
    return _impl(
        root_folder_id=root_folder_id,
        root_folder_path=root_folder_path,
        root_folder_name=root_folder_name,
        max_depth=max_depth,
        include_counts=include_counts,
        include_ids=include_ids,
        include_store=include_store,
        include_paths=include_paths,
    )

def get_folder_metadata(
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
    include_children: bool = False,
    max_children: int = 20,
    include_counts: bool = True,
) -> str:
    from outlook_mcp.tools.folders import get_folder_metadata as _impl
    return _impl(
        folder_id=folder_id,
        folder_path=folder_path,
        folder_name=folder_name,
        include_children=include_children,
        max_children=max_children,
        include_counts=include_counts,
    )

def create_folder(
    new_folder_name: str,
    parent_folder_id: Optional[str] = None,
    parent_folder_path: Optional[str] = None,
    parent_folder_name: Optional[str] = None,
    item_type: Optional[Any] = None,
    allow_existing: bool = False,
) -> str:
    from outlook_mcp.tools.folders import create_folder as _impl
    return _impl(
        new_folder_name=new_folder_name,
        parent_folder_id=parent_folder_id,
        parent_folder_path=parent_folder_path,
        parent_folder_name=parent_folder_name,
        item_type=item_type,
        allow_existing=allow_existing,
    )

def rename_folder(
    folder_id: Optional[str] = None,
    new_name: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
) -> str:
    from outlook_mcp.tools.folders import rename_folder as _impl
    return _impl(
        folder_id=folder_id,
        new_name=new_name,
        folder_path=folder_path,
        folder_name=folder_name,
    )

def delete_folder(
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
    confirm: bool = False,
) -> str:
    from outlook_mcp.tools.folders import delete_folder as _impl
    return _impl(
        folder_id=folder_id,
        folder_path=folder_path,
        folder_name=folder_name,
        confirm=confirm,
    )


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
    from outlook_mcp.tools.email_list import list_recent_emails as _impl
    return _impl(
        days=days,
        folder_name=folder_name,
        max_results=max_results,
        include_preview=include_preview,
        include_all_folders=include_all_folders,
        folder_ids=folder_ids,
        folder_paths=folder_paths,
        offset=offset,
        unread_only=unread_only,
    )

def list_sent_emails(
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    offset: int = 0,
) -> str:
    from outlook_mcp.tools.email_list import list_sent_emails as _impl
    return _impl(days=days, folder_name=folder_name, max_results=max_results, include_preview=include_preview, offset=offset)


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
    from outlook_mcp.tools.email_list import search_emails as _impl
    return _impl(
        search_term=search_term,
        days=days,
        folder_name=folder_name,
        max_results=max_results,
        include_preview=include_preview,
        include_all_folders=include_all_folders,
        folder_ids=folder_ids,
        folder_paths=folder_paths,
        offset=offset,
        unread_only=unread_only,
    )

def list_pending_replies(
    days: int = 14,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
    include_unread_only: bool = False,
    conversation_lookback_days: Optional[int] = None,
) -> str:
    from outlook_mcp.tools.email_list import list_pending_replies as _impl
    return _impl(days=days, folder_name=folder_name, max_results=max_results, include_preview=include_preview, include_all_folders=include_all_folders, include_unread_only=include_unread_only, conversation_lookback_days=conversation_lookback_days)
def list_upcoming_events(
    days: int = 7,
    calendar_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_description: bool = False,
    include_all_calendars: bool = False,
) -> str:
    from outlook_mcp.tools.calendar_read import list_upcoming_events as _impl
    return _impl(days=days, calendar_name=calendar_name, max_results=max_results, include_description=include_description, include_all_calendars=include_all_calendars)
def search_calendar_events(
    search_term: str,
    days: int = 30,
    calendar_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_description: bool = False,
    include_all_calendars: bool = False,
) -> str:
    from outlook_mcp.tools.calendar_read import search_calendar_events as _impl
    return _impl(search_term=search_term, days=days, calendar_name=calendar_name, max_results=max_results, include_description=include_description, include_all_calendars=include_all_calendars)
def get_event_by_number(event_number: int) -> str:
    from outlook_mcp.tools.calendar_read import get_event_by_number as _impl
    return _impl(event_number)



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
    from outlook_mcp.tools.calendar_write import create_calendar_event as _impl
    return _impl(
        subject=subject,
        start_time=start_time,
        duration_minutes=duration_minutes,
        location=location,
        body=body,
        attendees=attendees,
        reminder_minutes=reminder_minutes,
        calendar_name=calendar_name,
        all_day=all_day,
        send_invitations=send_invitations,
    )
def ensure_domain_folder(
    email_number: Optional[int] = None,
    sender_email: Optional[str] = None,
    root_folder_name: Optional[str] = None,
    subfolders: Optional[str] = None,
) -> str:
    from outlook_mcp.tools.domain_rules import ensure_domain_folder as _impl
    return _impl(email_number=email_number, sender_email=sender_email, root_folder_name=root_folder_name, subfolders=subfolders)


def move_email_to_domain_folder(
    email_number: int,
    root_folder_name: Optional[str] = None,
    create_if_missing: bool = True,
    subfolders: Optional[str] = None,
) -> str:
    from outlook_mcp.tools.domain_rules import move_email_to_domain_folder as _impl
    return _impl(
        email_number=email_number,
        root_folder_name=root_folder_name,
        create_if_missing=create_if_missing,
        subfolders=subfolders,
    )


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

def move_email_to_folder(
    target_folder_id: Optional[str] = None,
    target_folder_path: Optional[str] = None,
    target_folder_name: Optional[str] = None,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    create_if_missing: bool = False,
) -> str:
    from outlook_mcp.tools.email_actions import move_email_to_folder as _impl
    return _impl(target_folder_id=target_folder_id, target_folder_path=target_folder_path, target_folder_name=target_folder_name, email_number=email_number, message_id=message_id, create_if_missing=create_if_missing)
def mark_email_read_unread(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    unread: Optional[bool] = None,
    flag: Optional[str] = None,
) -> str:
    from outlook_mcp.tools.email_actions import mark_email_read_unread as _impl
    return _impl(email_number=email_number, message_id=message_id, unread=unread, flag=flag)
def apply_category(
    categories: Optional[Any] = None,
    category: Optional[str] = None,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    overwrite: bool = False,
    append: bool = False,
) -> str:
    from outlook_mcp.tools.email_actions import apply_category as _impl
    return _impl(categories=categories, category=category, email_number=email_number, message_id=message_id, overwrite=overwrite, append=append)
def reply_to_email_by_number(
    email_number: Optional[int] = None,
    reply_text: str = "",
    message_id: Optional[str] = None,
    reply_all: bool = False,
    send: bool = True,
    attachments: Optional[Any] = None,
    use_html: bool = False,
) -> str:
    from outlook_mcp.tools.email_actions import reply_to_email_by_number as _impl
    return _impl(email_number=email_number, reply_text=reply_text, message_id=message_id, reply_all=reply_all, send=send, attachments=attachments, use_html=use_html)
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
    from outlook_mcp.tools.email_actions import compose_email as _impl
    return _impl(recipient_email=recipient_email, subject=subject, body=body, cc_email=cc_email, bcc_email=bcc_email, attachments=attachments, send=send, use_html=use_html)
def get_attachments(
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    save_to: Optional[str] = None,
    download: bool = False,
    limit: Optional[int] = None,
) -> str:
    from outlook_mcp.tools.attachments import get_attachments as _impl
    return _impl(email_number=email_number, message_id=message_id, save_to=save_to, download=download, limit=limit)


def search_contacts(
    search_term: Optional[str] = None,
    max_results: int = 50,
) -> str:
    from outlook_mcp.tools.contacts import search_contacts as _impl
    return _impl(search_term=search_term, max_results=max_results)

def attach_to_email(
    attachments: Any,
    email_number: Optional[int] = None,
    message_id: Optional[str] = None,
    send: bool = False,
) -> str:
    from outlook_mcp.tools.attachments import attach_to_email as _impl
    return _impl(attachments=attachments, email_number=email_number, message_id=message_id, send=send)

def batch_manage_emails(
    email_numbers: Optional[Any] = None,
    message_ids: Optional[Any] = None,
    move_to_folder_id: Optional[str] = None,
    move_to_folder_path: Optional[str] = None,
    move_to_folder_name: Optional[str] = None,
    mark_as: Optional[str] = None,
    delete: bool = False,
) -> str:
    from outlook_mcp.tools.email_actions import batch_manage_emails as _impl
    return _impl(email_numbers=email_numbers, message_ids=message_ids, move_to_folder_id=move_to_folder_id, move_to_folder_path=move_to_folder_path, move_to_folder_name=move_to_folder_name, mark_as=mark_as, delete=delete)
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
            # Feature-gate check before actual invocation
            group = get_tool_group(tool_name)
            if not is_tool_enabled(tool_name, group):
                raise HTTPException(status_code=403, detail=f"Tool '{tool_name}' disabilitato dal server")  # type: ignore[misc]
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
        visible = [t for t in tools if is_tool_enabled(t.name, get_tool_group(t.name))]
        return {
            "message": "Outlook MCP HTTP Bridge attivo. Usa POST /tools/{tool_name} oppure POST / con {\"tool\": \"nome\", \"arguments\": {...}}.",
            "available_tools": [tool.name for tool in visible],
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
        visible = [t for t in tools if is_tool_enabled(t.name, get_tool_group(t.name))]
        return {"tools": [_serialize_tool_metadata(tool) for tool in visible]}

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





