import datetime
import logging
import os
import win32com.client
from logging.handlers import RotatingFileHandler
from typing import Any, Dict, List, Optional, Set
from mcp.server.fastmcp import FastMCP, Context

# Initialize FastMCP server
mcp = FastMCP("outlook-assistant")

# Constants
MAX_DAYS = 30
# Email cache for storing retrieved emails by number
email_cache = {}
calendar_cache = {}
BODY_PREVIEW_MAX_CHARS = 220
DEFAULT_MAX_RESULTS = 25
ATTACHMENT_NAME_PREVIEW_MAX = 5
CONVERSATION_ID_PREVIEW_MAX = 16
LOG_DIR_NAME = "logs"
LOG_FILE_NAME = "outlook_mcp_server.log"
MAX_EVENT_LOOKAHEAD_DAYS = 90
PR_LAST_VERB_EXECUTED = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
PR_LAST_VERB_EXECUTION_TIME = "http://schemas.microsoft.com/mapi/proptag/0x10820040"
LAST_VERB_REPLY_CODES = {102, 103}
DEFAULT_CONVERSATION_SAMPLE_LIMIT = 15
MAX_CONVERSATION_LOOKBACK_DAYS = 180
PENDING_SCAN_MULTIPLIER = 4

def _setup_logger() -> logging.Logger:
    """Configure application-wide logging once."""
    logger = logging.getLogger("outlook_mcp_server")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(base_dir, LOG_DIR_NAME)
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, LOG_FILE_NAME)

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
    )

    rotating_handler = RotatingFileHandler(
        log_path,
        maxBytes=5 * 1024 * 1024,  # 5 MB per file
        backupCount=3,
        encoding="utf-8",
    )
    rotating_handler.setFormatter(formatter)
    logger.addHandler(rotating_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    logger.debug("Logger inizializzato: scrittura su %s", log_path)
    return logger

logger = _setup_logger()

def _trim_conversation_id(conversation_id: Optional[str], max_chars: int = CONVERSATION_ID_PREVIEW_MAX) -> Optional[str]:
    """Shorten long conversation identifiers so they stay readable."""
    if not conversation_id:
        return None
    if len(conversation_id) <= max_chars:
        return conversation_id
    return conversation_id[:max_chars] + "..."

def _coerce_bool(value: Any) -> bool:
    """Best-effort conversion of user-provided values into booleans."""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "y", "yes", "on"}
    return bool(value)

def _normalize_whitespace(text: Optional[str]) -> str:
    """Collapse whitespace so previews stay compact."""
    if not text:
        return ""
    return " ".join(text.split())

def _build_body_preview(body: Optional[str], max_chars: int = BODY_PREVIEW_MAX_CHARS) -> str:
    """Create a trimmed preview of the email body for quick inspection."""
    normalized = _normalize_whitespace(body)
    if not normalized:
        return ""
    if len(normalized) <= max_chars:
        return normalized
    return normalized[: max_chars - 3].rstrip() + "..."

def _extract_recipients(mail_item) -> Dict[str, List[str]]:
    """Return recipients grouped by address type."""
    recipients_by_type = {"to": [], "cc": [], "bcc": []}
    if not hasattr(mail_item, "Recipients") or not mail_item.Recipients:
        return recipients_by_type

    type_mapping = {1: "to", 2: "cc", 3: "bcc"}  # Outlook constants
    for i in range(1, mail_item.Recipients.Count + 1):
        recipient = mail_item.Recipients(i)
        display_name = recipient.Name or "Sconosciuto"
        address = getattr(recipient, "Address", "") or ""
        formatted = f"{display_name} <{address}>" if address else display_name
        address_type = type_mapping.get(getattr(recipient, "Type", 1), "to")
        recipients_by_type[address_type].append(formatted)
    return recipients_by_type

def _safe_folder_path(mail_item) -> str:
    """Return a readable folder path if available."""
    try:
        parent = getattr(mail_item, "Parent", None)
        return parent.FolderPath if parent else ""
    except Exception:
        return ""

def _extract_attachment_names(mail_item, max_names: int = ATTACHMENT_NAME_PREVIEW_MAX) -> List[str]:
    """Return a small list of attachment names without downloading them."""
    names: List[str] = []
    if not hasattr(mail_item, "Attachments"):
        return names
    try:
        attachment_count = mail_item.Attachments.Count
    except Exception:
        attachment_count = 0
    if not attachment_count:
        return names

    for index in range(1, min(attachment_count, max_names) + 1):
        try:
            names.append(mail_item.Attachments(index).FileName)
        except Exception:
            continue
    if attachment_count > max_names:
        names.append(f"... (+{attachment_count - max_names} more)")
    return names

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

def _to_python_datetime(value: Any) -> Optional[datetime.datetime]:
    """Best-effort conversion from COM datetime to naive datetime."""
    if not value:
        return None
    try:
        return datetime.datetime(
            value.year,
            value.month,
            value.day,
            value.hour,
            value.minute,
            value.second,
        )
    except Exception:
        try:
            return datetime.datetime.strptime(
                value.strftime("%Y-%m-%d %H:%M:%S"), "%Y-%m-%d %H:%M:%S"
            )
        except Exception:
            return None

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
        last_time = _to_python_datetime(getattr(mail_item, "LastVerbExecutionTime", None))
        if not baseline or not last_time or last_time >= baseline:
            return True

    accessor = getattr(mail_item, "PropertyAccessor", None)
    if accessor:
        try:
            verb_value = accessor.GetProperty(PR_LAST_VERB_EXECUTED)
            if isinstance(verb_value, int) and verb_value in LAST_VERB_REPLY_CODES:
                time_value = accessor.GetProperty(PR_LAST_VERB_EXECUTION_TIME)
                time_dt = _to_python_datetime(time_value)
                if not baseline or not time_dt or time_dt >= baseline:
                    return True
        except Exception:
            pass
    return False

# Helper functions
def connect_to_outlook():
    """Connect to Outlook application using COM"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        logger.debug("Connessione a Outlook MAPI completata.")
        return outlook, namespace
    except Exception as e:
        logger.exception("Errore durante la connessione a Outlook.")
        raise Exception(f"Impossibile connettersi a Outlook: {str(e)}")

def get_folder_by_name(namespace, folder_name: str):
    """Get a specific Outlook folder by name"""
    try:
        # First check inbox subfolder
        inbox = namespace.GetDefaultFolder(6)  # 6 is the index for inbox folder
        
        # Check inbox subfolders first (most common)
        for folder in inbox.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
                
        # Then check all folders at root level
        for folder in namespace.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
            
            # Also check subfolders
            for subfolder in folder.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    return subfolder
                    
        # If not found
        return None
    except Exception as e:
        logger.exception("Impossibile accedere alla cartella '%s'.", folder_name)
        raise Exception(f"Impossibile accedere alla cartella {folder_name}: {str(e)}")

def format_email(mail_item) -> Dict[str, Any]:
    """Format an Outlook mail item into a structured dictionary"""
    try:
        # Extract recipients grouped by type
        recipients_by_type = _extract_recipients(mail_item)
        all_recipients = (
            recipients_by_type["to"]
            + recipients_by_type["cc"]
            + recipients_by_type["bcc"]
        )

        # Capture body and preview
        body_content = getattr(mail_item, "Body", "") or ""
        preview = _build_body_preview(body_content)

        # Prepare received time representations
        received_iso = None
        received_display = None
        if hasattr(mail_item, "ReceivedTime") and mail_item.ReceivedTime:
            received_dt = _to_python_datetime(mail_item.ReceivedTime)
            if received_dt:
                received_display = received_dt.strftime("%Y-%m-%d %H:%M:%S")
                received_iso = received_dt.strftime("%Y-%m-%dT%H:%M:%S")
            else:
                received_display = str(mail_item.ReceivedTime)
                received_iso = received_display

        sent_iso = None
        sent_display = None
        if hasattr(mail_item, "SentOn") and mail_item.SentOn:
            sent_dt = _to_python_datetime(mail_item.SentOn)
            if sent_dt:
                sent_display = sent_dt.strftime("%Y-%m-%d %H:%M:%S")
                sent_iso = sent_dt.strftime("%Y-%m-%dT%H:%M:%S")
            else:
                sent_display = str(mail_item.SentOn)
                sent_iso = sent_display

        last_modified_iso = None
        last_modified_display = None
        if hasattr(mail_item, "LastModificationTime") and mail_item.LastModificationTime:
            last_dt = _to_python_datetime(mail_item.LastModificationTime)
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
                    attachment_names = _extract_attachment_names(mail_item)
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
            "folder_path": _safe_folder_path(mail_item),
        }
        return email_data
    except Exception as e:
        logger.exception("Impossibile formattare il messaggio con EntryID=%s.", getattr(mail_item, "EntryID", "Sconosciuto"))
        raise Exception(f"Impossibile formattare il messaggio: {str(e)}")

def clear_email_cache():
    """Clear the email cache"""
    global email_cache
    email_cache = {}
    logger.debug("Cache dei messaggi svuotata.")

def clear_calendar_cache():
    """Clear the calendar event cache"""
    global calendar_cache
    calendar_cache = {}
    logger.debug("Cache degli eventi svuotata.")

def get_emails_from_folder(folder, days: int, search_term: Optional[str] = None):
    """Get emails from a folder with optional search filter"""
    emails_list = []
    
    # Calculate the date threshold
    now = datetime.datetime.now()
    threshold_date = now - datetime.timedelta(days=days)
    
    try:
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
                        continue
                    
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
            folder = get_folder_by_name(namespace, name)
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
    """Return a flat list of all accessible mail folders."""
    folders = []
    queue = []
    try:
        for root_folder in namespace.Folders:
            queue.append(root_folder)
    except Exception:
        logger.warning("Impossibile enumerare le cartelle principali dell'account Outlook.")
        return folders

    while queue:
        folder = queue.pop(0)
        folders.append(folder)
        try:
            for subfolder in folder.Folders:
                queue.append(subfolder)
        except Exception:
            continue
    logger.debug("Rilevate %s cartelle complessive per la scansione globale.", len(folders))
    return folders

def collect_emails_across_folders(folders: List, days: int, search_term: Optional[str] = None) -> List[Dict[str, Any]]:
    """Aggregate emails from multiple folders into a single newest-first list."""
    aggregated: Dict[str, Dict[str, Any]] = {}
    for folder in folders:
        try:
            folder_emails = get_emails_from_folder(folder, days, search_term)
        except Exception:
            logger.debug("Cartella ignorata durante la raccolta globale: %s", getattr(folder, "FolderPath", folder))
            continue

        for email in folder_emails:
            email_id = email.get("id")
            if not email_id:
                continue
            if email_id not in aggregated:
                aggregated[email_id] = email

    # Sort by ISO timestamp if available, otherwise fallback to received_time string
    sorted_emails = sorted(
        aggregated.values(),
        key=lambda item: item.get("received_iso") or item.get("received_time") or "",
        reverse=True,
    )
    logger.info(
        "Raccolti %s messaggi totali attraversando %s cartelle.",
        len(sorted_emails),
        len(folders),
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
        start_dt = _to_python_datetime(getattr(appointment, "Start", None))
        end_dt = _to_python_datetime(getattr(appointment, "End", None))

        def fmt(dt: Optional[datetime.datetime]) -> Optional[str]:
            if not dt:
                return None
            return dt.strftime("%Y-%m-%d %H:%M")

        start_iso = start_dt.strftime("%Y-%m-%dT%H:%M") if start_dt else None
        end_iso = end_dt.strftime("%Y-%m-%dT%H:%M") if end_dt else None

        required = getattr(appointment, "RequiredAttendees", "") or ""
        optional = getattr(appointment, "OptionalAttendees", "") or ""
        body = getattr(appointment, "Body", "") or ""
        preview = _build_body_preview(body, max_chars=320)

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
            "folder_path": _safe_folder_path(appointment),
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
            start_dt = _to_python_datetime(getattr(appointment, "Start", None))
            end_dt = _to_python_datetime(getattr(appointment, "End", None))

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
    if not email_data:
        return False

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
            except Exception:
                logger.debug(
                    "Errore durante la ricerca dei messaggi correlati per %s.",
                    email_data.get("id"),
                    exc_info=True,
                )

        for related in related_entries:
            sender_email = _normalize_email_address(related.get("sender_email")) or _normalize_email_address(
                related.get("sender")
            )
            if not sender_email or sender_email not in normalized_user_addresses:
                continue
            related_dt = _extract_best_timestamp(related)
            if not baseline_dt or not related_dt or related_dt >= baseline_dt:
                return True

    if mail_item and _mail_item_marked_replied(mail_item, baseline_dt):
        return True

    return False

# MCP Tools
@mcp.tool()
def get_current_datetime(include_utc: bool = True) -> str:
    """
    Restituisce la data e ora correnti formattate.

    Args:
        include_utc: Includere o meno il riferimento UTC nella risposta
    """
    include_utc_bool = _coerce_bool(include_utc)
    logger.info("get_current_datetime chiamato con include_utc=%s", include_utc_bool)
    try:
        local_dt = datetime.datetime.now()
        lines = [
            "Data e ora correnti:",
            f"- Locale: {local_dt.strftime('%Y-%m-%d %H:%M:%S')}",
            f"- Locale ISO: {local_dt.isoformat()}",
        ]
        if include_utc_bool:
            utc_dt = datetime.datetime.utcnow().replace(tzinfo=datetime.timezone.utc)
            lines.append(f"- UTC: {utc_dt.strftime('%Y-%m-%d %H:%M:%S')}")
            lines.append(f"- UTC ISO: {utc_dt.isoformat()}")
        return "\n".join(lines)
    except Exception as exc:
        logger.exception("Errore durante get_current_datetime.")
        return f"Errore durante il calcolo della data/ora corrente: {exc}"


@mcp.tool()
def list_folders() -> str:
    """
    List all available mail folders in Outlook
    
    Returns:
        A list of available mail folders
    """
    try:
        logger.info("list_folders chiamato.")
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        result = "Cartelle di posta disponibili:\n\n"
        
        # List all root folders and their subfolders
        for folder in namespace.Folders:
            result += f"- {folder.Name}\n"
            
            # List subfolders
            for subfolder in folder.Folders:
                result += f"  - {subfolder.Name}\n"
                
                # List subfolders (one more level)
                try:
                    for subsubfolder in subfolder.Folders:
                        result += f"    - {subsubfolder.Name}\n"
                except:
                    pass
        
        return result
    except Exception as e:
        logger.exception("Errore durante l'elenco delle cartelle di Outlook.")
        return f"Errore durante l'elenco delle cartelle: {str(e)}"

def _present_email_listing(
    emails: List[Dict[str, Any]],
    folder_display: str,
    days: int,
    max_results: int,
    include_preview: bool,
    log_context: str,
    search_term: Optional[str] = None,
    focus_on_recipients: bool = False,
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

    visible_emails = emails[:max_results]
    visible_count = len(visible_emails)
    total_count = len(emails)

    if search_term:
        if total_count > visible_count:
            header = (
                f"Trovati {total_count} messaggi che corrispondono a '{search_term}' in {folder_display} "
                f"negli ultimi {days} giorni. Mostro i primi {visible_count} risultati."
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
                f"Mostro i primi {visible_count} risultati."
            )
        else:
            header = f"Trovati {visible_count} messaggi in {folder_display} negli ultimi {days} giorni."

    logger.info(
        "%s: restituiti %s messaggi su %s (termine=%s cartella=%s).",
        log_context,
        visible_count,
        total_count,
        search_term,
        folder_display,
    )

    result = header + "\n\n"

    for idx, email in enumerate(visible_emails, 1):
        email_cache[idx] = email

        folder_path = email.get("folder_path") or folder_display
        importance_label = email.get("importance_label") or _describe_importance(email.get("importance"))
        trimmed_conv = _trim_conversation_id(email.get("conversation_id"))
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

    result += (
        "Per leggere il contenuto completo usa lo strumento get_email_by_number con il numero del messaggio.\n"
        "Per ottenere il contesto della conversazione usa lo strumento get_email_context."
    )
    return result

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

    result += (
        "Per leggere il dettaglio completo usa lo strumento get_event_by_number con il numero dell'evento."
    )
    return result

@mcp.tool()
def list_recent_emails(
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
) -> str:
    """
    List email titles from the specified number of days
    
    Args:
        days: Number of days to look back for emails (max 30)
        folder_name: Name of the folder to check (if not specified, checks the Inbox)
        max_results: Maximum number of emails to display (1-200)
        include_preview: Include a trimmed body preview for each email
        include_all_folders: Scan every Outlook folder (ignores folder_name)
        
    Returns:
        Numbered list of email titles with sender information
    """
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        logger.warning("Valore 'days' non valido passato a list_recent_emails: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a list_recent_emails: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"

    include_preview = _coerce_bool(include_preview)
    include_all = _coerce_bool(include_all_folders)
    logger.info(
        "list_recent_emails chiamato con giorni=%s cartella=%s max_risultati=%s anteprima=%s tutte_le_cartelle=%s",
        days,
        folder_name,
        max_results,
        include_preview,
        include_all,
    )
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        if include_all:
            if folder_name:
                logger.info("Parametro folder_name ignorato perché tutte_le_cartelle=True.")
            folders = get_all_mail_folders(namespace)
            emails = collect_emails_across_folders(folders, days)
            folder_display = "Tutte le cartelle"
        else:
            # Get the appropriate folder
            if folder_name:
                folder = get_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella '{folder_name}' non trovata"
            else:
                folder = namespace.GetDefaultFolder(6)  # Default inbox
            folder_display = f"'{folder_name}'" if folder_name else "Posta in arrivo"
            emails = get_emails_from_folder(folder, days)

        return _present_email_listing(
            emails=emails,
            folder_display=folder_display,
            days=days,
            max_results=max_results,
            include_preview=include_preview,
            log_context="list_recent_emails",
        )
    
    except Exception as e:
        logger.exception("Errore nel recupero dei messaggi per la cartella '%s'.", folder_name or "Posta in arrivo")
        return f"Errore durante il recupero dei messaggi: {str(e)}"

@mcp.tool()
def list_sent_emails(
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
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

    include_preview = _coerce_bool(include_preview)
    logger.info(
        "list_sent_emails chiamato con giorni=%s cartella=%s max_risultati=%s anteprima=%s",
        days,
        folder_name,
        max_results,
        include_preview,
    )

    try:
        _, namespace = connect_to_outlook()

        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
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
) -> str:
    """
    Search emails by contact name or keyword within a time period
    
    Args:
        search_term: Name or keyword to search for
        days: Number of days to look back (max 30)
        folder_name: Name of the folder to search (if not specified, searches the Inbox)
        max_results: Maximum number of emails to display (1-200)
        include_preview: Include a trimmed body preview for each email
        include_all_folders: Scan every Outlook folder (ignores folder_name)
        
    Returns:
        Numbered list of matching email titles
    """
    if not search_term:
        logger.warning("search_emails chiamato senza termine di ricerca.")
        return "Errore: inserisci un termine di ricerca"
        
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        logger.warning("Valore 'days' non valido passato a search_emails: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a search_emails: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"

    include_preview = _coerce_bool(include_preview)
    include_all = _coerce_bool(include_all_folders)
    logger.info(
        "search_emails chiamato con termine='%s' giorni=%s cartella=%s max_risultati=%s anteprima=%s tutte_le_cartelle=%s",
        search_term,
        days,
        folder_name,
        max_results,
        include_preview,
        include_all,
    )
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the appropriate folder
        if include_all:
            if folder_name:
                logger.info("Parametro folder_name ignorato perché tutte_le_cartelle=True.")
            folders = get_all_mail_folders(namespace)
            emails = collect_emails_across_folders(folders, days, search_term)
            folder_display = "Tutte le cartelle"
        else:
            if folder_name:
                folder = get_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella '{folder_name}' non trovata"
            else:
                folder = namespace.GetDefaultFolder(6)  # Default inbox
            folder_display = f"'{folder_name}'" if folder_name else "Posta in arrivo"
            emails = get_emails_from_folder(folder, days, search_term)
        return _present_email_listing(
            emails=emails,
            folder_display=folder_display,
            days=days,
            max_results=max_results,
            include_preview=include_preview,
            log_context="search_emails",
            search_term=search_term,
        )
    
    except Exception as e:
        logger.exception(
            "Errore durante la ricerca dei messaggi con termine '%s' nella cartella '%s'.",
            search_term,
            folder_name or "Posta in arrivo",
        )
        return f"Errore nella ricerca dei messaggi: {str(e)}"

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

    include_preview_bool = _coerce_bool(include_preview)
    include_all_bool = _coerce_bool(include_all_folders)
    unread_only_bool = _coerce_bool(include_unread_only)

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
            candidate_emails = collect_emails_across_folders(folders, days)
            folder_display = "Tutte le cartelle (senza risposta)"
        else:
            if folder_name:
                folder = get_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella '{folder_name}' non trovata"
                candidate_emails = get_emails_from_folder(folder, days)
                folder_display = f"'{folder_name}' (senza risposta)"
            else:
                folder = namespace.GetDefaultFolder(6)
                candidate_emails = get_emails_from_folder(folder, days)
                folder_display = "Posta in arrivo (senza risposta)"

        pending_emails: List[Dict[str, Any]] = []
        processed = 0
        max_processed_before_break = max(max_results * PENDING_SCAN_MULTIPLIER, max_results + 25)
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
            try:
                already_replied = _email_has_user_reply(
                    namespace=namespace,
                    email_data=email,
                    user_addresses=user_addresses,
                    conversation_limit=DEFAULT_CONVERSATION_SAMPLE_LIMIT,
                    lookback_days=lookback_days,
                )
            except Exception:
                logger.debug(
                    "Controllo risposta fallito per il messaggio %s.",
                    email.get("id"),
                    exc_info=True,
                )

            if not already_replied:
                pending_emails.append(email)

            if len(pending_emails) >= max_results and processed >= max_processed_before_break:
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

    include_desc = _coerce_bool(include_description)
    include_all = _coerce_bool(include_all_calendars)
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

    include_desc = _coerce_bool(include_description)
    include_all = _coerce_bool(include_all_calendars)
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
def get_email_by_number(email_number: int) -> str:
    """
    Get detailed content of a specific email by its number from the last listing
    
    Args:
        email_number: The number of the email from the list results
        
    Returns:
        Full details of the specified email
    """
    try:
        if not email_cache:
            logger.warning("get_email_by_number chiamato ma la cache e vuota.")
            return "Errore: nessun elenco disponibile. Usa prima list_recent_emails o search_emails."
        
        if email_number not in email_cache:
            logger.warning("Messaggio numero %s non presente in cache per get_email_by_number.", email_number)
            return f"Errore: il messaggio #{email_number} non e presente nell'elenco corrente."
        
        email_data = email_cache[email_number]
        logger.info("Recupero dettagli completi per il messaggio #%s.", email_number)
        
        # Connect to Outlook to get the full email content
        _, namespace = connect_to_outlook()
        
        # Retrieve the specific email
        email = namespace.GetItemFromID(email_data["id"])
        if not email:
            return f"Errore: il messaggio #{email_number} non puo essere recuperato da Outlook."
        
        trimmed_conv = _trim_conversation_id(email_data.get("conversation_id"), max_chars=32)
        importance_label = email_data.get("importance_label") or _describe_importance(email_data.get("importance"))
        attachment_names_preview = email_data.get("attachment_names") or []
        to_line = ", ".join(email_data.get("to_recipients", []))
        cc_line = ", ".join(email_data.get("cc_recipients", []))
        bcc_line = ", ".join(email_data.get("bcc_recipients", []))
        
        result_lines = [
            f"Dettagli del messaggio #{email_number}:",
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

        result_lines.extend(
            [
                f"Ricevuto: {email_data.get('received_time', 'Sconosciuto')}",
                f"Cartella: {email_data.get('folder_path', 'Cartella sconosciuta')}",
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

        result_lines.append(f"Allegati: {_yes_no(email_data.get('has_attachments'))}")
        if attachment_names_preview:
            result_lines.append(f"Nomi allegati: {', '.join(attachment_names_preview)}")

        attachment_lines = []
        if email_data.get("has_attachments") and hasattr(email, "Attachments"):
            try:
                for i in range(1, email.Attachments.Count + 1):
                    attachment = email.Attachments(i)
                    attachment_lines.append(f"  - {attachment.FileName}")
            except Exception:
                pass

        result_lines.append("")
        if attachment_lines:
            result_lines.append("Allegati:")
            result_lines.extend(attachment_lines)
            result_lines.append("")

        result_lines.append("Corpo:")
        result_lines.append(email_data.get("body", "(Nessun contenuto)"))
        
        result_lines.append("")
        result_lines.append(
            "Per rispondere a questo messaggio usa lo strumento reply_to_email_by_number con questo numero."
        )
        
        return "\n".join(result_lines)
    
    except Exception as e:
        logger.exception("Errore nel recupero dei dettagli per il messaggio #%s.", email_number)
        return f"Errore durante il recupero dei dettagli del messaggio: {str(e)}"

@mcp.tool()
def get_event_by_number(event_number: int) -> str:
    """
    Recupera i dettagli completi di un evento dall'ultimo elenco.
    """
    try:
        if not calendar_cache:
            logger.warning("get_event_by_number chiamato ma la cache eventi e vuota.")
            return "Errore: nessun elenco eventi disponibile. Usa prima list_upcoming_events o search_calendar_events."

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
            return "Errore: nessun elenco disponibile. Usa prima list_recent_emails o search_emails."
        
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

        include_thread_bool = _coerce_bool(include_thread)
        include_sent_bool = _coerce_bool(include_sent)
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
            trimmed_conv = _trim_conversation_id(email_data["conversation_id"], max_chars=32)
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
        context_lines.append("Per il corpo completo usa get_email_by_number con questo numero di messaggio.")

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
            "Suggerimento: usa reply_to_email_by_number per rispondere oppure compose_email per iniziare un nuovo thread."
        )

        return "\n".join(context_lines)

    except Exception as e:
        logger.exception("Errore nel recupero del contesto per il messaggio #%s.", email_number)
        return f"Errore durante il recupero del contesto del messaggio: {str(e)}"

@mcp.tool()
def reply_to_email_by_number(email_number: int, reply_text: str) -> str:
    """
    Reply to a specific email by its number from the last listing
    
    Args:
        email_number: The number of the email from the list results
        reply_text: The text content for the reply
        
    Returns:
        Status message indicating success or failure
    """
    try:
        if not email_cache:
            logger.warning("reply_to_email_by_number chiamato ma la cache e vuota.")
            return "Errore: nessun elenco disponibile. Usa prima list_recent_emails o search_emails."
        
        if email_number not in email_cache:
            logger.warning("Messaggio numero %s non presente in cache per reply_to_email_by_number.", email_number)
            return f"Errore: il messaggio #{email_number} non e presente nell'elenco corrente."
        
        email_id = email_cache[email_number]["id"]
        logger.info("Preparazione risposta per il messaggio #%s.", email_number)
        
        # Connect to Outlook
        outlook, namespace = connect_to_outlook()
        
        # Retrieve the specific email
        email = namespace.GetItemFromID(email_id)
        if not email:
            return f"Errore: il messaggio #{email_number} non puo essere recuperato da Outlook."
        
        # Create reply
        reply = email.Reply()
        reply.Body = reply_text
        
        # Send the reply
        reply.Send()
        
        logger.info("Risposta al messaggio #%s inviata correttamente.", email_number)
        return f"Risposta inviata a: {email.SenderName} <{email.SenderEmailAddress}>"
    
    except Exception as e:
        logger.exception("Errore durante la risposta al messaggio #%s.", email_number)
        return f"Errore durante l'invio della risposta: {str(e)}"

@mcp.tool()
def compose_email(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> str:
    """
    Compose and send a new email
    
    Args:
        recipient_email: Email address of the recipient
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Email address for CC (optional)
        
    Returns:
        Status message indicating success or failure
    """
    try:
        logger.info(
            "compose_email chiamato per destinatario=%s oggetto='%s' cc=%s",
            recipient_email,
            subject,
            cc_email,
        )
        # Connect to Outlook
        outlook, _ = connect_to_outlook()
        
        # Create a new email
        mail = outlook.CreateItem(0)  # 0 is the value for a mail item
        mail.Subject = subject
        mail.To = recipient_email
        
        if cc_email:
            mail.CC = cc_email
        
        # Add signature to the body
        mail.Body = body
        
        # Send the email
        mail.Send()
        
        logger.info("Email inviata correttamente a %s con oggetto '%s'.", recipient_email, subject)
        return f"Email inviata a: {recipient_email}"
    
    except Exception as e:
        logger.exception("Errore durante l'invio dell'email a %s con oggetto '%s'.", recipient_email, subject)
        return f"Errore durante l'invio dell'email: {str(e)}"

# Run the server
if __name__ == "__main__":
    print("Avvio di Outlook MCP Server...")
    print("Connessione a Outlook...")
    logger.info("Avvio di Outlook MCP Server.")
    
    try:
        # Test Outlook connection
        outlook, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)  # 6 is inbox
        print(f"Connessione a Outlook riuscita. La Posta in arrivo contiene {inbox.Items.Count} elementi.")
        logger.info("Connessione a Outlook riuscita. La Posta in arrivo contiene %s elementi.", inbox.Items.Count)
        
        # Run the MCP server
        print("Avvio del server MCP. Premi Ctrl+C per interrompere.")
        logger.info("Server MCP avviato. In attesa di richieste.")
        mcp.run()
    except Exception as e:
        print(f"Errore durante l'avvio del server: {str(e)}")
        logger.exception("Errore durante l'avvio di Outlook MCP Server.")
