import datetime
import os
import win32com.client
from typing import List, Optional, Dict, Any
from mcp.server.fastmcp import FastMCP, Context

# Initialize FastMCP server
mcp = FastMCP("outlook-assistant")

# Constants
MAX_DAYS = 30
# Email cache for storing retrieved emails by number
email_cache = {}
BODY_PREVIEW_MAX_CHARS = 220
DEFAULT_MAX_RESULTS = 25
ATTACHMENT_NAME_PREVIEW_MAX = 5
CONVERSATION_ID_PREVIEW_MAX = 16

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
        display_name = recipient.Name or "Unknown"
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

def _describe_importance(value: Any) -> str:
    """Map Outlook importance levels to descriptive labels."""
    importance_map = {0: "Low", 1: "Normal", 2: "High"}
    if isinstance(value, int) and value in importance_map:
        return importance_map[value]
    return str(value) if value is not None else "Unknown"

# Helper functions
def connect_to_outlook():
    """Connect to Outlook application using COM"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as e:
        raise Exception(f"Failed to connect to Outlook: {str(e)}")

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
        raise Exception(f"Failed to access folder {folder_name}: {str(e)}")

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
            try:
                received_display = mail_item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                received_iso = mail_item.ReceivedTime.strftime("%Y-%m-%dT%H:%M:%S")
            except Exception:
                received_display = str(mail_item.ReceivedTime)
                received_iso = received_display

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
        raise Exception(f"Failed to format email: {str(e)}")

def clear_email_cache():
    """Clear the email cache"""
    global email_cache
    email_cache = {}

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
                print(f"Warning: Error processing email: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Error retrieving emails: {str(e)}")
        
    return emails_list

def get_related_conversation_emails(namespace, mail_item, max_items: int = 5, lookback_days: int = 30):
    """Collect other emails from the same conversation to build context."""
    conversation_id = getattr(mail_item, "ConversationID", None)
    if not conversation_id:
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
    default_folder_ids = [6, 5]  # Inbox, Sent Items
    for folder_id in default_folder_ids:
        try:
            folder = namespace.GetDefaultFolder(folder_id)
            potential_folders.append(folder)
        except Exception:
            continue

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
            continue

        manual_filter = False
        candidate_items = items
        try:
            filter_query = f"[ConversationID] = '{conversation_id}'"
            candidate_items = items.Restrict(filter_query)
        except Exception:
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
                continue

        if len(related_entries) >= max_items:
            break

    # Sort newest first
    related_entries.sort(
        key=lambda entry: entry[0] if entry[0] else datetime.datetime.min,
        reverse=True,
    )
    return [entry[1] for entry in related_entries]

# MCP Tools
@mcp.tool()
def list_folders() -> str:
    """
    List all available mail folders in Outlook
    
    Returns:
        A list of available mail folders
    """
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        result = "Available mail folders:\n\n"
        
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
        return f"Error listing mail folders: {str(e)}"

@mcp.tool()
def list_recent_emails(
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
) -> str:
    """
    List email titles from the specified number of days
    
    Args:
        days: Number of days to look back for emails (max 30)
        folder_name: Name of the folder to check (if not specified, checks the Inbox)
        max_results: Maximum number of emails to display (1-200)
        include_preview: Include a trimmed body preview for each email
        
    Returns:
        Numbered list of email titles with sender information
    """
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        return "Error: 'max_results' must be an integer between 1 and 200"

    include_preview = _coerce_bool(include_preview)
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the appropriate folder
        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder '{folder_name}' not found"
        else:
            folder = namespace.GetDefaultFolder(6)  # Default inbox
        
        # Clear previous cache
        clear_email_cache()
        
        # Get emails from folder
        emails = get_emails_from_folder(folder, days)
        
        # Format the output and cache emails
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails found in {folder_display} from the last {days} days."

        visible_emails = emails[:max_results]
        visible_count = len(visible_emails)
        
        if len(emails) > visible_count:
            header = (
                f"Found {len(emails)} emails in {folder_display} from the last {days} days. "
                f"Showing first {visible_count} result(s)."
            )
        else:
            header = f"Found {visible_count} emails in {folder_display} from the last {days} days."
        
        result = header + "\n\n"
        
        # Cache emails and build result
        for i, email in enumerate(visible_emails, 1):
            # Store in cache
            email_cache[i] = email
            
            folder_path = email.get("folder_path") or folder_display
            importance_label = email.get("importance_label") or _describe_importance(email.get("importance"))
            trimmed_conv = _trim_conversation_id(email.get("conversation_id"))
            attachments_line = None
            if email.get("attachment_names"):
                attachments_line = f"Attachment Names: {', '.join(email['attachment_names'])}"
            
            result += f"Email #{i}\n"
            result += f"Subject: {email.get('subject', '(No subject)')}\n"
            result += f"From: {email.get('sender', 'Unknown')} <{email.get('sender_email', '')}>\n"
            result += f"Received: {email.get('received_time', 'Unknown')}\n"
            result += f"Folder: {folder_path}\n"
            result += f"Importance: {importance_label}\n"
            result += f"Read Status: {'Read' if not email.get('unread') else 'Unread'}\n"
            result += f"Has Attachments: {'Yes' if email.get('has_attachments') else 'No'}\n"
            if attachments_line:
                result += attachments_line + "\n"
            if include_preview and email.get("preview"):
                result += f"Preview: {email['preview']}\n"
            if email.get("categories"):
                result += f"Categories: {email['categories']}\n"
            if trimmed_conv:
                result += f"Conversation ID: {trimmed_conv}\n"
            result += "\n"
        
        result += (
            "To view the full content of an email, use the get_email_by_number tool with the email number.\n"
            "To gather the wider conversation context, use the get_email_context tool."
        )
        return result
    
    except Exception as e:
        return f"Error retrieving email titles: {str(e)}"

@mcp.tool()
def search_emails(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
) -> str:
    """
    Search emails by contact name or keyword within a time period
    
    Args:
        search_term: Name or keyword to search for
        days: Number of days to look back (max 30)
        folder_name: Name of the folder to search (if not specified, searches the Inbox)
        max_results: Maximum number of emails to display (1-200)
        include_preview: Include a trimmed body preview for each email
        
    Returns:
        Numbered list of matching email titles
    """
    if not search_term:
        return "Error: Please provide a search term"
        
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        return "Error: 'max_results' must be an integer between 1 and 200"

    include_preview = _coerce_bool(include_preview)
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the appropriate folder
        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder '{folder_name}' not found"
        else:
            folder = namespace.GetDefaultFolder(6)  # Default inbox
        
        # Clear previous cache
        clear_email_cache()
        
        # Get emails matching search term
        emails = get_emails_from_folder(folder, days, search_term)
        
        # Format the output and cache emails
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails matching '{search_term}' found in {folder_display} from the last {days} days."

        visible_emails = emails[:max_results]
        visible_count = len(visible_emails)

        if len(emails) > visible_count:
            header = (
                f"Found {len(emails)} emails matching '{search_term}' in {folder_display} from the last {days} days. "
                f"Showing first {visible_count} result(s)."
            )
        else:
            header = (
                f"Found {visible_count} emails matching '{search_term}' in {folder_display} from the last {days} days."
            )
        
        result = header + "\n\n"
        
        # Cache emails and build result
        for i, email in enumerate(visible_emails, 1):
            # Store in cache
            email_cache[i] = email
            
            folder_path = email.get("folder_path") or folder_display
            importance_label = email.get("importance_label") or _describe_importance(email.get("importance"))
            trimmed_conv = _trim_conversation_id(email.get("conversation_id"))
            attachments_line = None
            if email.get("attachment_names"):
                attachments_line = f"Attachment Names: {', '.join(email['attachment_names'])}"
            
            # Format for display
            result += f"Email #{i}\n"
            result += f"Subject: {email.get('subject', '(No subject)')}\n"
            result += f"From: {email.get('sender', 'Unknown')} <{email.get('sender_email', '')}>\n"
            result += f"Received: {email.get('received_time', 'Unknown')}\n"
            result += f"Folder: {folder_path}\n"
            result += f"Importance: {importance_label}\n"
            result += f"Read Status: {'Read' if not email.get('unread') else 'Unread'}\n"
            result += f"Has Attachments: {'Yes' if email.get('has_attachments') else 'No'}\n"
            if attachments_line:
                result += attachments_line + "\n"
            if include_preview and email.get("preview"):
                result += f"Preview: {email['preview']}\n"
            if email.get("categories"):
                result += f"Categories: {email['categories']}\n"
            if trimmed_conv:
                result += f"Conversation ID: {trimmed_conv}\n"
            result += "\n"
        
        result += (
            "To view the full content of an email, use the get_email_by_number tool with the email number.\n"
            "To gather the wider conversation context, use the get_email_context tool."
        )
        return result
    
    except Exception as e:
        return f"Error searching emails: {str(e)}"

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
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."
        
        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."
        
        email_data = email_cache[email_number]
        
        # Connect to Outlook to get the full email content
        _, namespace = connect_to_outlook()
        
        # Retrieve the specific email
        email = namespace.GetItemFromID(email_data["id"])
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."
        
        trimmed_conv = _trim_conversation_id(email_data.get("conversation_id"), max_chars=32)
        importance_label = email_data.get("importance_label") or _describe_importance(email_data.get("importance"))
        attachment_names_preview = email_data.get("attachment_names") or []
        to_line = ", ".join(email_data.get("to_recipients", []))
        cc_line = ", ".join(email_data.get("cc_recipients", []))
        bcc_line = ", ".join(email_data.get("bcc_recipients", []))
        
        result_lines = [
            f"Email #{email_number} Details:",
            "",
            f"Subject: {email_data.get('subject', '(No subject)')}",
            f"From: {email_data.get('sender', 'Unknown')} <{email_data.get('sender_email', '')}>",
        ]

        if to_line:
            result_lines.append(f"To: {to_line}")
        if cc_line:
            result_lines.append(f"Cc: {cc_line}")
        if bcc_line:
            result_lines.append(f"Bcc: {bcc_line}")

        result_lines.extend(
            [
                f"Received: {email_data.get('received_time', 'Unknown')}",
                f"Folder: {email_data.get('folder_path', 'Unknown folder')}",
                f"Importance: {importance_label}",
                f"Read Status: {'Read' if not email_data.get('unread') else 'Unread'}",
            ]
        )

        if email_data.get("categories"):
            result_lines.append(f"Categories: {email_data['categories']}")
        if trimmed_conv:
            result_lines.append(f"Conversation ID: {trimmed_conv}")
        if email_data.get("preview"):
            result_lines.append(f"Body Preview: {email_data['preview']}")

        result_lines.append(f"Has Attachments: {'Yes' if email_data.get('has_attachments') else 'No'}")
        if attachment_names_preview:
            result_lines.append(f"Attachment Names: {', '.join(attachment_names_preview)}")

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
            result_lines.append("Attachments:")
            result_lines.extend(attachment_lines)
            result_lines.append("")

        result_lines.append("Body:")
        result_lines.append(email_data.get("body", "(No body content)"))
        
        result_lines.append("")
        result_lines.append(
            "To reply to this email, use the reply_to_email_by_number tool with this email number."
        )
        
        return "\n".join(result_lines)
    
    except Exception as e:
        return f"Error retrieving email details: {str(e)}"

@mcp.tool()
def get_email_context(
    email_number: int,
    include_thread: bool = True,
    thread_limit: int = 5,
    lookback_days: int = 30,
) -> str:
    """
    Provide conversation-aware context for a previously listed email.
    
    Args:
        email_number: The number of the email from the last list/search result
        include_thread: Whether to include other emails from the same conversation
        thread_limit: Maximum number of related conversation emails to include
        lookback_days: How far back to look for related messages
    
    Returns:
        Detailed context summary for the specified email
    """
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."
        
        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."

        if not isinstance(thread_limit, int) or thread_limit < 1:
            return "Error: 'thread_limit' must be a positive integer."

        if not isinstance(lookback_days, int) or lookback_days < 1 or lookback_days > 180:
            return "Error: 'lookback_days' must be an integer between 1 and 180."
        
        email_data = email_cache[email_number]
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_data["id"])
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."

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
        sender_display = f"{email_data.get('sender', 'Unknown')} <{email_data.get('sender_email', '')}>".strip()
        if sender_display:
            participants.add(sender_display)
        for recipient in email_data.get("recipients", []):
            if recipient:
                participants.add(recipient)

        context_lines = [
            f"Context for Email #{email_number}",
            "",
            f"Subject: {email_data.get('subject', 'Unknown subject')}",
            f"From: {email_data.get('sender', 'Unknown sender')} <{email_data.get('sender_email', '')}>",
        ]

        if email_data.get("to_recipients"):
            context_lines.append(f"To: {', '.join(email_data['to_recipients'])}")
        if email_data.get("cc_recipients"):
            context_lines.append(f"Cc: {', '.join(email_data['cc_recipients'])}")
        if email_data.get("bcc_recipients"):
            context_lines.append(f"Bcc: {', '.join(email_data['bcc_recipients'])}")

        context_lines.extend(
            [
                f"Received: {email_data.get('received_time', 'Unknown')}",
                f"Folder: {email_data.get('folder_path', 'Unknown')}",
                f"Importance: {importance_label}",
                f"Read Status: {'Read' if not email_data.get('unread') else 'Unread'}",
            ]
        )

        if email_data.get("categories"):
            context_lines.append(f"Categories: {email_data['categories']}")
        if email_data.get("conversation_id"):
            trimmed_conv = _trim_conversation_id(email_data["conversation_id"], max_chars=32)
            conv_line = f"Conversation ID: {trimmed_conv}" if trimmed_conv else "Conversation ID: (Unavailable)"
            if trimmed_conv and trimmed_conv.endswith("..."):
                conv_line += " (truncated)"
            context_lines.append(conv_line)

        if participants:
            context_lines.append(f"Participants Involved: {', '.join(sorted(participants))}")

        if email_data.get("preview"):
            context_lines.append(f"Body Preview: {email_data['preview']}")

        if attachment_names:
            context_lines.append(f"Attachments: {', '.join(attachment_names)}")

        # Always leave a blank line before additional sections
        body_content = email_data.get("body", "")
        if body_content and len(body_content) > 4000:
            truncated_body = body_content[:4000].rstrip() + "\n[Body truncated for brevity]"
        else:
            truncated_body = body_content or "(No body content)"

        context_lines.append("")
        context_lines.append("Current Email Body:")
        context_lines.append(truncated_body)
        context_lines.append("For the full body, use get_email_by_number with this email number.")

        if include_thread:
            context_lines.append("")
            context_lines.append("Related Conversation Messages:")
            related_emails = get_related_conversation_emails(
                namespace=namespace,
                mail_item=email,
                max_items=thread_limit,
                lookback_days=lookback_days,
            )

            if not related_emails:
                context_lines.append("- No additional conversation messages found within the specified window.")
            else:
                for idx, related in enumerate(related_emails, 1):
                    summary_header = (
                        f"{idx}. {related.get('received_time', 'Unknown time')} | "
                        f"{related.get('sender', 'Unknown sender')} | "
                        f"{related.get('folder_path', 'Unknown folder')}"
                    )
                    context_lines.append(summary_header)
                    if related.get("subject") and related["subject"] != email_data.get("subject"):
                        context_lines.append(f"   Subject: {related['subject']}")
                    if related.get("preview"):
                        context_lines.append(f"   Preview: {related['preview']}")
                    if related.get("has_attachments"):
                        context_lines.append(
                            f"   Attachments: {related.get('attachment_count', 0)} file(s) attached."
                        )
                        if related.get("attachment_names"):
                            context_lines.append(f"   Attachment Names: {', '.join(related['attachment_names'])}")

        context_lines.append("")
        context_lines.append(
            "Tip: Use reply_to_email_by_number to respond or compose_email to start a new thread."
        )

        return "\n".join(context_lines)

    except Exception as e:
        return f"Error retrieving email context: {str(e)}"

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
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."
        
        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."
        
        email_id = email_cache[email_number]["id"]
        
        # Connect to Outlook
        outlook, namespace = connect_to_outlook()
        
        # Retrieve the specific email
        email = namespace.GetItemFromID(email_id)
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."
        
        # Create reply
        reply = email.Reply()
        reply.Body = reply_text
        
        # Send the reply
        reply.Send()
        
        return f"Reply sent successfully to: {email.SenderName} <{email.SenderEmailAddress}>"
    
    except Exception as e:
        return f"Error replying to email: {str(e)}"

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
        
        return f"Email sent successfully to: {recipient_email}"
    
    except Exception as e:
        return f"Error sending email: {str(e)}"

# Run the server
if __name__ == "__main__":
    print("Starting Outlook MCP Server...")
    print("Connecting to Outlook...")
    
    try:
        # Test Outlook connection
        outlook, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)  # 6 is inbox
        print(f"Successfully connected to Outlook. Inbox has {inbox.Items.Count} items.")
        
        # Run the MCP server
        print("Starting MCP server. Press Ctrl+C to stop.")
        mcp.run()
    except Exception as e:
        print(f"Error starting server: {str(e)}")
