# Outlook MCP Server

Outlook MCP Server exposes Microsoft Outlook email and calendar data through Anthropic's Model Context Protocol (MCP). It enables MCP-aware assistants to browse folders, search and summarize conversations, monitor pending replies, inspect attachments, surface calendar events, and respond without leaving Outlook.

## Features

- Unified coverage for Outlook mailboxes and calendars, including Inbox, Sent Items, shared mail folders, and additional calendars that are already available in your Outlook profile.
- Rich email listings with sender metadata, read state, categories, conversation identifiers, attachment previews (up to five filenames), and optional cross-folder aggregation.
- Sent-mail and conversation awareness, including pending-reply detection that inspects your sent items, Last Verb metadata, and historical conversation windows.
- Deep thread expansion via cached list results and `get_email_context`, which can pull related messages from Sent Items and custom folders.
- Calendar exploration across one or many calendars with recurrence-aware listings, keyword search, and detailed event retrieval.
- Action helpers that let you reply inline (`reply_to_email_by_number`) or draft new outbound mail (`compose_email`) directly from MCP clients.
- Built-in rotating logging (`logs/outlook_mcp_server.log`) and caches that keep long MCP sessions observable and responsive.

## Requirements

- Windows with Microsoft Outlook installed, configured, and signed in
- Python 3.10 or newer
- Python packages from `requirements.txt` (`mcp>=1.2.0`, `pywin32>=305`)
- An MCP-compatible client (Claude Desktop or another MCP host)

## Installation

1. Clone or download this repository.
2. (Optional) Create and activate a virtual environment.
3. Install dependencies: `pip install -r requirements.txt`.
4. Ensure Outlook is open (or can be launched) under the Windows profile running the server.

## Configuration

### Claude Desktop configuration

Add (or merge) the following block into your `MCP_client_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["C:\\\\path\\\\to\\\\outlook_mcp_server.py"],
      "env": {}
    }
  }
}
```

If you rely on a virtual environment, point `command` and `args` to the corresponding `python.exe`.

### Prompt helper (optional)

`prompt.txt` ships with an Italian primer that lists every tool and recommended workflow. Paste it into your MCP client if you want inline assistance.

## Usage

### Starting the server

```bash
python outlook_mcp_server.py
```

On startup the script validates the Outlook COM bridge, writes status information to stdout, and begins serving FastMCP requests.

### Tool reference

Each MCP tool accepts keyword arguments so clients can override defaults as needed.

- `get_current_datetime(include_utc=True)` - Returns local time plus UTC details (ISO stamps included when requested).
- `list_folders()` - Enumerates top-level folders and two nested levels of subfolders for the current Outlook profile.
- `list_recent_emails(days=7, folder_name=None, max_results=25, include_preview=True, include_all_folders=False)` - Lists newest messages with sender data, read state, attachment previews, categories, and optional multi-folder scans.
- `list_sent_emails(days=7, folder_name=None, max_results=25, include_preview=True)` - Mirrors `list_recent_emails` for Sent Items (or a custom folder) and highlights recipient lists for historical context.
- `search_emails(search_term, days=7, folder_name=None, max_results=25, include_preview=True, include_all_folders=False)` - Searches for keywords or names (supports `OR` separators) within Inbox or across all mail folders.
- `list_pending_replies(days=14, folder_name=None, max_results=25, include_preview=True, include_all_folders=False, include_unread_only=False, conversation_lookback_days=None)` - Surfaces incoming mail that likely lacks an outgoing reply by checking conversation metadata, sent mail, and optional unread filters.
- `get_email_by_number(email_number)` - Retrieves the full body, attachment list, conversation ID, and metadata for any message cached by the latest listing/search result.
- `get_email_context(email_number, include_thread=True, thread_limit=5, lookback_days=30, include_sent=True, additional_folders=None)` - Expands a cached email with related conversation entries (including Sent Items), participants, and attachment summaries.
- `list_upcoming_events(days=7, calendar_name=None, max_results=25, include_description=False, include_all_calendars=False)` - Lists appointments up to 90 days ahead, with optional aggregation across every accessible calendar.
- `search_calendar_events(search_term, days=30, calendar_name=None, max_results=25, include_description=False, include_all_calendars=False)` - Keyword-searches Outlook calendars with the same aggregation and preview controls as the listing command.
- `get_event_by_number(event_number)` - Returns full appointment metadata (attendees, recurrence flags, location, description preview) for events cached by the latest listing/search command.
- `reply_to_email_by_number(email_number, reply_text)` - Prepares and sends a plain-text reply to a cached message.
- `compose_email(recipient_email, subject, body, cc_email=None)` - Sends a new plain-text message with optional CC recipients.

### Workflow tips

- Listing/search tools refresh the caches that power `get_email_by_number`, `get_email_context`, `get_event_by_number`, `reply_to_email_by_number`, and `compose_email`. Call a listing command first if a cache warning appears.
- Use `include_all_folders=True` (mail) or `include_all_calendars=True` (events) to scan shared folders when the message location is unknown.
- `list_pending_replies` automatically increases its conversation lookback (up to 180 days) so it can confirm whether a response exists in Sent Items.
- Email previews are trimmed to 220 characters, and attachment previews contain at most five filenames for legibility.
- Calendar scans include recurrences and cap out after 500 inspected appointments per folder to keep COM calls responsive.

## Examples

```text
List unread-focused pending replies for the last 10 days:
list_pending_replies(days=10, include_unread_only=True, max_results=15)
```

```text
Search all folders for project updates with OR keywords:
search_emails("project update OR rollout", include_all_folders=True)
```

```text
Expand an email thread before summarizing it:
get_email_context(email_number=3, thread_limit=8, lookback_days=60)
```

```text
Inspect upcoming meetings across calendars and drill into one:
list_upcoming_events(days=30, include_all_calendars=True, include_description=True)
get_event_by_number(event_number=2)
```

```text
Reply inline after reviewing a cached message:
reply_to_email_by_number(email_number=1, reply_text="Grazie per l'aggiornamento, ti rispondo entro domani.")
```

```text
Compose a new outbound message with CC:
compose_email("team@example.com", "Report settimanale", "Allego il riepilogo delle attivita...", cc_email="manager@example.com")
```

## Logging and troubleshooting

- Detailed logs rotate in `logs/outlook_mcp_server.log` (5 MB per file, three retained backups).
- If Outlook is closed or prompts for credentials, reopen it (or accept the prompt) before restarting the MCP server.
- Cache errors normally mean a list/search command was not run during the current session; repeat the listing and retry.
- `pywin32` security prompts may appear the first time the script automates Outlook; allow access and (optionally) whitelist the automation in the Outlook Trust Center.

## Limitations

- Email listings cap their window at 30 days (`MAX_DAYS`); calendar listings cap at 90 days (`MAX_EVENT_LOOKAHEAD_DAYS`).
- Event scans stop after 500 appointments per folder; if your calendar is very busy, narrow `days` for better coverage.
- `get_email_by_number` and `get_email_context` truncate bodies beyond ~4000 characters for readability.
- Replies and composed emails are sent as plain text; existing Outlook signatures are not injected automatically.
- Attachments are not downloaded; only their names and counts are exposed.
- MCP interactions operate against the Outlook profile of the Windows session running the server; shared mailboxes must already be visible in Outlook.
