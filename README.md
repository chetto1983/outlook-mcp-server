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
- Python packages from `requirements.txt` (core: `mcp>=1.2.0`, `pywin32>=305`; HTTP bridge extras: `fastapi>=0.110`, `uvicorn[standard]>=0.27`)
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

### Streamable HTTP transport (n8n, Docker)

The MCP specification defines a **streamable HTTP** transport that n8n supports out of the box. Run the server with this transport so the n8n MCP Client node can list and invoke tools directly.

1. Install the optional runtime dependency if you have not already: `pip install uvicorn[standard]`.
2. Start the server from the Windows host that can reach Outlook:

   ```bash
   python outlook_mcp_server.py --mode mcp --transport streamable-http --host 0.0.0.0 --port 8000 --stream-path /mcp
   ```

   The server checks the Outlook COM bridge, binds to `http://0.0.0.0:8000`, and exposes a streamable MCP endpoint at `/mcp`.

3. In n8n (Docker) configure the **MCP Client Tool** node:
   - `Endpoint`: `http://host.docker.internal:8000/mcp` (adjust if Docker cannot resolve `host.docker.internal`)
   - `Transport`: **HTTP Streamable**
   - Authentication: set according to your network needs (defaults to “None”)
   - Select the Outlook tools you want to expose to your workflow

If you prefer Server-Sent Events, start the server with `--transport sse --sse-path /sse --mount-path /` and point the n8n node’s SSE endpoint to `http://host.docker.internal:8000/sse`.

### REST bridge mode (optional)

For simple automations that only need to trigger specific tools over plain REST (without implementing the MCP protocol), you can launch the lightweight FastAPI shim:

1. Install the optional dependencies: `pip install fastapi uvicorn[standard]`.
2. Start the bridge:

   ```bash
   python outlook_mcp_server.py --mode http --host 0.0.0.0 --port 8000
   ```

3. Call the HTTP endpoints from your automation platform:
   - `POST http://host.docker.internal:8000/tools/list_recent_emails` with body `{"arguments": {...}}`
   - Alternatively `POST http://host.docker.internal:8000/` with `{"tool": "list_recent_emails", "arguments": {...}}`

Available endpoints in REST mode:

- `GET /health` – readiness probe
- `GET /tools` – list tool metadata (names, schemas, descriptions)
- `GET /` – quick-start message plus the currently registered tool names
- `POST /tools/{tool_name}` – execute any MCP tool; supply arguments as JSON body
- `POST /` – alternative execution form using a body like `{"tool": "list_recent_emails", "arguments": {...}}` (also accepts `{"list_recent_emails": {...}}`)

> **Note:** Regardless of the transport, the server must run on the Windows host where Outlook is available. Docker containers cannot access the COM automation layer directly.

### Tool reference

Each MCP tool accepts keyword arguments so clients can override defaults as needed.

- `params()` - Returns general metadata and HTTP endpoint pointers for automation handshakes (used by some n8n MCP nodes).
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

## Outlook automation roadmap

The next Outlook-first enhancements (all executable directly through the MCP server) are grouped into five pillars:

1. **Dynamic domain folders**
   - Rilevare il dominio del mittente quando arriva un nuovo messaggio (via hook COM o strumenti MCP) e creare automaticamente la gerarchia `Clienti/<dominio>/...`.
   - Spostare subito il messaggio nella sottocartella adeguata, riutilizzando la struttura per tutte le email successive dello stesso dominio.
2. **Priorità e promemoria interni**
   - Applicare categorie/flag Outlook per messaggi “Azione/Critico”, così n8n o altri flussi possono reagire senza scansioni pesanti.
   - Mantenere un log leggero delle azioni critiche per offrirle via MCP senza riesaminare tutte le cartelle.
3. **Monitoraggio meeting e inviti**
   - Agganciarsi agli eventi calendario per intercettare aggiornamenti/cancellazioni e leggere `responseStatus` degli invitati.
   - Esporre strumenti MCP dedicati per inviare promemoria o riconfermare la partecipazione direttamente da Outlook.
4. **Mailbox condivise e nuovo Outlook**
   - Estendere regole e categorie a mailbox condivise già montate; valutare limiti del “New Outlook” e usare flag/stati quando le categorie non sono disponibili.
5. **Bozze intelligenti e pianificazione**
   - Aggiungere tool MCP per depositare bozze in Draft a partire dal contesto email (senza API esterne).
   - Creare appuntamenti follow-up con COM `Appointments.Add`, così n8n deve solo orchestrare notifiche esterne.

Ogni passo resta compatibile con n8n: il server MCP gestisce Outlook, mentre n8n si occupa delle automazioni cross-canale (Telegram, CRM, task manager).

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
