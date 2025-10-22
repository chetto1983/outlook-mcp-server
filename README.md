# Outlook MCP Server
Outlook MCP Server espone email e calendario di Microsoft Outlook tramite il Model Context Protocol (MCP). Permette ad assistenti MCP di elencare e cercare messaggi, creare riassunti, individuare risposte mancanti, gestire allegati, consultare eventi e rispondere senza uscire da Outlook. Include anche un bridge HTTP opzionale per piattaforme di automazione (es. n8n).

Caratteristiche
- Copertura unificata di posta e calendari (Inbox, Posta inviata, cartelle condivise gia' montate nel profilo).
- Elenchi ricchi con mittente, stato lettura, categorie, ID conversazione, anteprima corpo e nomi allegati.
- Consapevolezza della conversazione: controllo delle risposte gia' inviate e outline compatta del thread.
- Calendario con occorrenze ricorrenti, ricerca per parole chiave e dettaglio evento.
- Azioni: risposta inline (`reply_to_email_by_number`), composizione (`compose_email`), spostamenti, letti/non letti, categorie, allegati, batch.
- Amministrazione cartelle: `list_folders`, `get_folder_metadata`, `create_folder`, `rename_folder`, `delete_folder`.
- Rotating logging (`logs/outlook_mcp_server.log`) e cache per sessioni lunghe.

Requisiti
- Windows con Microsoft Outlook installato e configurato (profilo aperto/accessibile)
- Python 3.10+
- Dipendenze: `pip install -r requirements.txt` (core: `mcp`, `pywin32`; opzionali HTTP: `fastapi`, `uvicorn[standard]`)
- Un client MCP compatibile (Claude Desktop o altro host MCP)

Installazione
1. Clona o scarica questa repository.
2. (Consigliato) Crea un virtualenv e attivalo.
3. Installa le dipendenze: `pip install -r requirements.txt`.
4. Assicurati che Outlook sia utilizzabile dall'utente che avvia il server.

Quick Start - 3 minuti
----------------------
```bash
# 1. Installa dipendenze core
pip install mcp pywin32

# 2. Avvia server MCP (stdio - per Claude Desktop)
python outlook_mcp_server.py

# 3. (Opzionale) Testa con modalità HTTP
pip install fastapi uvicorn[standard]
python outlook_mcp_server.py --mode http --port 8000
curl http://localhost:8000/tools
```

**🚀 Automazioni con n8n:** Vedi [N8N_QUICKSTART.md](N8N_QUICKSTART.md) per setup rapido automazioni email con n8n (workflow pronti all'uso inclusi!)

Avvio rapido (stdio MCP)
```bash
python outlook_mcp_server.py
```
All'avvio il server verifica la connessione COM verso Outlook e poi accetta richieste FastMCP su stdio.

Trasporti MCP per automazioni
- Streamable HTTP (consigliato per n8n):
  ```bash
  pip install uvicorn[standard]
  python outlook_mcp_server.py --mode mcp --transport streamable-http --host 0.0.0.0 --port 8000 --stream-path /mcp
  ```
  Endpoints MCP: `http://HOST:8000/mcp`.

- Server‑Sent Events (SSE):
  ```bash
  python outlook_mcp_server.py --mode mcp --transport sse --host 0.0.0.0 --port 8000 --sse-path /sse --mount-path /
  ```
  Endpoint SSE: `http://HOST:8000/sse`.

Bridge REST (opzionale)
Per chiamare i tool MCP via REST semplice (senza implementare MCP):
```bash
pip install fastapi uvicorn[standard]
python outlook_mcp_server.py --mode http --host 0.0.0.0 --port 8000
```
Endpoint utili:
- `GET /health` – probe di readiness
- `GET /tools` – lista tool (nome, schema, descrizione)
- `GET /` – messaggio di benvenuto + tool disponibili
- `POST /tools/{tool_name}` – esegue un tool con body `{ "arguments": { ... } }`
- `POST /` – alternativa con `{ "tool": "list_recent_emails", "arguments": { ... } }`

Nota: il server deve girare sullo stesso host Windows che ha accesso a Outlook (COM). Docker non puo' accedere a COM diretto.

Configurazione: Filtri Contenuti
- Il file `config.json` include la sezione `filters.promotional_keywords`, usata dai briefing per ignorare newsletter, promozioni e inviti marketing.
- Personalizza la lista inserendo parole chiave coerenti con le tue campagne ricorrenti, ad esempio:
  ```json
  "filters": {
    "promotional_keywords": [
      "newsletter",
      "promo vip",
      "iscriviti",
      "unsubscribe"
    ]
  }
  ```
- Dopo aver modificato il file esegui il tool `reload_configuration()` (gruppo `system`) per ricaricare i filtri senza riavviare il server.

Configurazione: Feature Flags
Abilita/disabilita gruppi o singoli tool senza modificare il codice. Di default tutto e' abilitato.
- File `features.json` nella root del progetto o variabile `OUTLOOK_MCP_FEATURES_FILE` con path ad un JSON.
- Variabili d'ambiente (separate da virgola o punto e virgola):
  - `OUTLOOK_MCP_ENABLED_GROUPS`, `OUTLOOK_MCP_DISABLED_GROUPS`
  - `OUTLOOK_MCP_ENABLED_TOOLS`, `OUTLOOK_MCP_DISABLED_TOOLS`

Esempio di `features.json`:
```json
{
  "enabled_groups": [],
  "disabled_groups": [],
  "enabled_tools": [],
  "disabled_tools": []
}
```
Gruppi rilevanti: `system`, `general`, `folders`, `email.list`, `email.detail`, `email.actions`, `attachments`, `contacts`, `calendar.read`, `calendar.write`, `domain.rules`, `batch`.

**Esempi pratici:**

Disabilita composizione e invio email (solo lettura):
```json
{
  "disabled_tools": ["compose_email", "reply_to_email_by_number"]
}
```

Modalità solo lettura completa (disabilita tutte le operazioni di scrittura):
```json
{
  "disabled_groups": ["email.actions", "calendar.write", "batch"]
}
```

Abilita solo funzionalità sistema e ricerca:
```json
{
  "enabled_groups": ["system", "email.list", "calendar.read", "contacts"],
  "disabled_groups": []
}
```

Tool di amministrazione runtime disponibili nel gruppo `system`:
- `reload_configuration()` �?" ricarica `features.json` e le variabili d'ambiente senza riavviare il server.
- `feature_status()` �?" riepiloga gruppi/tool attivi e disabilitati.

Prompt (facoltativo)
Il file `prompt.txt` contiene un primer in italiano con regole e workflow consigliato da incollare nel tuo client MCP.

Uso dei tool principali
- `params()` – metadati generali e hint per trasporti HTTP.
- `get_current_datetime(include_utc=True)` – orario locale/UTC.
- `list_folders(...)` – navigazione gerarchia cartelle con contatori/ID/path.
- `list_recent_emails(...)` / `list_sent_emails(...)` – elenchi di posta con anteprima allegati/corpo.
- `search_emails("term OR altro", ...)` – ricerca per parole chiave con supporto OR.
- `list_pending_replies(...)` – messaggi senza risposta incrociando Posta inviata e metadati conversazione.
- `get_email_by_number(...)` / `get_email_context(...)` – dettaglio e outline conversazione dalla cache corrente.
- `get_attachments(...)` / `attach_to_email(...)` – ispezione/allegato file.
- `reply_to_email_by_number(...)` / `compose_email(...)` – risposte e nuove email (plain‑text) con invio opzionale.
- `move_email_to_folder(...)`, `mark_email_read_unread(...)`, `apply_category(...)`, `batch_manage_emails(...)` – manutenzione messaggi.
- `list_upcoming_events(...)` / `search_calendar_events(...)` / `get_event_by_number(...)` – calendario. Per default le ricerche eventi scandiscono tutti i calendari visibili; usa `calendar_name` per limitarle.
- `create_calendar_event(...)` – creazione eventi (all-day o a durata) con invito opzionale.
- `move_calendar_event(...)` – riprogramma o sposta eventi esistenti (orario, durata, luogo, calendario) con aggiornamenti facoltativi ai partecipanti.

Per usare `move_calendar_event` recupera prima l'evento con `list_upcoming_events`/`search_calendar_events` e passa il relativo `event_number` (oppure l'`entry_id` se già noto). Puoi impostare:
- `new_start_time`: data/ora locale in formato `YYYY-MM-DD HH:MM` o ISO.
- `new_duration_minutes`: durata in minuti (solo eventi non all-day).
- `new_location`: testo libero per spostare la sede.
- `new_calendar_name`: sposta l'evento in un altro calendario accessibile.
- `send_updates`: `true` per inoltrare gli aggiornamenti ai partecipanti quando l'evento è un meeting.

Suggerimenti di workflow
- Esegui sempre prima un elenco/ricerca: alimenta le cache usate dai tool di dettaglio/azione.
- Usa `include_all_folders=True` (posta) o `include_all_calendars=True` (eventi) quando non conosci la collocazione.
- `list_pending_replies` estende automaticamente il lookback di conversazione (fino a 180 giorni) per garantire accuratezza.
- Le anteprime email sono tagliate a ~220 caratteri; i nomi allegati max 5.
- Le scansioni calendario includono ricorrenze e limitano a ~500 elementi per cartella.

Logging e troubleshooting
- Log a rotazione in `logs/outlook_mcp_server.log` (5MB x 3 file).
- Se Outlook e' chiuso o chiede credenziali, aprilo e riprova.
- Gli avvisi di sicurezza `pywin32` possono comparire al primo avvio: consenti l'accesso.
- Errori di cache indicano che non hai ancora eseguito un elenco/ricerca nella sessione corrente.

Risoluzione Problemi Comuni
---------------------------

| Errore | Causa | Soluzione |
|--------|-------|-----------|
| `pywintypes.com_error` | Outlook chiuso o non accessibile | Apri Outlook manualmente e verifica che il profilo sia accessibile |
| `Cache not found for email #N` | Nessun list/search eseguito | Esegui `list_recent_emails()` o `search_emails()` prima di chiamare tool di dettaglio |
| `Permission denied` | Profilo Outlook non accessibile | Verifica credenziali Windows e permessi sul profilo Outlook |
| `Tool not found` | Tool disabilitato via features.json | Controlla `features.json` o usa `feature_status()` per verificare tool abilitati |
| `Folder not found` | Cartella specificata non esiste | Usa `list_folders()` per verificare nomi cartelle corretti, oppure `create_folder()` |
| `COM object not responding` | Outlook in stato instabile | Chiudi e riapri Outlook, poi riavvia il server MCP |
| `Max retries exceeded` | Outlook sovraccarico | Riduci `max_results` nelle chiamate o aumenta timeout nelle richieste |
| `Invalid datetime format` | Formato data non corretto | Usa formato ISO 8601: `YYYY-MM-DDTHH:MM:SS` (es. `2025-10-20T14:30:00`) |
| Server non risponde su HTTP | Porta già in uso | Cambia porta con `--port` o termina processo che usa la porta |
| `ModuleNotFoundError: fastapi` | Dipendenze HTTP non installate | Esegui `pip install fastapi uvicorn[standard]` |

Test
- Unita': `python -m pytest`
- Integrazione reale (richiede Outlook aperto): `OUTLOOK_MCP_REAL=1 python -m pytest tests/test_outlook_real_integration.py`

Struttura del codice
- `outlook_mcp_server.py`: server MCP/HTTP, CLI, helper condivisi.
- `outlook_mcp/`: pacchetto con costanti, logger, connessione, util, cache, feature flags.
- `outlook_mcp/tools/`: definizione dei tool per cartelle, email (liste/dettaglio/azioni), allegati, contatti, calendario, regole per dominio, batch.

Limitazioni
- Gli elenchi email coprono fino a 30 giorni (`MAX_DAYS`), il calendario fino a 90 (`MAX_EVENT_LOOKAHEAD_DAYS`).
- I corpi molto lunghi possono essere troncati in output.
- Le email sono inviate in plain‑text; le firme di Outlook non sono aggiunte automaticamente.
- Le interazioni MCP operano sul profilo Outlook dell'utente Windows che esegue il server.
