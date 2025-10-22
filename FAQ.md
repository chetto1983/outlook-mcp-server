# FAQ - Outlook MCP Server

## Domande Generali

### Cos'è MCP (Model Context Protocol)?
MCP è un protocollo standard sviluppato da Anthropic per permettere agli assistenti AI di interagire con sistemi esterni tramite tool standardizzati. Outlook MCP Server implementa questo protocollo per esporre le funzionalità di Outlook.

### Posso usare questo server senza Outlook installato?
No. Il server richiede Microsoft Outlook installato e configurato su Windows, poiché utilizza l'interfaccia COM di Windows per comunicare direttamente con Outlook.

### Funziona con Outlook.com (versione web)?
No. Il server funziona solo con Microsoft Outlook desktop (parte di Microsoft 365 o Office). Non supporta Outlook.com o altre webmail.

### Funziona su Mac o Linux?
No. Il server richiede Windows perché usa pywin32 per accedere alle API COM di Outlook, disponibili solo su Windows.

### Posso usarlo con più account email?
Sì. Il server accede a tutti gli account configurati nel profilo Outlook attivo dell'utente Windows che esegue il server. Le cartelle condivise già montate nel profilo sono accessibili.

---

## Sicurezza e Privacy

### Il server invia dati a servizi esterni?
No. Il server opera completamente in locale. Tutti i dati rimangono sul tuo computer e vengono esposti solo tramite le interfacce che configuri (stdio, HTTP locale).

### È sicuro usare la modalità HTTP?
La modalità HTTP è pensata per uso locale o in reti fidate. Se esponi il server su Internet:
- Usa un reverse proxy con autenticazione (nginx, Caddy)
- Considera VPN o tunnel sicuri (Tailscale, WireGuard)
- Configura firewall appropriati

### Posso limitare le operazioni disponibili?
Sì, tramite `features.json`. Puoi creare profili read-only disabilitando i gruppi di scrittura:
```json
{
  "disabled_groups": ["email.actions", "calendar.write", "batch"]
}
```

### Come proteggo l'invio accidentale di email?
Il prompt consiglia di usare `send=False` di default nelle chiamate a `reply_to_email_by_number()` e `compose_email()`. Questo crea bozze che puoi rivedere manualmente prima dell'invio.

---

## Configurazione e Setup

### Come configuro il server per Claude Desktop?
Aggiungi questa configurazione a `claude_desktop_config.json` (vedi documentazione Claude per il percorso):
```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["C:\\path\\to\\outlook_mcp_server.py"],
      "env": {}
    }
  }
}
```

### Come uso il server con n8n?
Usa la modalità `streamable-http`:
```bash
python outlook_mcp_server.py --mode mcp --transport streamable-http --port 8000
```
Configura n8n per chiamare `http://localhost:8000/mcp` con il protocollo MCP.

### Posso eseguire il server come servizio Windows?
Sì. Puoi usare NSSM (Non-Sucking Service Manager) o Task Scheduler di Windows per avviare il server automaticamente. Assicurati che:
- L'utente del servizio abbia accesso a Outlook
- Outlook sia già aperto quando il servizio parte
- Il profilo Outlook non richieda password all'avvio

### Come cambio il percorso dei log?
Modifica `outlook_mcp/logger.py` e cambia la variabile `LOG_DIR`. Alternativamente, imposta una variabile d'ambiente `OUTLOOK_MCP_LOG_DIR`.

---

## Uso dei Tool

### Perché ricevo "Cache not found for email #5"?
I tool di dettaglio (`get_email_by_number`, `get_email_context`) richiedono che tu abbia prima eseguito un tool di lista (`list_recent_emails`, `search_emails`) nella stessa sessione. Questo popola la cache con la numerazione delle email.

Esempio corretto:
```python
# 1. Prima: popola la cache
list_recent_emails(days=7, max_results=10)

# 2. Poi: accedi ai dettagli
get_email_by_number(email_number=3)
```

### Come cerco email in tutte le cartelle?
Usa il parametro `include_all_folders=True`:
```python
search_emails(
    search_term="preventivo",
    days=30,
    include_all_folders=True
)
```

### Come gestisco le ricorrenze del calendario?
Le ricorrenze sono gestite automaticamente. `list_upcoming_events()` espande gli eventi ricorrenti mostrando ogni occorrenza individuale nel periodo richiesto.

### Posso allegare file a email esistenti?
Sì, ma solo a bozze o email non ancora inviate. Usa `attach_to_email()` specificando il numero dell'email dalla cache:
```python
attach_to_email(
    email_number=5,
    attachments=["C:\\percorso\\file.pdf"],
    send=False  # False per non inviare immediatamente
)
```

### Come funziona la validazione dei contatti?
Il prompt consiglia questo workflow:
1. Usa `search_contacts(nome_destinatario)` per cercare nella rubrica
2. Se non trovi contatti, usa `get_email_context()` o `search_emails()` per estrarre indirizzi da conversazioni passate
3. Se ancora incerto, chiedi conferma all'utente prima di inviare

### Posso usare HTML nelle email?
Sì, imposta `use_html=True` in `compose_email()` o `reply_to_email_by_number()`. Di default è `False` (plain text).

### Come faccio a escludere newsletter o promozioni dal briefing?
1. Apri `config.json` e aggiorna l'elenco `filters.promotional_keywords` con le parole chiave da ignorare (es. `"newsletter"`, `"promo"`, `"unsubscribe"`).
2. Esegui il tool MCP `reload_configuration()` per ricaricare le impostazioni senza riavviare il server.
3. Ripeti il briefing: i messaggi con quelle parole chiave in oggetto/corpo verranno filtrati automaticamente.

---

## Performance e Limiti

### Quali sono i limiti temporali di ricerca?
- Email: 30 giorni (`MAX_DAYS` in `constants.py`)
- Eventi calendario: 90 giorni in avanti (`MAX_EVENT_LOOKAHEAD_DAYS`)
- Conversazioni: 180 giorni di lookback (`MAX_CONVERSATION_LOOKBACK_DAYS`)

Puoi modificare questi limiti editando `outlook_mcp/constants.py`.

### Quante email posso processare in batch?
Non c'è un limite hard-coded, ma batch molto grandi (>100 email) possono richiedere tempo. Il server elabora item uno alla volta tramite COM. Considera di:
- Dividere operazioni batch in chunk più piccoli
- Usare `offset` in `list_recent_emails()` per paginazione
- Monitorare i log per eventuali timeout

### La cache ha una scadenza?
Sì. La cache email/eventi usa TTL (Time To Live). Di default:
- Cache email: valida per la durata della sessione
- Cache eventi: valida per la durata della sessione

Riavviando il server o chiamando `reload_configuration()`, le cache vengono svuotate.

### Posso aumentare il numero di email nella cache?
Sì. Modifica `outlook_mcp/cache.py` e aumenta `maxsize` del `TimedLRUCache`. Ad esempio:
```python
email_cache = TimedLRUCache(maxsize=1000, ttl=3600)  # 1000 email, 1 ora TTL
```

---

## Errori Comuni

### "Outlook is not running or cannot be accessed"
**Causa:** Outlook non è aperto o il profilo richiede autenticazione.

**Soluzione:**
1. Apri Outlook manualmente
2. Verifica che il profilo sia accessibile (nessuna richiesta password)
3. Riavvia il server MCP

### "Permission denied when accessing folder"
**Causa:** Permessi insufficienti sulla cartella (es. cartella condivisa con accesso limitato).

**Soluzione:**
- Verifica i permessi in Outlook manualmente
- Usa un account con privilegi appropriati
- Contatta l'amministratore della mailbox condivisa

### "COM error: Exception occurred"
**Causa:** Errore generico COM, spesso dovuto a Outlook in stato instabile.

**Soluzione:**
1. Chiudi completamente Outlook (controlla Task Manager per processi OUTLOOK.EXE residui)
2. Riapri Outlook
3. Riavvia il server MCP
4. Se persiste, riavvia Windows

### "Invalid datetime format in create_calendar_event"
**Causa:** Formato data/ora non corretto.

**Soluzione:** Usa formato ISO 8601:
```python
create_calendar_event(
    subject="Riunione",
    start_time="2025-10-20T14:30:00",  # Corretto
    duration_minutes=60
)
```

### "Tool 'compose_email' not found"
**Causa:** Tool disabilitato in `features.json`.

**Soluzione:**
1. Controlla `features.json`
2. Rimuovi il tool da `disabled_tools`
3. Esegui `reload_configuration()` o riavvia il server

---

## Integrazione e Automazione

### Come integro il server in Python?
Usa un client MCP o chiama l'API HTTP direttamente:

**Modalità HTTP (più semplice):**
```python
import requests

response = requests.post(
    "http://localhost:8000/tools/list_recent_emails",
    json={"arguments": {"days": 7, "max_results": 5}}
)
emails = response.json()
```

**Modalità MCP stdio (richiede client MCP):**
```python
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

server_params = StdioServerParameters(
    command="python",
    args=["outlook_mcp_server.py"]
)

async with stdio_client(server_params) as (read, write):
    async with ClientSession(read, write) as session:
        result = await session.call_tool("list_recent_emails", {
            "days": 7,
            "max_results": 5
        })
```

### Posso schedulare operazioni automatiche?
Sì. Usa Task Scheduler di Windows o cron (WSL) per eseguire script Python che chiamano il server HTTP:

```python
# auto_check_emails.py
import requests
import datetime

response = requests.post(
    "http://localhost:8000/tools/list_pending_replies",
    json={"arguments": {"days": 7, "max_results": 10}}
)

pending = response.json()
if len(pending.get("emails", [])) > 5:
    # Invia notifica
    print(f"Attenzione: {len(pending['emails'])} email senza risposta!")
```

### Come implemento rate limiting?
Il server non ha rate limiting nativo. Per aggiungerlo:

**Via reverse proxy (nginx):**
```nginx
limit_req_zone $binary_remote_addr zone=mcp:10m rate=10r/s;

server {
    location /tools/ {
        limit_req zone=mcp burst=20;
        proxy_pass http://localhost:8000;
    }
}
```

**Via codice (modifica `outlook_mcp_server.py`):**
```python
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address

limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter

@app.post("/tools/{tool_name}")
@limiter.limit("10/minute")
async def execute_tool(...):
    ...
```

---

## Sviluppo e Contributi

### Come aggiungo un nuovo tool?
1. Crea un nuovo modulo in `outlook_mcp/tools/` (es. `tasks.py`)
2. Importa il decoratore: `from outlook_mcp import mcp_tool`
3. Definisci la funzione con il decoratore:
```python
@mcp_tool(group="tasks")
def list_tasks(folder_name: str = "Tasks", max_results: int = 10):
    """Elenca le attività di Outlook"""
    # Implementazione...
    return {"tasks": [...]}
```
4. Importa il modulo in `outlook_mcp/tools/__init__.py`
5. Il tool sarà automaticamente registrato all'avvio

### Come eseguo i test?
```bash
# Test unità (non richiedono Outlook)
python -m pytest

# Test integrazione (richiedono Outlook aperto)
set OUTLOOK_MCP_REAL=1
python -m pytest tests/test_outlook_real_integration.py

# Test specifico
python -m pytest tests/test_cache_utils.py -v
```

### Dove trovo i log per debugging?
- Percorso: `logs/outlook_mcp_server.log`
- Rotazione: 5MB per file, 3 backup
- Livello: INFO di default

Per aumentare il dettaglio, modifica `outlook_mcp/logger.py`:
```python
logger.setLevel(logging.DEBUG)  # Cambia da INFO a DEBUG
```

### Come contribuisco al progetto?
Questo è un progetto personale. Per contributi:
1. Fai un fork del repository
2. Crea un branch per le tue modifiche
3. Invia una pull request con descrizione dettagliata
4. Assicurati che i test passino

---

## Risorse Aggiuntive

### Documentazione utile
- [MCP Documentation](https://github.com/modelcontextprotocol)
- [Claude Desktop Setup](https://claude.ai/docs)
- [pywin32 Documentation](https://github.com/mhammond/pywin32)
- [Outlook COM API](https://learn.microsoft.com/en-us/office/vba/api/overview/outlook)

### Supporto
- Issues GitHub: [outlook-mcp-server/issues](https://github.com/your-repo/issues)
- Log file: `logs/outlook_mcp_server.log`
- Feature requests: Apri una issue con label "enhancement"

### Licenza
Verifica il file `LICENSE` nella repository per i termini d'uso.
