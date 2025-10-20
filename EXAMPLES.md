# Esempi d'Uso - Outlook MCP Server

Questa guida contiene esempi pratici per i casi d'uso più comuni.

---

## Indice
- [Setup Iniziale](#setup-iniziale)
- [Gestione Email](#gestione-email)
- [Calendario](#calendario)
- [Contatti](#contatti)
- [Automazioni Avanzate](#automazioni-avanzate)
- [Integrazione con n8n](#integrazione-con-n8n)

---

## Setup Iniziale

### Esempio 1: Verifica Connessione e Feature Status

**Scenario:** Verificare che il server sia operativo e vedere quali tool sono abilitati.

**Comandi MCP:**
```python
# 1. Verifica parametri server
params()

# 2. Ottieni data/ora corrente
get_current_datetime(include_utc=True)

# 3. Verifica feature abilitate
feature_status()
```

**Risposta attesa:**
```json
{
  "enabled_groups": ["system", "email.list", "email.detail", ...],
  "disabled_groups": [],
  "enabled_tools": 45,
  "disabled_tools": 0
}
```

---

### Esempio 2: Configurazione Read-Only

**Scenario:** Configurare il server per permettere solo lettura (nessun invio email, nessuna modifica).

**File `features.json`:**
```json
{
  "disabled_groups": [
    "email.actions",
    "calendar.write",
    "batch",
    "domain.rules"
  ]
}
```

**Reload configurazione senza riavvio:**
```python
reload_configuration()
```

---

## Gestione Email

### Esempio 3: Briefing Mattutino

**Scenario:** Ogni mattina, ottenere un riepilogo delle email urgenti e degli appuntamenti della giornata.

**Workflow:**
```python
# 1. Ottieni data corrente
current_time = get_current_datetime(include_utc=True)

# 2. Email non lette delle ultime 24 ore
unread = list_recent_emails(
    days=1,
    unread_only=True,
    max_results=10,
    include_preview=True
)

# 3. Email che richiedono risposta (ultimi 7 giorni)
pending = list_pending_replies(
    days=7,
    max_results=5,
    unread_only=False
)

# 4. Eventi di oggi
today_events = list_upcoming_events(
    days=1,
    include_description=False,
    include_all_calendars=True
)
```

**Risposta strutturata:**
```
🕐 Ore 08:30 - 20 Ottobre 2025

📧 EMAIL NON LETTE (24h): 3 messaggi
  #1 | Mario Rossi | Preventivo Q4 | 07:45 | 📎 2 allegati
  #2 | Anna Bianchi | Urgente: Riunione posticipata | 08:15
  #3 | Supporto IT | Manutenzione programmata | 08:20

⚠️  RICHIESTE DI RISPOSTA: 2 messaggi
  #1 | Cliente Acme | Conferma ordine | 3 giorni fa
  #2 | Fornitore XYZ | Disponibilità materiale | 5 giorni fa

📅 EVENTI OGGI: 2 appuntamenti
  #1 | 10:00-11:00 | Riunione Team | Sala Conferenze
  #2 | 15:30-16:30 | Call Cliente | Online (Teams)
```

---

### Esempio 4: Ricerca Email con Filtri Avanzati

**Scenario:** Trovare tutte le email di un cliente specifico negli ultimi 30 giorni, in tutte le cartelle.

**Comandi MCP:**
```python
# Ricerca con termini multipli (OR)
results = search_emails(
    search_term="Acme Corp OR acme@example.com",
    days=30,
    include_all_folders=True,
    max_results=50,
    include_preview=True
)

# Ottieni dettagli del terzo risultato
email_detail = get_email_by_number(
    email_number=3,
    include_body=True
)

# Ottieni contesto conversazione
thread_context = get_email_context(
    email_number=3,
    include_thread=True,
    thread_limit=15,
    lookback_days=180
)
```

---

### Esempio 5: Rispondere a Email con Validazione Contatti

**Scenario:** Rispondere a una email verificando prima i contatti del destinatario.

**Workflow Sicuro:**
```python
# 1. Lista email recenti
list_recent_emails(days=7, max_results=10)

# 2. Ottieni dettagli email
email = get_email_by_number(email_number=2, include_body=True)

# 3. VALIDAZIONE CONTATTI OBBLIGATORIA
contacts = search_contacts(search_term="Mario Rossi")

# 4. Se contatti non trovati, cerca nella corrispondenza
if not contacts:
    # Estrai da thread
    context = get_email_context(email_number=2, include_thread=True)

    # Cerca email passate
    past_emails = search_emails(
        search_term="Mario Rossi",
        days=30,
        include_all_folders=True
    )

    # Se ancora dubbi → CHIEDI CONFERMA UTENTE
    # Esempio: "Confermi l'indirizzo mario.rossi@acme.com?"

# 5. SOLO DOPO VALIDAZIONE → Invia risposta
reply = reply_to_email_by_number(
    email_number=2,
    reply_text="Buongiorno Mario,\n\nGrazie per il messaggio...",
    reply_all=False,
    send=False  # False = crea bozza per revisione manuale
)

# 6. Se bozza OK → Invia manualmente o imposta send=True
```

---

### Esempio 6: Composizione Email con Allegati

**Scenario:** Creare e inviare una nuova email con allegati.

**Comandi MCP:**
```python
# 1. Valida destinatari
recipient_check = search_contacts(search_term="Anna Bianchi")
cc_check = search_contacts(search_term="Mario Rossi")

# 2. Componi email con allegati
compose_email(
    recipient_email="anna.bianchi@example.com",
    subject="Documenti Q4 2025",
    body="Buongiorno Anna,\n\nAllego i documenti richiesti.\n\nCordiali saluti,\nDavide",
    cc_email="mario.rossi@example.com",
    attachments=[
        "C:\\Documenti\\Report_Q4.pdf",
        "C:\\Documenti\\Budget_2025.xlsx"
    ],
    send=False  # Crea bozza per revisione
)
```

---

### Esempio 7: Organizzazione Automatica per Dominio

**Scenario:** Archiviare automaticamente email in cartelle per dominio mittente.

**Comandi MCP:**
```python
# 1. Lista email recenti
emails = list_recent_emails(days=1, max_results=20)

# 2. Per ogni email, crea struttura dominio e sposta
for i in range(1, 6):  # Prime 5 email
    # Crea cartella se non esiste
    ensure_domain_folder(
        email_number=i,
        root_folder_name="Clienti",
        subfolders=["Inbox", "Progetti"]
    )

    # Sposta email nella cartella dominio
    move_email_to_domain_folder(
        email_number=i,
        root_folder_name="Clienti",
        create_if_missing=True
    )

    # Marca come letto
    mark_email_read_unread(
        email_number=i,
        unread=False
    )
```

**Risultato:**
```
Clienti/
├── acme.com/
│   ├── Inbox/
│   └── Progetti/
├── example.com/
│   ├── Inbox/
│   └── Progetti/
```

---

### Esempio 8: Gestione Batch di Email

**Scenario:** Spostare tutte le newsletter in una cartella dedicata e marcarle come lette.

**Comandi MCP:**
```python
# 1. Cerca newsletter
newsletters = search_emails(
    search_term="newsletter OR unsubscribe",
    days=7,
    include_all_folders=False,
    max_results=50
)

# 2. Crea cartella Newsletter se non esiste
create_folder(
    new_folder_name="Newsletter",
    parent_folder_name="Inbox",
    allow_existing=True
)

# 3. Batch move + mark read
batch_manage_emails(
    email_numbers=[1, 2, 3, 4, 5],  # IDs dalle ricerca
    move_to_folder_name="Newsletter",
    mark_as="read"
)
```

---

## Calendario

### Esempio 9: Lista Eventi della Settimana

**Scenario:** Ottenere tutti gli eventi dei prossimi 7 giorni da tutti i calendari.

**Comandi MCP:**
```python
events = list_upcoming_events(
    days=7,
    max_results=50,
    include_description=True,
    include_all_calendars=True
)
```

**Risposta formattata:**
```
📅 EVENTI PROSSIMI 7 GIORNI

Lunedì 20 Ottobre
  #1 | 09:00-10:00 | Standup Team | Sala Meeting
  #2 | 14:00-15:30 | Presentazione Cliente | Online

Martedì 21 Ottobre
  #3 | 10:00-12:00 | Workshop Sicurezza | Sede Centrale
  #4 | 15:00-16:00 | 1-on-1 Manager | Ufficio 402

Mercoledì 22 Ottobre
  #5 | TUTTO IL GIORNO | Ferie Mario Rossi
```

---

### Esempio 10: Ricerca Eventi per Keyword

**Scenario:** Trovare tutte le riunioni che menzionano un progetto specifico.

**Comandi MCP:**
```python
# Ricerca eventi
project_events = search_calendar_events(
    search_term="Progetto Alpha",
    days=30,
    include_all_calendars=True,
    include_description=True
)

# Dettaglio evento specifico
event_detail = get_event_by_number(event_number=2)
```

---

### Esempio 11: Creazione Evento con Inviti

**Scenario:** Creare una riunione e invitare partecipanti.

**Comandi MCP:**
```python
# 1. Valida contatti partecipanti
search_contacts("Anna Bianchi")
search_contacts("Mario Rossi")

# 2. Crea evento
create_calendar_event(
    subject="Review Progetto Alpha",
    start_time="2025-10-25T10:00:00",
    duration_minutes=90,
    location="Sala Conferenze A",
    attendees=[
        "anna.bianchi@example.com",
        "mario.rossi@example.com"
    ],
    body="Agenda:\n1. Stato avanzamento\n2. Budget Q4\n3. Prossimi step",
    send_invitations=True,
    reminder_minutes_before_start=15
)
```

---

### Esempio 12: Preparazione Riunione Automatica

**Scenario:** Prima di una riunione, raccogliere automaticamente email recenti con i partecipanti.

**Workflow:**
```python
# 1. Trova prossima riunione
events = list_upcoming_events(days=1, max_results=5)

# 2. Ottieni dettagli evento
event = get_event_by_number(event_number=1)
attendees = event.get("attendees", [])

# 3. Per ogni partecipante, cerca email recenti
for attendee in attendees:
    # Cerca contatto
    contact = search_contacts(attendee)

    # Cerca email ultimi 30 giorni
    emails = search_emails(
        search_term=attendee,
        days=30,
        include_all_folders=True,
        max_results=10
    )

    # Output: Lista email rilevanti per partecipante
```

---

## Contatti

### Esempio 13: Ricerca Contatti con Dettagli

**Scenario:** Cercare un contatto e ottenere tutti i suoi recapiti.

**Comandi MCP:**
```python
contacts = search_contacts(
    search_term="Mario",
    max_results=10
)
```

**Risposta esempio:**
```json
{
  "contacts": [
    {
      "name": "Mario Rossi",
      "email": "mario.rossi@acme.com",
      "company": "Acme Corp",
      "job_title": "Sales Manager",
      "phone": "+39 02 1234567",
      "mobile": "+39 333 1234567"
    }
  ]
}
```

---

## Automazioni Avanzate

### Esempio 14: Script Python - Auto-Risposta Fuori Ufficio

**Scenario:** Script automatico che rileva email urgenti e crea bozze di risposta mentre sei fuori ufficio.

**File: `auto_ooo_responses.py`**
```python
import requests
import json

SERVER_URL = "http://localhost:8000"

def check_urgent_emails():
    # Lista email non lette
    response = requests.post(
        f"{SERVER_URL}/tools/list_recent_emails",
        json={"arguments": {
            "days": 1,
            "unread_only": True,
            "max_results": 20
        }}
    )

    emails = response.json().get("emails", [])

    for email in emails:
        # Rileva urgenti (oggetto con URGENTE, subject con !)
        if "URGENTE" in email["subject"].upper() or "!" in email["subject"]:
            create_ooo_draft(email["number"], email["from"])

def create_ooo_draft(email_number, sender):
    # Valida contatto
    requests.post(
        f"{SERVER_URL}/tools/search_contacts",
        json={"arguments": {"search_term": sender}}
    )

    # Crea bozza risposta
    ooo_message = f"""Buongiorno,

Grazie per il messaggio. Sono attualmente fuori ufficio fino al 25/10/2025.

Per urgenze, contattare:
- Collega: anna.bianchi@example.com
- Tel: +39 02 1234567

Risponderò al rientro.

Cordiali saluti,
Davide Marchetto"""

    response = requests.post(
        f"{SERVER_URL}/tools/reply_to_email_by_number",
        json={"arguments": {
            "email_number": email_number,
            "reply_text": ooo_message,
            "reply_all": False,
            "send": False  # Crea bozza
        }}
    )

    print(f"✓ Bozza creata per email #{email_number}")

if __name__ == "__main__":
    check_urgent_emails()
```

**Esecuzione via Task Scheduler:**
- Frequenza: Ogni 2 ore durante OOO
- Comando: `python auto_ooo_responses.py`

---

### Esempio 15: Monitoraggio SLA Risposte

**Scenario:** Report automatico di email che richiedono risposta da più di X giorni (violazione SLA).

**File: `sla_monitor.py`**
```python
import requests
from datetime import datetime, timedelta

SERVER_URL = "http://localhost:8000"
SLA_DAYS = 3  # SLA: rispondere entro 3 giorni

def check_sla_violations():
    # Lista email senza risposta (ultimi 14 giorni)
    response = requests.post(
        f"{SERVER_URL}/tools/list_pending_replies",
        json={"arguments": {
            "days": 14,
            "max_results": 100
        }}
    )

    pending = response.json().get("emails", [])
    violations = []

    for email in pending:
        # Calcola giorni da ricezione
        received = datetime.fromisoformat(email["received_time"])
        days_ago = (datetime.now() - received).days

        if days_ago > SLA_DAYS:
            violations.append({
                "number": email["number"],
                "from": email["from"],
                "subject": email["subject"],
                "days_ago": days_ago,
                "sla_breach": days_ago - SLA_DAYS
            })

    # Report
    if violations:
        print(f"⚠️  VIOLAZIONI SLA: {len(violations)} email")
        print("-" * 80)
        for v in sorted(violations, key=lambda x: x["days_ago"], reverse=True):
            print(f"[{v['days_ago']} giorni] {v['from']} - {v['subject']}")
            print(f"  Superamento SLA: +{v['sla_breach']} giorni\n")
    else:
        print("✓ Nessuna violazione SLA")

if __name__ == "__main__":
    check_sla_violations()
```

---

### Esempio 16: Backup Metadata Email

**Scenario:** Esportare metadata di tutte le email degli ultimi 30 giorni in JSON per backup/analisi.

**File: `backup_emails.py`**
```python
import requests
import json
from datetime import datetime

SERVER_URL = "http://localhost:8000"

def backup_email_metadata():
    all_emails = []

    # Lista tutte email (usa pagination con offset)
    offset = 0
    batch_size = 100

    while True:
        response = requests.post(
            f"{SERVER_URL}/tools/list_recent_emails",
            json={"arguments": {
                "days": 30,
                "max_results": batch_size,
                "offset": offset,
                "include_all_folders": True,
                "include_preview": True
            }}
        )

        batch = response.json().get("emails", [])
        if not batch:
            break

        all_emails.extend(batch)
        offset += batch_size
        print(f"Processati {len(all_emails)} email...")

    # Salva backup
    backup_file = f"email_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(backup_file, "w", encoding="utf-8") as f:
        json.dump({
            "backup_date": datetime.now().isoformat(),
            "total_emails": len(all_emails),
            "emails": all_emails
        }, f, indent=2, ensure_ascii=False)

    print(f"✓ Backup completato: {backup_file} ({len(all_emails)} email)")

if __name__ == "__main__":
    backup_email_metadata()
```

---

## Integrazione con n8n

### Esempio 17: n8n Workflow - Notifica Email Urgenti su Telegram

**Scenario:** Workflow n8n che controlla ogni 15 minuti le email urgenti e invia notifica Telegram.

**Configurazione n8n:**

1. **Nodo Cron:** Trigger ogni 15 minuti

2. **Nodo HTTP Request (MCP):**
   - Method: POST
   - URL: `http://localhost:8000/tools/list_recent_emails`
   - Body (JSON):
   ```json
   {
     "arguments": {
       "days": 1,
       "unread_only": true,
       "max_results": 10
     }
   }
   ```

3. **Nodo Function:** Filtra email urgenti
   ```javascript
   const emails = $input.item.json.emails || [];
   const urgent = emails.filter(e =>
     e.subject.includes("URGENTE") ||
     e.subject.includes("URGENT") ||
     e.subject.includes("!")
   );

   return urgent.map(e => ({
     json: {
       from: e.from,
       subject: e.subject,
       preview: e.body_preview,
       received: e.received_time
     }
   }));
   ```

4. **Nodo Telegram:** Invia messaggio
   - Text:
   ```
   🚨 Email Urgente

   Da: {{$json.from}}
   Oggetto: {{$json.subject}}

   {{$json.preview}}
   ```

---

### Esempio 18: n8n Workflow - Auto-Archiviazione Newsletter

**Scenario:** Workflow che identifica newsletter e le sposta automaticamente.

**Nodi n8n:**

1. **Schedule Trigger:** Daily 22:00

2. **HTTP Request:**
   - URL: `http://localhost:8000/tools/search_emails`
   - Body:
   ```json
   {
     "arguments": {
       "search_term": "newsletter OR unsubscribe OR marketing",
       "days": 1,
       "max_results": 50
     }
   }
   ```

3. **Function:** Estrai numeri email
   ```javascript
   const emails = $input.item.json.emails || [];
   return [{
     json: {
       email_numbers: emails.map(e => e.number)
     }
   }];
   ```

4. **HTTP Request:** Batch move
   - URL: `http://localhost:8000/tools/batch_manage_emails`
   - Body:
   ```json
   {
     "arguments": {
       "email_numbers": "{{$json.email_numbers}}",
       "move_to_folder_name": "Newsletter",
       "mark_as": "read"
     }
   }
   ```

---

## Best Practices

### ✅ Raccomandazioni

1. **Cache Management:**
   - Esegui sempre `list_*` o `search_*` prima di `get_*_by_number`
   - La cache è valida per la sessione corrente

2. **Validazione Contatti:**
   - USA SEMPRE `search_contacts()` prima di `compose_email()` o `reply_to_email_by_number()`
   - Estrai indirizzi da conversazioni passate se contatto non in rubrica
   - Chiedi conferma utente per indirizzi non validati

3. **Performance:**
   - Usa `offset` per paginazione con dataset grandi
   - Limita `max_results` a valori ragionevoli (≤100)
   - Preferisci `include_all_folders=False` quando possibile

4. **Safety:**
   - Usa `send=False` per creare bozze e revisione manuale
   - Testa workflow con account non-produzione prima
   - Monitora i log regolarmente

5. **Error Handling:**
   - Implementa retry logic per errori COM transitori
   - Valida esistenza cartelle prima di `move_email_to_folder()`
   - Cattura eccezioni e logga per troubleshooting

### ❌ Da Evitare

- ❌ Chiamare `get_email_by_number()` senza prima eseguire `list_recent_emails()`
- ❌ Usare `send=True` senza validazione contatti
- ❌ Batch troppo grandi (>200 item) senza chunking
- ❌ Ignorare errori COM (possono indicare instabilità Outlook)
- ❌ Eseguire operazioni critiche senza logging

---

## Conclusione

Questi esempi coprono i casi d'uso principali. Per ulteriori informazioni:
- **FAQ.md**: Domande frequenti e troubleshooting
- **README.md**: Documentazione completa
- **prompt.txt**: Prompt ottimizzato per assistenti AI

Buon lavoro con Outlook MCP Server! 🚀
