# Assistente Automatico con n8n - Guida Completa

Questa guida mostra come creare un assistente intelligente con n8n che gestisce automaticamente Outlook utilizzando il server MCP.

---

## Indice
- [Setup Base n8n](#setup-base-n8n)
- [Assistente Email Automatico](#assistente-email-automatico)
- [Assistente AI con Claude/OpenAI](#assistente-ai-con-claudeopenai)
- [Dashboard Telegram Bot](#dashboard-telegram-bot)
- [Assistente Vocale (Webhook)](#assistente-vocale-webhook)
- [Workflow Avanzati](#workflow-avanzati)

---

## Setup Base n8n

### Prerequisiti
```bash
# Installa n8n
npm install -g n8n

# Avvia n8n
n8n start

# URL locale: http://localhost:5678
```

### Configurazione Outlook MCP Server

**Avvia il server in modalità HTTP:**
```bash
python outlook_mcp_server.py --mode http --host 0.0.0.0 --port 8000
```

**Credenziali n8n (per riuso):**
Crea credenziale "Outlook MCP" con:
- Type: HTTP Request → Generic Credential Type
- Name: `Outlook MCP Server`
- Base URL: `http://localhost:8000`

---

## Assistente Email Automatico

### 1. Assistente Briefing Mattutino

**Trigger:** Ogni giorno alle 08:00
**Azioni:** Raccoglie info e invia report via email/Telegram

#### Workflow n8n

```
[Schedule Trigger] → [Get DateTime] → [Get Unread Emails] → [Get Pending Replies]
                                                                        ↓
                           [Format Report] ← [Get Today Events] ←──────┘
                                  ↓
                           [Send via Email/Telegram]
```

#### Configurazione Nodi

**1. Schedule Trigger**
- Mode: `Cron`
- Cron Expression: `0 8 * * *` (08:00 ogni giorno)

**2. HTTP Request - Get DateTime**
- Method: POST
- URL: `http://localhost:8000/tools/get_current_datetime`
- Body:
```json
{
  "arguments": {
    "include_utc": true
  }
}
```

**3. HTTP Request - Get Unread Emails**
- Method: POST
- URL: `http://localhost:8000/tools/list_recent_emails`
- Body:
```json
{
  "arguments": {
    "days": 1,
    "unread_only": true,
    "max_results": 10,
    "include_preview": true
  }
}
```

**4. HTTP Request - Get Pending Replies**
- Method: POST
- URL: `http://localhost:8000/tools/list_pending_replies`
- Body:
```json
{
  "arguments": {
    "days": 7,
    "max_results": 5
  }
}
```

**5. HTTP Request - Get Today Events**
- Method: POST
- URL: `http://localhost:8000/tools/list_upcoming_events`
- Body:
```json
{
  "arguments": {
    "days": 1,
    "include_all_calendars": true
  }
}
```

**6. Function - Format Report**
```javascript
const datetime = $('Get DateTime').item.json;
const unread = $('Get Unread Emails').item.json.emails || [];
const pending = $('Get Pending Replies').item.json.emails || [];
const events = $('Get Today Events').item.json.events || [];

// Formatta report
let report = `📅 BRIEFING GIORNALIERO - ${datetime.local_time}\n`;
report += `${'='.repeat(50)}\n\n`;

// Email non lette
report += `📧 EMAIL NON LETTE (24h): ${unread.length}\n`;
if (unread.length > 0) {
  unread.slice(0, 5).forEach((email, i) => {
    report += `  ${i+1}. ${email.from} | ${email.subject}\n`;
    report += `     ${email.received_time} | ${email.unread ? '🔵 Non letto' : '✓ Letto'}\n`;
  });
  if (unread.length > 5) {
    report += `  ... e altri ${unread.length - 5}\n`;
  }
}
report += `\n`;

// Risposte pendenti
report += `⚠️  RICHIESTE DI RISPOSTA: ${pending.length}\n`;
if (pending.length > 0) {
  pending.forEach((email, i) => {
    const daysAgo = Math.floor((Date.now() - new Date(email.received_time)) / (1000*60*60*24));
    report += `  ${i+1}. ${email.from} | ${email.subject}\n`;
    report += `     ${daysAgo} giorni fa${daysAgo > 3 ? ' ⚠️ URGENTE' : ''}\n`;
  });
}
report += `\n`;

// Eventi oggi
report += `📅 EVENTI OGGI: ${events.length}\n`;
if (events.length > 0) {
  events.forEach((event, i) => {
    const time = event.all_day ? 'TUTTO IL GIORNO' : `${event.start_time} - ${event.end_time}`;
    report += `  ${i+1}. ${time} | ${event.subject}\n`;
    if (event.location) report += `     📍 ${event.location}\n`;
  });
}

return {
  json: {
    report: report,
    unread_count: unread.length,
    pending_count: pending.length,
    events_count: events.length,
    has_urgent: pending.some(e => {
      const daysAgo = Math.floor((Date.now() - new Date(e.received_time)) / (1000*60*60*24));
      return daysAgo > 3;
    })
  }
};
```

**7. Telegram/Email Node**
- Text: `{{$json.report}}`
- Parse Mode: `Markdown` (per formattazione)

---

### 2. Auto-Risposta Email Urgenti

**Trigger:** Ogni 30 minuti
**Logica:** Rileva email urgenti non lette e crea bozze di risposta automatica

#### Workflow

```
[Schedule] → [List Unread] → [Filter Urgent] → [Search Contacts] → [Create Draft Reply]
                                     ↓
                              [Notify Telegram]
```

#### Nodi Chiave

**Filter Urgent (Function Node):**
```javascript
const emails = $input.item.json.emails || [];

// Filtra urgenti
const urgent = emails.filter(email => {
  const subject = email.subject.toUpperCase();
  const from = email.from.toLowerCase();

  // Criteri urgenza
  return (
    subject.includes('URGENTE') ||
    subject.includes('URGENT') ||
    subject.includes('!!!') ||
    subject.includes('ASAP') ||
    from.includes('@cliente-importante.com')
  );
});

return urgent.map(email => ({
  json: {
    email_number: email.number,
    from: email.from,
    subject: email.subject,
    preview: email.body_preview
  }
}));
```

**Create Draft Reply (HTTP Request per ogni email):**
```json
{
  "arguments": {
    "email_number": "{{$json.email_number}}",
    "reply_text": "Buongiorno,\n\nHo ricevuto il suo messaggio urgente e lo sto prendendo in carico.\nRisponderò nel dettaglio entro [TEMPO].\n\nCordiali saluti,\nDavide Marchetto",
    "reply_all": false,
    "send": false
  }
}
```
- URL: `http://localhost:8000/tools/reply_to_email_by_number`

---

### 3. Organizzazione Automatica Newsletter

**Trigger:** Ogni sera alle 22:00
**Azione:** Sposta newsletter in cartella dedicata

#### Workflow

```
[Schedule 22:00] → [Search Newsletter] → [Create Folder If Missing] → [Batch Move & Mark Read]
```

#### Search Newsletter (HTTP Request):
```json
{
  "arguments": {
    "search_term": "newsletter OR unsubscribe OR marketing OR promotional",
    "days": 1,
    "max_results": 100,
    "include_all_folders": false
  }
}
```

#### Extract Email Numbers (Function):
```javascript
const emails = $input.item.json.emails || [];
return [{
  json: {
    email_numbers: emails.map(e => e.number),
    count: emails.length
  }
}];
```

#### Batch Move (HTTP Request):
```json
{
  "arguments": {
    "email_numbers": "{{$json.email_numbers}}",
    "move_to_folder_name": "Newsletter",
    "mark_as": "read"
  }
}
```
- URL: `http://localhost:8000/tools/batch_manage_emails`

---

## Assistente AI con Claude/OpenAI

### Setup: Assistente Email Intelligente con AI

Questo workflow usa Claude/OpenAI per leggere email e generare risposte contestuali.

#### Architettura

```
[Webhook/Schedule] → [Get Email] → [Get Context] → [AI: Analyze & Draft] → [Create Reply] → [Notify User]
```

#### Workflow Dettagliato

**1. Webhook Trigger**
- HTTP Method: POST
- Path: `/process-email`
- Payload: `{ "email_number": 5 }`

**2. Get Email Detail**
- URL: `http://localhost:8000/tools/get_email_by_number`
- Body:
```json
{
  "arguments": {
    "email_number": "{{$json.body.email_number}}",
    "include_body": true
  }
}
```

**3. Get Email Context**
- URL: `http://localhost:8000/tools/get_email_context`
- Body:
```json
{
  "arguments": {
    "email_number": "{{$json.body.email_number}}",
    "include_thread": true,
    "thread_limit": 10
  }
}
```

**4. Search Contact Info**
- URL: `http://localhost:8000/tools/search_contacts`
- Body:
```json
{
  "arguments": {
    "search_term": "{{$('Get Email Detail').item.json.from}}"
  }
}
```

**5. OpenAI/Claude Node - Generate Reply**

**Prompt:**
```
Sei l'assistente email di Davide Marchetto. Analizza questa email e genera una risposta professionale.

CONTESTO CONVERSAZIONE:
{{$('Get Email Context').item.json.thread_outline}}

EMAIL CORRENTE:
Da: {{$('Get Email Detail').item.json.from}}
Oggetto: {{$('Get Email Detail').item.json.subject}}
Data: {{$('Get Email Detail').item.json.received_time}}

Corpo:
{{$('Get Email Detail').item.json.body}}

INFO CONTATTO:
{{$('Search Contact Info').item.json}}

ISTRUZIONI:
1. Analizza il contenuto e il contesto della conversazione
2. Identifica la richiesta principale
3. Genera una risposta professionale, concisa e in italiano
4. Usa tono formale ma cordiale
5. Includi saluti appropriati
6. Se richieste informazioni che non hai, indica chiaramente cosa serve verificare

FORMATO OUTPUT:
{
  "analysis": "Breve analisi della richiesta",
  "reply_text": "Testo completo della risposta da inviare",
  "urgency": "low|medium|high",
  "needs_review": true|false,
  "action_items": ["lista", "azioni", "necessarie"]
}
```

**Modello:** GPT-4 o Claude 3.5 Sonnet
**Temperature:** 0.3 (per risposte coerenti)

**6. Function - Parse AI Response**
```javascript
const aiResponse = JSON.parse($input.item.json.choices[0].message.content);

return {
  json: {
    reply_text: aiResponse.reply_text,
    analysis: aiResponse.analysis,
    urgency: aiResponse.urgency,
    needs_review: aiResponse.needs_review,
    action_items: aiResponse.action_items,
    email_number: $('Webhook').item.json.body.email_number
  }
};
```

**7. Create Draft Reply (HTTP Request)**
```json
{
  "arguments": {
    "email_number": "{{$json.email_number}}",
    "reply_text": "{{$json.reply_text}}",
    "reply_all": false,
    "send": false
  }
}
```
- URL: `http://localhost:8000/tools/reply_to_email_by_number`

**8. Telegram Notification**
```
🤖 Bozza Risposta Generata

📧 Email: {{$('Get Email Detail').item.json.subject}}
👤 Da: {{$('Get Email Detail').item.json.from}}

📊 Analisi:
{{$('Parse AI Response').item.json.analysis}}

🚦 Urgenza: {{$('Parse AI Response').item.json.urgency}}
✍️  Revisione necessaria: {{$('Parse AI Response').item.json.needs_review ? 'Sì' : 'No'}}

✅ Azioni suggerite:
{{$('Parse AI Response').item.json.action_items.join('\n')}}

📝 Bozza creata in Outlook
```

---

## Dashboard Telegram Bot

### Assistente Interattivo via Telegram

Crea un bot Telegram che permette di controllare Outlook con comandi testuali.

#### Setup Telegram Bot

1. **Crea bot con @BotFather**
   - `/newbot`
   - Nome: "Outlook Assistant"
   - Username: `outlook_assistant_bot`
   - Salva il **token**

2. **Configura credenziali in n8n**
   - Credentials → Add → Telegram API
   - Access Token: `<token-da-botfather>`

#### Workflow: Telegram Command Handler

```
[Telegram Trigger] → [Parse Command] → [Execute MCP Tool] → [Format Response] → [Reply Telegram]
```

#### Comandi Supportati

**Parse Command (Function Node):**
```javascript
const message = $input.item.json.message.text;
const chatId = $input.item.json.message.chat.id;

// Parse comando
let command = '';
let args = {};

if (message.startsWith('/briefing')) {
  command = 'briefing';

} else if (message.startsWith('/unread')) {
  command = 'unread';
  const match = message.match(/\/unread\s+(\d+)?/);
  args.days = match && match[1] ? parseInt(match[1]) : 1;

} else if (message.startsWith('/pending')) {
  command = 'pending';
  const match = message.match(/\/pending\s+(\d+)?/);
  args.days = match && match[1] ? parseInt(match[1]) : 7;

} else if (message.startsWith('/search')) {
  command = 'search';
  args.term = message.replace('/search', '').trim();

} else if (message.startsWith('/events')) {
  command = 'events';
  const match = message.match(/\/events\s+(\d+)?/);
  args.days = match && match[1] ? parseInt(match[1]) : 7;

} else if (message.startsWith('/help')) {
  command = 'help';

} else {
  command = 'unknown';
}

return {
  json: {
    command: command,
    args: args,
    chatId: chatId,
    originalMessage: message
  }
};
```

#### Switch per Comandi (Switch Node)

**Routing basato su `{{$json.command}}`**

**Route 1: Briefing**
- HTTP Request → `list_recent_emails` (unread=true, days=1)
- HTTP Request → `list_pending_replies` (days=7)
- HTTP Request → `list_upcoming_events` (days=1)
- Function → Format briefing
- Telegram → Send message

**Route 2: Unread**
- HTTP Request → `list_recent_emails`
  ```json
  {
    "arguments": {
      "days": "{{$json.args.days}}",
      "unread_only": true,
      "max_results": 10
    }
  }
  ```

**Route 3: Search**
- HTTP Request → `search_emails`
  ```json
  {
    "arguments": {
      "search_term": "{{$json.args.term}}",
      "days": 30,
      "include_all_folders": true,
      "max_results": 10
    }
  }
  ```

**Route 4: Events**
- HTTP Request → `list_upcoming_events`
  ```json
  {
    "arguments": {
      "days": "{{$json.args.days}}",
      "include_all_calendars": true,
      "max_results": 20
    }
  }
  ```

**Route 5: Help**
- Function → Return help text
  ```javascript
  return {
    json: {
      text: `🤖 Comandi Outlook Assistant

/briefing - Riepilogo giornaliero
/unread [giorni] - Email non lette (default: 1)
/pending [giorni] - Email senza risposta (default: 7)
/search <termine> - Cerca email
/events [giorni] - Eventi prossimi (default: 7)
/help - Mostra questo messaggio

Esempi:
  /unread 3
  /search contratto Acme
  /events 14`
    }
  };
  ```

#### Format Response (Function per ogni route)

**Esempio per Unread:**
```javascript
const emails = $input.item.json.emails || [];

if (emails.length === 0) {
  return {
    json: {
      text: '✅ Nessuna email non letta!'
    }
  };
}

let text = `📧 Email Non Lette: ${emails.length}\n\n`;

emails.forEach((email, i) => {
  text += `${i+1}. *${email.from}*\n`;
  text += `   ${email.subject}\n`;
  text += `   📅 ${email.received_time}\n`;
  if (email.has_attachments) {
    text += `   📎 Allegati\n`;
  }
  text += `\n`;
});

return {
  json: {
    text: text,
    parse_mode: 'Markdown'
  }
};
```

---

## Assistente Vocale (Webhook)

### Integrazione con Assistenti Vocali (Alexa, Google Home, Siri)

Usa webhook per integrare con servizi vocali come iOS Shortcuts o IFTTT.

#### Workflow: Webhook Processor

```
[Webhook] → [Authenticate] → [Parse Intent] → [Execute Action] → [Return Response]
```

#### Webhook Configuration

**URL:** `https://your-domain.com/webhook/outlook-assistant`
**Method:** POST
**Auth:** Bearer Token (configura in n8n)

**Payload:**
```json
{
  "intent": "get_briefing|get_unread|search_email|create_event",
  "parameters": {
    "days": 7,
    "search_term": "contratto",
    "event_subject": "Riunione"
  },
  "user_id": "davide"
}
```

#### iOS Shortcut Example

**Shortcut: "Outlook Briefing"**

1. **Get Current Date** → Variable: CurrentDate
2. **Text:** "Dammi il briefing email"
3. **Get Contents of URL:**
   - URL: `https://your-domain.com/webhook/outlook-assistant`
   - Method: POST
   - Headers: `Authorization: Bearer YOUR_TOKEN`
   - Body:
   ```json
   {
     "intent": "get_briefing",
     "parameters": {},
     "user_id": "davide"
   }
   ```
4. **Show Result:** Variable: Response
5. **Speak Text:** Response → summary

**Attivazione Siri:** "Hey Siri, Outlook Briefing"

---

## Workflow Avanzati

### 1. Monitoraggio SLA con Alert

**Scenario:** Alert se email cliente non risposta entro 24h

```
[Schedule ogni ora] → [List Pending] → [Calculate SLA] → [Filter Violations] → [Alert Team]
```

**Calculate SLA (Function):**
```javascript
const pending = $input.item.json.emails || [];
const SLA_HOURS = 24;
const now = Date.now();

const violations = pending
  .map(email => {
    const received = new Date(email.received_time);
    const hoursAgo = (now - received) / (1000 * 60 * 60);
    return {
      ...email,
      hoursAgo: Math.floor(hoursAgo),
      violation: hoursAgo > SLA_HOURS
    };
  })
  .filter(e => e.violation)
  .sort((a, b) => b.hoursAgo - a.hoursAgo);

return violations.map(v => ({ json: v }));
```

**Alert (Multiple channels):**
- Telegram → Team chat
- Slack → #support channel
- Email → Manager
- SMS → On-call person (via Twilio)

---

### 2. AI Email Classifier & Router

**Scenario:** Classifica email in categorie e instrada automaticamente

```
[New Email] → [AI Classify] → [Switch Category] → [Route to Folder] → [Notify Team Member]
```

**AI Classify (OpenAI):**
```
Classifica questa email in una categoria:

Email:
Da: {{$json.from}}
Oggetto: {{$json.subject}}
Corpo: {{$json.body_preview}}

Categorie possibili:
- support: Richieste supporto tecnico
- sales: Opportunità vendita
- billing: Fatturazione/pagamenti
- partnership: Proposte partnership
- hr: Risorse umane
- spam: Spam/marketing non richiesto

Output JSON:
{
  "category": "...",
  "confidence": 0-1,
  "reason": "breve spiegazione",
  "suggested_assignee": "nome persona"
}
```

**Route to Folder (Switch + HTTP):**
```javascript
// Switch basato su category
switch ($json.category) {
  case 'support':
    return { folder: 'Support', assignee: 'team-support@example.com' };
  case 'sales':
    return { folder: 'Vendite', assignee: 'sales@example.com' };
  case 'billing':
    return { folder: 'Amministrazione', assignee: 'admin@example.com' };
  // ...
}
```

---

### 3. Meeting Prep Automation

**Scenario:** 30 min prima di ogni meeting, prepara automaticamente briefing partecipanti

```
[Schedule ogni 15min] → [Get Upcoming Events] → [Filter Next 30min] → [For Each Event]
                                                                              ↓
                           [Send Briefing] ← [Format] ← [Search Emails] ← [Get Attendees]
```

**Search Emails per Attendee:**
```javascript
// Per ogni partecipante
const attendees = $json.attendees || [];

const searches = attendees.map(async (email) => {
  const response = await fetch('http://localhost:8000/tools/search_emails', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      arguments: {
        search_term: email,
        days: 30,
        max_results: 5
      }
    })
  });
  return {
    attendee: email,
    recentEmails: await response.json()
  };
});

return await Promise.all(searches);
```

---

## Best Practices n8n + Outlook MCP

### ✅ Raccomandazioni

1. **Error Handling:**
   - Sempre configurare "Continue On Fail" per nodi HTTP
   - Usa nodo "Error Trigger" per gestire fallimenti
   - Log errori in file/database per debugging

2. **Performance:**
   - Usa "Batch" mode per elaborare multiple email
   - Limita `max_results` a valori ragionevoli (≤50)
   - Cache response in variabili n8n quando possibile

3. **Security:**
   - Usa credenziali n8n per API keys
   - Non loggare dati sensibili
   - Valida sempre input da webhook pubblici

4. **Monitoring:**
   - Configura webhook failures su Slack/Telegram
   - Monitora execution times via n8n metrics
   - Set timeout appropriati (default: 300s)

5. **Testing:**
   - Testa workflow con `send=False` prima
   - Usa account test per development
   - Implementa dry-run mode con variabili

### ❌ Da Evitare

- ❌ Batch troppo grandi (>100 items) senza chunking
- ❌ Loop infiniti senza exit condition
- ❌ Hardcode credenziali/tokens nei workflow
- ❌ Inviare email (`send=true`) senza validazione
- ❌ Ignorare rate limits API

---

## Template Workflow Pronti

### Download JSON Workflows

**1. Daily Briefing:**
```json
{
  "name": "Outlook Daily Briefing",
  "nodes": [
    // ... (vedere esempi sopra per configurazione completa)
  ]
}
```

**2. AI Email Assistant:**
```json
{
  "name": "AI Email Assistant",
  "nodes": [
    // ...
  ]
}
```

**3. Telegram Bot:**
```json
{
  "name": "Telegram Outlook Bot",
  "nodes": [
    // ...
  ]
}
```

---

## Supporto e Risorse

- **n8n Documentation:** https://docs.n8n.io
- **n8n Community:** https://community.n8n.io
- **Template Library:** https://n8n.io/workflows

Per problemi specifici con Outlook MCP + n8n:
- Verifica che il server HTTP sia raggiungibile: `curl http://localhost:8000/health`
- Controlla log n8n: `~/.n8n/logs/`
- Testa endpoint manualmente con Postman/curl prima di integrare in workflow

---

Buon lavoro con il tuo assistente automatico! 🚀
