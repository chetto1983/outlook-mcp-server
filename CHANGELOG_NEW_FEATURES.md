# Nuove Funzionalità Implementate

Questo documento descrive le nuove funzionalità aggiunte all'Outlook MCP Server basandosi sull'analisi delle API Microsoft Outlook disponibili.

## 1. Gestione Tasks (Attività)

### Nuovi Tools
- **`list_tasks`**: Elenca le attività con filtri su giorni, cartelle, stato completamento
- **`search_tasks`**: Cerca attività per parole chiave
- **`get_task_by_number`**: Recupera dettagli completi di un'attività
- **`create_task`**: Crea nuove attività con priorità, scadenze, promemoria
- **`update_task`**: Aggiorna attività esistenti (oggetto, stato, scadenza, priorità, %)
- **`mark_task_complete`**: Contrassegna rapidamente un'attività come completata
- **`delete_task`**: Elimina un'attività

### Feature Group
`tasks`

### Esempi di Utilizzo

```python
# Lista attività non completate degli ultimi 30 giorni
list_tasks(days=30, include_completed=False)

# Crea un'attività con priorità alta e scadenza
create_task(
    subject="Preparare presentazione Q4",
    body="Slide su revenue e proiezioni",
    due_date="2025-10-25T17:00",
    priority="alta",
    reminder_time="2025-10-25T15:00"
)

# Aggiorna percentuale completamento
update_task(task_number=1, percent_complete=75)

# Segna come completata
mark_task_complete(task_number=1)
```

### Proprietà Supportate
- Subject, Body, Preview
- Status (Non iniziata, In corso, Completata, In attesa, Differita)
- Priority (Bassa, Normale, Alta)
- Due Date, Start Date, Completed Date
- Percent Complete (0-100%)
- Reminder (Set/Time)
- Owner, Categories
- Folder Path

---

## 2. Rules (Regole di Outlook)

### Nuovi Tools
- **`list_rules`**: Elenca tutte le regole configurate con stato e dettagli
- **`get_rule_details`**: Mostra condizioni e azioni di una regola specifica
- **`create_move_rule`**: Crea regole semplificate per spostare email automaticamente
- **`enable_disable_rule`**: Abilita o disabilita una regola esistente
- **`delete_rule`**: Elimina una regola

### Feature Group
`rules`

### Esempi di Utilizzo

```python
# Elenca tutte le regole configurate
list_rules()

# Crea regola per spostare email da un mittente
create_move_rule(
    rule_name="Newsletter Tech",
    from_address="newsletter@techsite.com",
    target_folder_name="Newsletter",
    mark_as_read=True,
    enabled=True
)

# Crea regola basata sull'oggetto
create_move_rule(
    rule_name="Fatture",
    subject_contains="Fattura",
    target_folder_path="\\Amministrazione\\Fatture",
    enabled=True
)

# Disabilita temporaneamente una regola
enable_disable_rule(rule_name="Newsletter Tech", enabled=False)

# Dettagli completi di una regola
get_rule_details(rule_name="Fatture")
```

### Condizioni Supportate
- From (mittente)
- Subject Contains (oggetto contiene)
- Body Contains (corpo contiene)
- Sent To (inviato a)
- Importance (priorità)
- Message Size (dimensione)

### Azioni Supportate
- Move to Folder (sposta in cartella)
- Copy to Folder (copia in cartella)
- Mark as Read (segna come letto)
- Delete (elimina)
- Assign Category (assegna categoria)
- Forward (inoltra)
- Stop Processing (interrompi elaborazione)

---

## 3. Proprietà Email Avanzate

### Funzionalità Aggiunte

#### A. Importance (Priorità)
Parametro `importance` aggiunto a `compose_email` e `reply_to_email_by_number`.

Valori: `"bassa"`, `"normale"`, `"alta"`

```python
compose_email(
    recipient_email="manager@company.com",
    subject="Urgente: Approvazione richiesta",
    body="...",
    importance="alta"
)
```

#### B. Sensitivity (Riservatezza)
Parametro `sensitivity` aggiunto a `compose_email` e `reply_to_email_by_number`.

Valori: `"normale"`, `"personale"`, `"privato"`, `"confidenziale"`

```python
compose_email(
    recipient_email="hr@company.com",
    subject="Informazioni personali",
    body="...",
    sensitivity="confidenziale"
)
```

#### C. Read/Delivery Receipts (Conferme di Lettura/Consegna)
Parametri `request_read_receipt` e `request_delivery_receipt`.

```python
compose_email(
    recipient_email="client@partner.com",
    subject="Proposta commerciale",
    body="...",
    request_read_receipt=True,
    request_delivery_receipt=True
)
```

#### D. Voting Buttons (Pulsanti di Voto)
Parametro `voting_options` in `compose_email` per sondaggi rapidi.

```python
compose_email(
    recipient_email="team@company.com",
    subject="Sondaggio data meeting",
    body="Quale data preferite?",
    voting_options="Lunedì;Martedì;Mercoledì"
)
```

#### E. Follow-up Flags (Contrassegni)
Nuovo tool **`flag_email`** per impostare contrassegni follow-up con scadenze e promemoria.

```python
# Contrassegna email per follow-up
flag_email(
    email_number=5,
    flag_status="Follow up",
    due_date="2025-10-23T17:00",
    reminder_time="2025-10-23T09:00"
)

# Rimuovi contrassegno
flag_email(email_number=5, clear_flag=True)
```

---

## 4. Free/Busy (Disponibilità Calendario)

### Nuovi Tools
- **`get_freebusy_info`**: Recupera dati di disponibilità per un destinatario
- **`find_free_time_slots`**: Trova slot temporali liberi comuni per più partecipanti

### Feature Group
`freebusy`

### Esempi di Utilizzo

```python
# Verifica disponibilità di un collega
get_freebusy_info(
    recipient_email="collega@company.com",
    start_date="2025-10-22T08:00",
    end_date="2025-10-22T18:00",
    interval_minutes=30,
    merge_slots=True
)

# Trova slot liberi comuni per un meeting
find_free_time_slots(
    attendees="alice@company.com,bob@company.com,charlie@company.com",
    duration_minutes=60,
    start_date="2025-10-22",
    end_date="2025-10-25",
    working_hours_start="09:00",
    working_hours_end="17:00",
    max_results=10
)
```

### Stati di Disponibilità
- **Libero** (0): Nessun impegno
- **Provvisorio** (1): Evento tentativo
- **Occupato** (2): Occupato
- **Fuori ufficio** (3): Out of office

### Funzionalità
- Intervalli configurabili (1-1440 minuti)
- Unione automatica slot consecutivi con stesso stato
- Ricerca intelligente solo durante orari lavorativi
- Supporto multi-partecipante per meeting

---

## Nuovi File Creati

### Services
- `outlook_mcp/services/tasks.py` - Logica business per gestione tasks
- Estensioni in `outlook_mcp/services/common.py` - Helper per proprietà avanzate

### Tools
- `outlook_mcp/tools/tasks.py` - MCP tools per tasks
- `outlook_mcp/tools/rules.py` - MCP tools per rules
- `outlook_mcp/tools/freebusy.py` - MCP tools per free/busy
- Estensioni in `outlook_mcp/tools/email_actions.py` - Proprietà avanzate email

### Cache & Constants
- Aggiunta `task_cache` in `outlook_mcp/cache.py`
- Nuove costanti task-related in `outlook_mcp/constants.py`

---

## Feature Flags

Per abilitare/disabilitare le nuove funzionalità, usa `features.json`:

```json
{
  "enabled_groups": ["tasks", "rules", "freebusy"],
  "disabled_groups": [],
  "enabled_tools": [],
  "disabled_tools": []
}
```

O disabilita funzionalità specifiche:

```json
{
  "disabled_groups": ["rules"],
  "disabled_tools": ["create_move_rule"]
}
```

---

## Compatibilità

Tutte le nuove funzionalità:
- ✅ Mantengono compatibilità con il codice esistente
- ✅ Seguono le convenzioni di naming esistenti (italiano)
- ✅ Integrano il sistema di cache esistente
- ✅ Supportano feature flags per controllo granulare
- ✅ Includono logging dettagliato
- ✅ Gestiscono errori in modo consistente con il resto del sistema

---

## Testing

Per testare le nuove funzionalità:

1. Verifica che Outlook sia aperto e accessibile
2. Abilita i gruppi desiderati in `features.json`
3. Riavvia il server MCP
4. Usa `feature_status()` per verificare tool attivi

---

## Note di Implementazione

### Tasks
- Usa Outlook Folder ID 13 (olFolderTasks)
- Supporta ricerca multi-folder
- Cache separata con TTL 20 minuti
- Gestione date intelligente (ignora date invalide come 4501)

### Rules
- Accesso via `Store.GetRules()`
- Supporto regole server-side (non solo locali)
- `create_move_rule` è semplificata, per regole complesse usare UI Outlook
- Richiede `recipient.Resolve()` per mittenti

### Free/Busy
- Usa `Recipient.FreeBusy()` API
- CompleteFormat=True per 4 stati (Free/Tentative/Busy/OOF)
- Intervallo minimo 1 minuto, massimo 1440 (24 ore)
- Ricerca slot intelligente con considerazione orari lavorativi

### Proprietà Email Avanzate
- Non richiedono nuove API, solo parametri aggiuntivi
- Backward compatible (parametri opzionali)
- Voting options formato: "Option1;Option2;Option3"
- Follow-up flags integrati con reminder Outlook

---

## Roadmap Future Miglioramenti

Feature non ancora implementate ma disponibili in Outlook API:

### Priorità Media
- **Notes**: Gestione note di Outlook
- **Distribution Lists**: Gruppi di contatti
- **Conversation Actions**: Clean up, ignore, move conversation
- **Journal**: Tracciamento attività automatico

### Priorità Bassa
- **Form Regions**: Custom forms
- **Search Folders**: Cartelle di ricerca virtuali
- **Public Folders**: Accesso cartelle pubbliche Exchange
- **MAPI Properties**: Proprietà estese
- **PST Management**: Backup/import archivi

Queste feature possono essere implementate seguendo lo stesso pattern usato per tasks, rules e free/busy.
