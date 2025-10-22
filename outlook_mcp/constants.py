"""Shared constants used across Outlook MCP server modules."""

MAX_DAYS = 30
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
MAX_EMAIL_SCAN_PER_FOLDER = 400
DEFAULT_DOMAIN_ROOT_NAME = "Clienti"
DEFAULT_DOMAIN_SUBFOLDERS = [
    "00 - Generale",
    "01 - Offerte",
    "02 - Ordini",
    "03 - Service",
]
DEFAULT_ITEM_TYPE_LABELS = {
    0: "Posta",
    1: "Calendario",
    2: "Contatti",
    3: "Attivita",
    4: "Diario",
    5: "Note",
    6: "Post",
}
ITEM_TYPE_NAME_MAP = {
    "mail": 0,
    "posta": 0,
    "posta in arrivo": 0,
    "calendar": 1,
    "calendario": 1,
    "contact": 2,
    "contacts": 2,
    "contatti": 2,
    "task": 3,
    "tasks": 3,
    "attivita": 3,
    "journal": 4,
    "diario": 4,
    "note": 5,
    "post": 6,
}

# Task-specific constants
MAX_TASK_DAYS = 365
DEFAULT_TASK_MAX_RESULTS = 50
TASK_STATUS_MAP = {
    0: "Non iniziata",
    1: "In corso",
    2: "Completata",
    3: "In attesa",
    4: "Differita",
}
TASK_PRIORITY_MAP = {
    0: "Bassa",
    1: "Normale",
    2: "Alta",
}
TASK_STATUS_REVERSE_MAP = {
    "non iniziata": 0,
    "non iniziato": 0,
    "in corso": 1,
    "in progress": 1,
    "completata": 2,
    "completato": 2,
    "complete": 2,
    "in attesa": 3,
    "waiting": 3,
    "differita": 4,
    "deferred": 4,
}
TASK_PRIORITY_REVERSE_MAP = {
    "bassa": 0,
    "low": 0,
    "normale": 1,
    "normal": 1,
    "alta": 2,
    "high": 2,
}
