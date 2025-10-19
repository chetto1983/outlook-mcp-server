"""MCP tools for domain-based folder rules (Clienti/<dominio>/...)."""

from __future__ import annotations

from typing import Optional, List, Any, Dict

from ..features import feature_gate
from outlook_mcp.toolkit import mcp_tool

from outlook_mcp import logger
from outlook_mcp.utils import coerce_bool


def _connect():
    from outlook_mcp import connect_to_outlook

    return connect_to_outlook()


def _resolve(namespace, *, email_number: Optional[int], message_id: Optional[str] = None):
    from outlook_mcp.services.email import resolve_mail_item

    return resolve_mail_item(namespace, email_number=email_number, message_id=message_id)


def _ensure_structure(namespace, domain: str, root_folder_name: str, subfolders: Optional[List[str]]):
    from outlook_mcp.services.email import ensure_domain_folder_structure

    return ensure_domain_folder_structure(namespace, domain, root_folder_name, subfolders)


def _extract_domain(address: Optional[str]) -> Optional[str]:
    from outlook_mcp.services.email import extract_email_domain

    return extract_email_domain(address)


def _derive_sender(entry: Dict[str, Any]) -> Optional[str]:
    from outlook_mcp.services.email import derive_sender_email

    return derive_sender_email(entry)


@mcp_tool()
@feature_gate(group="domain.rules")
def ensure_domain_folder(
    email_number: Optional[int] = None,
    sender_email: Optional[str] = None,
    root_folder_name: Optional[str] = None,
    subfolders: Optional[str] = None,
) -> str:
    """Verifica/crea la gerarchia di cartelle per il dominio del mittente.

    Se `email_number` e' fornito, usa il mittente del messaggio in cache.
    """
    try:
        target_email = sender_email
        email_entry: Optional[Dict[str, Any]] = None
        if email_number:
            from outlook_mcp import email_cache

            if not email_cache:
                return "Errore: nessun elenco messaggi attivo. Mostra prima le email e riprova."
            email_entry = email_cache.get(email_number)
            if not email_entry:
                return f"Errore: il messaggio #{email_number} non e presente nella cache corrente."
            target_email = target_email or _derive_sender(email_entry)
        if not target_email:
            return "Errore: specifica un mittente (sender_email) oppure il numero di un messaggio gia elencato."

        domain = _extract_domain(target_email)
        if not domain:
            return f"Errore: impossibile determinare il dominio dal mittente '{target_email}'."

        custom_subfolders: Optional[List[str]] = None
        if subfolders:
            custom_subfolders = [folder.strip() for folder in subfolders.split("|") if folder.strip()]

        _, namespace = _connect()
        domain_folder, domain_created, created_subfolders = _ensure_structure(
            namespace=namespace,
            domain=domain,
            root_folder_name=root_folder_name or "Clienti",
            subfolders=custom_subfolders or ["Da leggere", "In lavorazione", "Archivio"],
        )
        folder_path = getattr(domain_folder, "FolderPath", f"{domain_folder}")
        summary_parts = [f"Cartella dominio '{domain}' pronta: {folder_path}"]
        if domain_created:
            summary_parts.append("Cartella dominio creata ex novo.")
        if created_subfolders:
            summary_parts.append(
                "Sottocartelle create: " + ", ".join(created_subfolders)
            )
        return " ".join(summary_parts)
    except Exception as exc:
        logger.exception("Errore durante ensure_domain_folder (email_number=%s).", email_number)
        return f"Errore durante la verifica/creazione della cartella dominio: {exc}"


@mcp_tool()
@feature_gate(group="domain.rules")
def move_email_to_domain_folder(
    email_number: int,
    root_folder_name: Optional[str] = None,
    create_if_missing: bool = True,
    subfolders: Optional[str] = None,
) -> str:
    """Sposta un messaggio nella cartella dedicata al suo dominio mittente.

    Crea automaticamente la struttura se `create_if_missing=True`.
    """
    try:
        from outlook_mcp import email_cache

        if not email_cache or email_number not in email_cache:
            return "Errore: nessun elenco messaggi attivo o numero non valido."

        email_entry = email_cache[email_number]
        sender = _derive_sender(email_entry)
        if not sender:
            return "Errore: il messaggio non contiene un mittente valido."
        domain = _extract_domain(sender)
        if not domain:
            return f"Errore: impossibile determinare il dominio dal mittente '{sender}'."

        _, namespace = _connect()

        if coerce_bool(create_if_missing):
            custom_subfolders = [seg.strip() for seg in subfolders.split("|") if seg.strip()] if subfolders else None
            domain_folder, _, _ = _ensure_structure(
                namespace=namespace,
                domain=domain,
                root_folder_name=root_folder_name or "Clienti",
                subfolders=custom_subfolders or ["Da leggere", "In lavorazione", "Archivio"],
            )
        else:
            inbox = namespace.GetDefaultFolder(6)
            root_folder = None
            try:
                for sub in inbox.Folders:
                    if sub.Name.lower() == (root_folder_name or "Clienti").lower():
                        root_folder = sub
                        break
            except Exception:
                root_folder = None
            domain_folder = None
            if root_folder:
                try:
                    for sub in root_folder.Folders:
                        if sub.Name.lower() == domain.lower():
                            domain_folder = sub
                            break
                except Exception:
                    pass
            if not domain_folder:
                return (
                    f"Errore: cartella dominio '{domain}' non trovata sotto '{root_folder_name or 'Clienti'}'. "
                    "Imposta create_if_missing=True per crearla automaticamente."
                )

        try:
            _, mail_item = _resolve(namespace, email_number=email_number)
        except Exception as exc:
            return f"Errore: {exc}"

        mail_item.Move(domain_folder)
        folder_path = getattr(domain_folder, "FolderPath", f"{domain_folder}")
        return (
            f"Messaggio #{email_number} spostato nella cartella dominio '{domain}' "
            f"({folder_path})."
        )
    except Exception as exc:
        logger.exception("Errore durante move_email_to_domain_folder per messaggio #%s.", email_number)
        return f"Errore durante lo spostamento nella cartella dominio: {exc}"
