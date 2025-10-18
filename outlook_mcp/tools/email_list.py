from __future__ import annotations

from typing import Any, Optional, List, Dict

from ..features import feature_gate
from outlook_mcp_server import mcp

from outlook_mcp import logger
from outlook_mcp import folders as folder_service
from outlook_mcp.utils import coerce_bool, ensure_string_list
from outlook_mcp import (
    MAX_DAYS,
    DEFAULT_MAX_RESULTS,
    MAX_CONVERSATION_LOOKBACK_DAYS,
    DEFAULT_CONVERSATION_SAMPLE_LIMIT,
    PENDING_SCAN_MULTIPLIER,
)

# Reuse shared helpers from server to avoid duplication
from outlook_mcp.services.email import (
    get_emails_from_folder,
    get_all_mail_folders,
    collect_emails_across_folders,
    present_email_listing,
    collect_user_addresses,
    normalize_email_address,
    email_has_user_reply_with_context,
    build_conversation_outline,
)

def _connect():
    from outlook_mcp import connect_to_outlook

    return connect_to_outlook()


@mcp.tool()
@feature_gate(group="email.list")
def list_recent_emails(
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
    folder_ids: Optional[Any] = None,
    folder_paths: Optional[Any] = None,
    offset: int = 0,
    unread_only: bool = False,
) -> str:
    """Elenca i messaggi piu' recenti con filtri su giorni/cartelle/anteprima."""
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        logger.warning("Valore 'days' non valido passato a list_recent_emails: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a list_recent_emails: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"
    try:
        offset_value = int(offset)
        if offset_value < 0:
            raise ValueError
    except (TypeError, ValueError):
        logger.warning("Valore 'offset' non valido passato a list_recent_emails: %s", offset)
        return "Errore: 'offset' deve essere un intero maggiore o uguale a zero."

    include_preview_bool = coerce_bool(include_preview)
    include_all_bool = coerce_bool(include_all_folders)
    unread_only_bool = coerce_bool(unread_only)
    folder_id_list = ensure_string_list(folder_ids)
    folder_path_list = ensure_string_list(folder_paths)
    logger.info(
        (
            "list_recent_emails chiamato con giorni=%s cartella=%s max_risultati=%s "
            "anteprima=%s tutte_le_cartelle=%s offset=%s unread_only=%s folder_ids=%s folder_paths=%s"
        ),
        days,
        folder_name,
        max_results,
        include_preview_bool,
        include_all_bool,
        offset_value,
        unread_only_bool,
        folder_id_list,
        folder_path_list,
    )

    try:
        _, namespace = _connect()

        emails: List[Dict[str, Any]]
        folder_display: str
        if folder_id_list or folder_path_list:
            selected_folders: List = []
            failures: List[str] = []
            for entry_id in folder_id_list:
                folder, attempts = folder_service.resolve_folder(namespace, folder_id=entry_id)
                if folder:
                    selected_folders.append(folder)
                else:
                    detail = ", ".join(attempts) if attempts else "non trovato"
                    failures.append(f"ID {entry_id}: {detail}")
            for folder_path in folder_path_list:
                folder, attempts = folder_service.resolve_folder(namespace, folder_path=folder_path)
                if folder:
                    selected_folders.append(folder)
                else:
                    detail = ", ".join(attempts) if attempts else "non trovato"
                    failures.append(f"Percorso {folder_path}: {detail}")
            if not selected_folders:
                detail = "; ".join(failures) if failures else "cartelle non trovate."
                return f"Errore: impossibile individuare le cartelle richieste ({detail})."
            emails = collect_emails_across_folders(selected_folders, days)
            folder_display = "Cartelle selezionate"
        elif include_all_bool:
            if folder_name:
                logger.info("Parametro folder_name ignorato perche include_all_folders=True.")
            folders = get_all_mail_folders(namespace)
            emails = collect_emails_across_folders(folders, days)
            folder_display = "Tutte le cartelle"
        else:
            if folder_name:
                folder = folder_service.get_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella '{folder_name}' non trovata"
            else:
                folder = namespace.GetDefaultFolder(6)
            folder_display = f"'{folder_name}'" if folder_name else "Posta in arrivo"
            emails = get_emails_from_folder(folder, days)

        if unread_only_bool:
            emails = [email for email in emails if email.get("unread")]

        return present_email_listing(
            emails=emails,
            folder_display=folder_display,
            days=days,
            max_results=max_results,
            include_preview=include_preview_bool,
            log_context="list_recent_emails",
            offset=offset_value,
        )
    except Exception as exc:
        logger.exception("Errore nel recupero dei messaggi per la cartella '%s'.", folder_name or "Posta in arrivo")
        return f"Errore durante il recupero dei messaggi: {exc}"


@mcp.tool()
@feature_gate(group="email.list")
def list_sent_emails(
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    offset: int = 0,
) -> str:
    """Elenca i messaggi inviati, con anteprima facoltativa e offset."""
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        logger.warning("Valore 'days' non valido passato a list_sent_emails: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a list_sent_emails: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"
    try:
        offset_value = int(offset)
        if offset_value < 0:
            raise ValueError
    except (TypeError, ValueError):
        logger.warning("Valore 'offset' non valido passato a list_sent_emails: %s", offset)
        return "Errore: 'offset' deve essere un intero maggiore o uguale a zero."

    include_preview = coerce_bool(include_preview)
    logger.info(
        "list_sent_emails chiamato con giorni=%s cartella=%s max_risultati=%s anteprima=%s offset=%s",
        days,
        folder_name,
        max_results,
        include_preview,
        offset_value,
    )

    try:
        _, namespace = _connect()

        if folder_name:
            folder = folder_service.get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Errore: cartella '{folder_name}' non trovata"
            folder_display = f"'{folder_name}'"
        else:
            folder = namespace.GetDefaultFolder(5)  # Sent Items
            folder_display = "Posta inviata"

        emails = get_emails_from_folder(folder, days)
        return present_email_listing(
            emails=emails,
            folder_display=folder_display,
            days=days,
            max_results=max_results,
            include_preview=include_preview,
            log_context="list_sent_emails",
            search_term=None,
            focus_on_recipients=True,
            offset=offset_value,
        )
    except Exception as e:
        logger.exception("Errore nel recupero dei messaggi inviati per la cartella '%s'.", folder_name or "Posta inviata")
        return f"Errore durante il recupero dei messaggi inviati: {str(e)}"


@mcp.tool()
@feature_gate(group="email.list")
def search_emails(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
    folder_ids: Optional[Any] = None,
    folder_paths: Optional[Any] = None,
    offset: int = 0,
    unread_only: bool = False,
) -> str:
    """Cerca messaggi per parole chiave (supporta 'OR') con filtri standard."""
    if not search_term:
        logger.warning("search_emails chiamato senza termine di ricerca.")
        return "Errore: inserisci un termine di ricerca"

    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        logger.warning("Valore 'days' non valido passato a search_emails: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a search_emails: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"
    try:
        offset_value = int(offset)
        if offset_value < 0:
            raise ValueError
    except (TypeError, ValueError):
        logger.warning("Valore 'offset' non valido passato a search_emails: %s", offset)
        return "Errore: 'offset' deve essere un intero maggiore o uguale a zero."

    include_preview_bool = coerce_bool(include_preview)
    include_all_bool = coerce_bool(include_all_folders)
    unread_only_bool = coerce_bool(unread_only)
    folder_id_list = ensure_string_list(folder_ids)
    folder_path_list = ensure_string_list(folder_paths)
    logger.info(
        (
            "search_emails chiamato con termine='%s' giorni=%s cartella=%s max_risultati=%s "
            "anteprima=%s tutte_le_cartelle=%s offset=%s unread_only=%s folder_ids=%s folder_paths=%s"
        ),
        search_term,
        days,
        folder_name,
        max_results,
        include_preview_bool,
        include_all_bool,
        offset_value,
        unread_only_bool,
        folder_id_list,
        folder_path_list,
    )

    try:
        _, namespace = _connect()

        emails: List[Dict[str, Any]]
        folder_display: str
        if folder_id_list or folder_path_list:
            selected_folders: List = []
            failures: List[str] = []
            for entry_id in folder_id_list:
                folder, attempts = folder_service.resolve_folder(namespace, folder_id=entry_id)
                if folder:
                    selected_folders.append(folder)
                else:
                    detail = ", ".join(attempts) if attempts else "non trovato"
                    failures.append(f"ID {entry_id}: {detail}")
            for folder_path in folder_path_list:
                folder, attempts = folder_service.resolve_folder(namespace, folder_path=folder_path)
                if folder:
                    selected_folders.append(folder)
                else:
                    detail = ", ".join(attempts) if attempts else "non trovato"
                    failures.append(f"Percorso {folder_path}: {detail}")
            if not selected_folders:
                detail = "; ".join(failures) if failures else "cartelle non trovate."
                return f"Errore: impossibile individuare le cartelle richieste ({detail})."
            emails = collect_emails_across_folders(selected_folders, days, search_term)
            folder_display = "Cartelle selezionate"
        elif include_all_bool:
            if folder_name:
                logger.info("Parametro folder_name ignorato perche include_all_folders=True.")
            folders = get_all_mail_folders(namespace)
            emails = collect_emails_across_folders(folders, days, search_term)
            folder_display = "Tutte le cartelle"
        else:
            if folder_name:
                folder = folder_service.get_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella '{folder_name}' non trovata"
            else:
                folder = namespace.GetDefaultFolder(6)
            folder_display = f"'{folder_name}'" if folder_name else "Posta in arrivo"
            emails = get_emails_from_folder(folder, days, search_term)

        if unread_only_bool:
            emails = [email for email in emails if email.get("unread")]

        return present_email_listing(
            emails=emails,
            folder_display=folder_display,
            days=days,
            max_results=max_results,
            include_preview=include_preview_bool,
            log_context="search_emails",
            search_term=search_term,
            offset=offset_value,
        )
    except Exception as exc:
        logger.exception(
            "Errore durante la ricerca dei messaggi con termine '%s' nella cartella '%s'.",
            search_term,
            folder_name or "Posta in arrivo",
        )
        return f"Errore nella ricerca dei messaggi: {exc}"


@mcp.tool()
@feature_gate(group="email.list")
def list_pending_replies(
    days: int = 14,
    folder_name: Optional[str] = None,
    max_results: int = DEFAULT_MAX_RESULTS,
    include_preview: bool = True,
    include_all_folders: bool = False,
    include_unread_only: bool = False,
    conversation_lookback_days: Optional[int] = None,
) -> str:
    """Evidenzia mail in attesa di risposta incrociando conversazione e Posta inviata."""
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        logger.warning("Valore 'days' non valido passato a list_pending_replies: %s", days)
        return f"Errore: 'days' deve essere un intero tra 1 e {MAX_DAYS}"
    if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
        logger.warning("Valore 'max_results' non valido passato a list_pending_replies: %s", max_results)
        return "Errore: 'max_results' deve essere un intero tra 1 e 200"

    include_preview_bool = coerce_bool(include_preview)
    include_all_bool = coerce_bool(include_all_folders)
    unread_only_bool = coerce_bool(include_unread_only)

    if conversation_lookback_days is None:
        lookback_days = max(days * 2, 14)
    elif isinstance(conversation_lookback_days, int) and 1 <= conversation_lookback_days <= MAX_CONVERSATION_LOOKBACK_DAYS:
        lookback_days = conversation_lookback_days
    else:
        logger.warning(
            "Valore 'conversation_lookback_days' non valido passato a list_pending_replies: %s",
            conversation_lookback_days,
        )
        return f"Errore: 'conversation_lookback_days' deve essere un intero tra 1 e {MAX_CONVERSATION_LOOKBACK_DAYS}"

    lookback_days = min(lookback_days, MAX_CONVERSATION_LOOKBACK_DAYS)
    max_processed_before_break = max(max_results * PENDING_SCAN_MULTIPLIER, max_results + 25)
    logger.info(
        (
            "list_pending_replies chiamato con giorni=%s cartella=%s max_risultati=%s "
            "anteprima=%s tutte_le_cartelle=%s solo_non_letti=%s lookback_conv=%s"
        ),
        days,
        folder_name,
        max_results,
        include_preview_bool,
        include_all_bool,
        unread_only_bool,
        lookback_days,
    )

    try:
        _, namespace = _connect()
        user_addresses = collect_user_addresses(namespace)
        normalized_user_addresses = {
            addr for addr in (normalize_email_address(addr) for addr in user_addresses) if addr
        }

        if include_all_bool:
            if folder_name:
                logger.info("Parametro folder_name ignorato poiche include_all_folders=True.")
            folders = get_all_mail_folders(namespace)
            candidate_emails = collect_emails_across_folders(
                folders,
                days,
                target_total=max_processed_before_break,
            )
            folder_display = "Tutte le cartelle (senza risposta)"
        else:
            if folder_name:
                folder = folder_service.get_folder_by_name(namespace, folder_name)
                if not folder:
                    return f"Errore: cartella '{folder_name}' non trovata"
                candidate_emails = get_emails_from_folder(folder, days)
                folder_display = f"'{folder_name}' (senza risposta)"
            else:
                folder = namespace.GetDefaultFolder(6)
                candidate_emails = get_emails_from_folder(folder, days)
                folder_display = "Posta in arrivo (senza risposta)"
            if len(candidate_emails) > max_processed_before_break:
                candidate_emails = candidate_emails[:max_processed_before_break]

        pending_emails: List[Dict[str, Any]] = []
        processed = 0
        truncated_scan = False

        for email in candidate_emails:
            processed += 1

            if unread_only_bool and not email.get("unread"):
                if processed >= max_processed_before_break and len(pending_emails) >= max_results:
                    truncated_scan = processed < len(candidate_emails)
                    break
                continue

            sender_email = normalize_email_address(email.get("sender_email")) or normalize_email_address(
                email.get("sender")
            )
            if sender_email and sender_email in normalized_user_addresses:
                if processed >= max_processed_before_break and len(pending_emails) >= max_results:
                    truncated_scan = processed < len(candidate_emails)
                    break
                continue

            already_replied = False
            related_entries: Optional[List[Dict[str, Any]]] = None
            mail_item_ref = None
            try:
                already_replied, related_entries, mail_item_ref = email_has_user_reply_with_context(
                    namespace=namespace,
                    email_data=email,
                    user_addresses=user_addresses,
                    conversation_limit=DEFAULT_CONVERSATION_SAMPLE_LIMIT,
                    lookback_days=lookback_days,
                    collect_related=True,
                )
            except Exception:
                logger.debug(
                    "Controllo risposta fallito per il messaggio %s.",
                    email.get("id"),
                    exc_info=True,
                )
                already_replied = False
                related_entries = None
                mail_item_ref = None

            if not already_replied:
                outline = build_conversation_outline(
                    namespace=namespace,
                    email_data=email,
                    lookback_days=lookback_days,
                    max_items=4,
                    preloaded_entries=related_entries,
                    mail_item=mail_item_ref,
                )
                if outline:
                    email["_conversation_outline"] = outline
                pending_emails.append(email)

            if len(pending_emails) >= max_results:
                truncated_scan = processed < len(candidate_emails)
                break

        logger.info(
            "list_pending_replies ha trovato %s messaggi da gestire su %s analizzati.",
            len(pending_emails),
            processed,
        )

        presentation = present_email_listing(
            emails=pending_emails,
            folder_display=folder_display,
            days=days,
            max_results=max_results,
            include_preview=include_preview_bool,
            log_context="list_pending_replies",
        )

        if truncated_scan:
            presentation += (
                "\nNota: la scansione Ã¨ stata troncata per motivi di performance; restringi i giorni o alza 'max_results'."
            )

        return presentation
    except Exception as exc:
        logger.exception("Errore durante list_pending_replies per la cartella '%s'.", folder_name or "Posta in arrivo")
        return f"Errore durante il calcolo delle risposte mancanti: {str(exc)}"


