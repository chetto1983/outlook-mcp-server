"""Thin wrapper around Outlook COM connection logic."""

import win32com.client  # type: ignore

from .logger import logger

_version_logged = False


def _log_outlook_version(app) -> None:
    """Log Outlook version information once per process."""
    global _version_logged
    if _version_logged:
        return
    try:
        version = str(getattr(app, "Version", "")).strip()
    except Exception:
        version = ""
    if version:
        logger.info("Versione Outlook rilevata: %s", version)
        try:
            major = int(version.split(".")[0])
            if major < 15:
                logger.warning(
                    "Versione Outlook %s potrebbe non supportare tutte le funzionalita' del server.",
                    version,
                )
        except Exception:
            logger.debug("Impossibile interpretare la versione Outlook: %s", version)
    else:
        logger.info("Versione di Outlook non disponibile via COM.")
    _version_logged = True


def connect_to_outlook():
    """Connect to Outlook application using COM."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        _log_outlook_version(outlook)
        logger.debug("Connessione a Outlook MAPI completata.")
        return outlook, namespace
    except Exception as exc:  # pragma: no cover - depends on Outlook runtime
        logger.exception("Errore durante la connessione a Outlook.")
        raise Exception(f"Impossibile connettersi a Outlook: {exc}") from exc


__all__ = ["connect_to_outlook"]
