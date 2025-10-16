"""Thin wrapper around Outlook COM connection logic."""

import win32com.client  # type: ignore

from .logger import logger


def connect_to_outlook():
    """Connect to Outlook application using COM."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        logger.debug("Connessione a Outlook MAPI completata.")
        return outlook, namespace
    except Exception as exc:  # pragma: no cover - depends on Outlook runtime
        logger.exception("Errore durante la connessione a Outlook.")
        raise Exception(f"Impossibile connettersi a Outlook: {exc}") from exc


__all__ = ["connect_to_outlook"]
