"""Helpers for resilient COM interactions with Outlook."""

from __future__ import annotations

import time
from typing import Callable, Optional, Tuple, TypeVar

from outlook_mcp import logger

T = TypeVar("T")

_TRANSIENT_MARKERS = [
    "rpc_e_servercall_retrylater",
    "call was rejected by callee",
    "0x8001010a",
]
_PERMANENT_NOT_FOUND = [
    "mapi_e_not_found",
    "object could not be found",
    "does not exist",
]


class OutlookComError(RuntimeError):
    """Wrap Outlook COM failures with additional guidance."""

    def __init__(
        self,
        description: str,
        original_exception: Exception,
        suggestion: Optional[str],
        transient: bool,
    ) -> None:
        message = f"{description}: {original_exception}"
        if suggestion:
            message = f"{message} ({suggestion})"
        super().__init__(message)
        self.description = description
        self.original_exception = original_exception
        self.suggestion = suggestion
        self.transient = transient


def _classify_exception(exc: Exception) -> Tuple[bool, Optional[str]]:
    text = str(exc).lower()
    for marker in _TRANSIENT_MARKERS:
        if marker in text:
            return True, "Outlook sembra occupato. Chiudi eventuali finestre di dialogo e riprova tra qualche secondo."
    for marker in _PERMANENT_NOT_FOUND:
        if marker in text:
            return False, "L'elemento non esiste piu' o e' stato spostato."
    return False, "Errore COM imprevisto. Riavvia Outlook se l'errore persiste."


def run_com_call(
    action: Callable[[], T],
    description: str,
    *,
    retries: int = 1,
    delay_seconds: float = 0.4,
) -> T:
    """Execute a COM call with retry logic and structured logging."""

    attempt = 0
    while True:
        try:
            return action()
        except Exception as exc:  # pylint: disable=broad-except
            attempt += 1
            transient, suggestion = _classify_exception(exc)
            logger.warning(
                "Operazione COM fallita (%s) [tentativo %s]: %s",
                description,
                attempt,
                exc,
            )
            if transient and attempt <= retries:
                time.sleep(delay_seconds)
                continue
            raise OutlookComError(description, exc, suggestion, transient) from exc


def wrap_com_exception(description: str, exc: Exception) -> OutlookComError:
    """Convert a raw COM exception into an OutlookComError without performing retries."""
    transient, suggestion = _classify_exception(exc)
    return OutlookComError(description, exc, suggestion, transient)


__all__ = ["run_com_call", "OutlookComError", "wrap_com_exception"]
