"""Logging configuration shared across Outlook MCP server components."""

import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path

from .constants import LOG_DIR_NAME, LOG_FILE_NAME


def _determine_base_dir() -> Path:
    """Return the project root directory (one level above this package)."""
    package_dir = Path(__file__).resolve().parent
    return package_dir.parent


def _setup_logger() -> logging.Logger:
    """Configure application-wide logging once and reuse the same logger."""
    logger = logging.getLogger("outlook_mcp_server")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    base_dir = _determine_base_dir()
    log_dir = base_dir / LOG_DIR_NAME
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / LOG_FILE_NAME

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s")

    rotating_handler = RotatingFileHandler(
        log_path,
        maxBytes=5 * 1024 * 1024,  # 5 MB per file
        backupCount=3,
        encoding="utf-8",
    )
    rotating_handler.setFormatter(formatter)
    logger.addHandler(rotating_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    logger.debug("Logger inizializzato: scrittura su %s", log_path)
    return logger


# Configure logger at import time so every module shares the same instance.
logger = _setup_logger()

__all__ = ["logger"]
