from __future__ import annotations

from typing import Any, Optional

from ..features import feature_gate
from outlook_mcp.toolkit import mcp_tool

from outlook_mcp import logger
from outlook_mcp.utils import coerce_bool, safe_folder_path
from outlook_mcp import folders as folder_service


def _connect():
    from outlook_mcp import connect_to_outlook

    return connect_to_outlook()


@mcp_tool()
@feature_gate(group="folders")
def list_folders(
    root_folder_id: Optional[str] = None,
    root_folder_path: Optional[str] = None,
    root_folder_name: Optional[str] = None,
    max_depth: int = 2,
    include_counts: bool = True,
    include_ids: bool = False,
    include_store: bool = False,
    include_paths: bool = True,
) -> str:
    """Elenca le cartelle di Outlook a partire da una radice opzionale."""
    if not isinstance(max_depth, int) or max_depth < 0 or max_depth > 10:
        logger.warning("Valore 'max_depth' non valido passato a list_folders: %s", max_depth)
        return "Errore: 'max_depth' deve essere un intero compreso tra 0 e 10."

    include_counts_flag = coerce_bool(include_counts)
    include_ids_flag = coerce_bool(include_ids)
    include_store_flag = coerce_bool(include_store)
    include_paths_flag = coerce_bool(include_paths)

    logger.info(
        (
            "list_folders chiamato (root_id=%s root_path=%s root_name=%s profondita=%s "
            "contatori=%s ids=%s store=%s paths=%s)."
        ),
        root_folder_id,
        root_folder_path,
        root_folder_name,
        max_depth,
        include_counts_flag,
        include_ids_flag,
        include_store_flag,
        include_paths_flag,
    )

    try:
        _, namespace = _connect()
        return folder_service.list_folders(
            namespace,
            root_folder_id=root_folder_id,
            root_folder_path=root_folder_path,
            root_folder_name=root_folder_name,
            max_depth=max_depth,
            include_counts=include_counts,
            include_ids=include_ids,
            include_store=include_store,
            include_paths=include_paths,
        )
    except Exception as exc:
        logger.exception("Errore durante l'elenco delle cartelle di Outlook.")
        return f"Errore durante l'elenco delle cartelle: {exc}"


@mcp_tool()
@feature_gate(group="folders")
def get_folder_metadata(
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
    include_children: bool = False,
    max_children: int = 20,
    include_counts: bool = True,
) -> str:
    """Mostra metadati dettagliati di una cartella (ID, path, contatori)."""
    if not isinstance(max_children, int) or max_children < 0:
        return "Errore: 'max_children' deve essere un intero non negativo."

    include_children_flag = coerce_bool(include_children)
    include_counts_flag = coerce_bool(include_counts)
    logger.info(
        "get_folder_metadata chiamato (id=%s path=%s nome=%s figli=%s max=%s contatori=%s).",
        folder_id,
        folder_path,
        folder_name,
        include_children_flag,
        max_children,
        include_counts_flag,
    )

    try:
        _, namespace = _connect()
        return folder_service.folder_metadata(
            namespace,
            folder_id=folder_id,
            folder_path=folder_path,
            folder_name=folder_name,
            include_children=include_children,
            max_children=max_children,
            include_counts=include_counts,
        )
    except Exception as exc:
        logger.exception("Errore durante get_folder_metadata.")
        return f"Errore durante il recupero dei metadati della cartella: {exc}"


@mcp_tool()
@feature_gate(group="folders")
def create_folder(
    new_folder_name: str,
    parent_folder_id: Optional[str] = None,
    parent_folder_path: Optional[str] = None,
    parent_folder_name: Optional[str] = None,
    item_type: Optional[Any] = None,
    allow_existing: bool = False,
) -> str:
    """Crea una nuova cartella (eventualmente sotto un percorso/id/nome indicato)."""
    if not new_folder_name or not new_folder_name.strip():
        return "Errore: specifica un nome valido per la nuova cartella."

    allow_existing_bool = coerce_bool(allow_existing)
    logger.info(
        "create_folder chiamato (nome=%s parent_id=%s parent_path=%s parent_name=%s tipo=%s allow_existing=%s).",
        new_folder_name,
        parent_folder_id,
        parent_folder_path,
        parent_folder_name,
        item_type,
        allow_existing_bool,
    )

    try:
        _, namespace = _connect()
        parent, attempts = folder_service.resolve_folder(
            namespace,
            folder_id=parent_folder_id,
            folder_path=parent_folder_path,
            folder_name=parent_folder_name,
        )
        if not parent:
            detail = "; ".join(attempts) if attempts else "cartella padre non trovata."
            return f"Errore: impossibile individuare la cartella padre ({detail})."

        try:
            _, message = folder_service.create_folder(
                parent,
                new_folder_name=new_folder_name,
                item_type=item_type,
                allow_existing=allow_existing_bool,
            )
            return message
        except ValueError as exc:
            return f"Errore: {exc}"
        except RuntimeError as exc:
            return f"Errore: {exc}"
    except Exception as exc:
        logger.exception("Errore durante create_folder.")
        return f"Errore durante la creazione della cartella: {exc}"


@mcp_tool()
@feature_gate(group="folders")
def rename_folder(
    folder_id: Optional[str] = None,
    new_name: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
) -> str:
    """Rinomina una cartella esistente risolta per id/path/nome."""
    if not new_name or not new_name.strip():
        return "Errore: specifica un nuovo nome valido per la cartella."

    logger.info(
        "rename_folder chiamato (id=%s path=%s nome=%s nuovo_nome=%s).",
        folder_id,
        folder_path,
        folder_name,
        new_name,
    )

    try:
        _, namespace = _connect()
        target, attempts = folder_service.resolve_folder(
            namespace,
            folder_id=folder_id,
            folder_path=folder_path,
            folder_name=folder_name,
        )
        if not target:
            detail = "; ".join(attempts) if attempts else "cartella non trovata."
            return f"Errore: impossibile individuare la cartella da rinominare ({detail})."

        try:
            folder_service.rename_folder(target, new_name)
        except ValueError as exc:
            return f"Errore: {exc}"
        except RuntimeError as exc:
            return f"Errore: {exc}"

        path_display = safe_folder_path(target) or new_name.strip()
        return f"Cartella rinominata in '{new_name.strip()}' (percorso attuale: {path_display})."
    except Exception as exc:
        logger.exception("Errore durante rename_folder.")
        return f"Errore durante la rinomina della cartella: {exc}"


@mcp_tool()
@feature_gate(group="folders")
def delete_folder(
    folder_id: Optional[str] = None,
    folder_path: Optional[str] = None,
    folder_name: Optional[str] = None,
    confirm: bool = False,
) -> str:
    """Elimina una cartella (richiede `confirm=True` per sicurezza)."""
    if not coerce_bool(confirm):
        return "Conferma mancante: imposta confirm=True per procedere con l'eliminazione della cartella."

    logger.info(
        "delete_folder chiamato (id=%s path=%s nome=%s).",
        folder_id,
        folder_path,
        folder_name,
    )

    try:
        _, namespace = _connect()
        target, attempts = folder_service.resolve_folder(
            namespace,
            folder_id=folder_id,
            folder_path=folder_path,
            folder_name=folder_name,
        )
        if not target:
            detail = "; ".join(attempts) if attempts else "cartella non trovata."
            return f"Errore: impossibile individuare la cartella da eliminare ({detail})."

        path_display = safe_folder_path(target) or getattr(target, "Name", "(sconosciuta)")
        try:
            folder_service.delete_folder(target)
        except RuntimeError as exc:
            return f"Errore: {exc}"

        return (
            f"Cartella eliminata: {path_display}. (Se previsto, Outlook l'ha spostata in Posta eliminata.)"
        )
    except Exception as exc:
        logger.exception("Errore durante delete_folder.")
        return f"Errore durante l'eliminazione della cartella: {exc}"


