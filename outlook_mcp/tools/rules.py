"""MCP tools for managing Outlook rules (automatic email filtering and actions)."""

from __future__ import annotations

from typing import Any, Optional, List, Dict

from ..features import feature_gate
from outlook_mcp.toolkit import mcp_tool

from outlook_mcp import logger
from outlook_mcp.utils import coerce_bool
from outlook_mcp import folders as folder_service


def _connect():
    from outlook_mcp import connect_to_outlook
    return connect_to_outlook()


@mcp_tool()
@feature_gate(group="rules")
def list_rules() -> str:
    """Elenca tutte le regole di Outlook configurate."""
    try:
        logger.info("list_rules chiamato")
        _, namespace = _connect()

        # Get default store (mailbox)
        try:
            default_store = namespace.DefaultStore
            rules = default_store.GetRules()
        except Exception as exc:
            logger.exception("Impossibile recuperare le regole dal default store.")
            return f"Errore: impossibile accedere alle regole ({exc})"

        if rules.Count == 0:
            return "Nessuna regola configurata in Outlook."

        lines = [
            f"Regole di Outlook configurate ({rules.Count} totali):",
            "",
        ]

        for idx in range(1, rules.Count + 1):
            rule = rules.Item(idx)
            try:
                name = getattr(rule, "Name", f"Regola {idx}")
                enabled = getattr(rule, "Enabled", False)
                is_local = getattr(rule, "IsLocalRule", False)
                execution_order = getattr(rule, "ExecutionOrder", idx)

                status = "✓ Attiva" if enabled else "✗ Disattivata"
                location = "Locale" if is_local else "Server"

                lines.append(f"{idx}. {name}")
                lines.append(f"   Stato: {status} | Tipo: {location} | Ordine: {execution_order}")

                # Try to get basic condition info
                try:
                    conditions = rule.Conditions
                    condition_list = []

                    if hasattr(conditions, "From") and conditions.From.Enabled:
                        recipients = conditions.From.Recipients
                        if recipients and recipients.Count > 0:
                            condition_list.append(f"Da: {recipients.Item(1).Name}")

                    if hasattr(conditions, "Subject") and conditions.Subject.Enabled:
                        text = ", ".join(conditions.Subject.Text) if hasattr(conditions.Subject.Text, "__iter__") else str(conditions.Subject.Text)
                        condition_list.append(f"Oggetto contiene: {text}")

                    if hasattr(conditions, "SentTo") and conditions.SentTo.Enabled:
                        recipients = conditions.SentTo.Recipients
                        if recipients and recipients.Count > 0:
                            condition_list.append(f"Inviato a: {recipients.Item(1).Name}")

                    if condition_list:
                        lines.append(f"   Condizioni: {' | '.join(condition_list)}")
                except Exception:
                    pass

                # Try to get basic action info
                try:
                    actions = rule.Actions
                    action_list = []

                    if hasattr(actions, "MoveToFolder") and actions.MoveToFolder.Enabled:
                        try:
                            folder = actions.MoveToFolder.Folder
                            folder_name = getattr(folder, "Name", "cartella sconosciuta")
                            action_list.append(f"Sposta in: {folder_name}")
                        except Exception:
                            action_list.append("Sposta in cartella")

                    if hasattr(actions, "MarkAsRead") and actions.MarkAsRead.Enabled:
                        action_list.append("Segna come letto")

                    if hasattr(actions, "Delete") and actions.Delete.Enabled:
                        action_list.append("Elimina")

                    if hasattr(actions, "AssignToCategory") and actions.AssignToCategory.Enabled:
                        try:
                            categories = ", ".join(actions.AssignToCategory.Categories) if hasattr(actions.AssignToCategory.Categories, "__iter__") else str(actions.AssignToCategory.Categories)
                            action_list.append(f"Categoria: {categories}")
                        except Exception:
                            action_list.append("Assegna categoria")

                    if action_list:
                        lines.append(f"   Azioni: {' | '.join(action_list)}")
                except Exception:
                    pass

                lines.append("")

            except Exception as exc:
                logger.warning("Errore nel processamento della regola #%s: %s", idx, exc)
                lines.append(f"{idx}. (Errore nel recupero dei dettagli)")
                lines.append("")

        return "\n".join(lines)

    except Exception as exc:
        logger.exception("Errore durante list_rules.")
        return f"Errore durante il recupero delle regole: {exc}"


@mcp_tool()
@feature_gate(group="rules")
def get_rule_details(rule_name: str) -> str:
    """Mostra i dettagli completi di una regola specifica."""
    try:
        if not rule_name.strip():
            return "Errore: specifica il nome della regola"

        logger.info("get_rule_details chiamato per regola '%s'", rule_name)
        _, namespace = _connect()

        # Get rules
        try:
            default_store = namespace.DefaultStore
            rules = default_store.GetRules()
        except Exception as exc:
            logger.exception("Impossibile recuperare le regole.")
            return f"Errore: impossibile accedere alle regole ({exc})"

        # Find rule by name
        rule = None
        for idx in range(1, rules.Count + 1):
            r = rules.Item(idx)
            if r.Name.lower() == rule_name.lower():
                rule = r
                break

        if not rule:
            return f"Errore: regola '{rule_name}' non trovata"

        # Build detailed report
        lines = [
            f"Dettagli regola: {rule.Name}",
            "",
            f"Stato: {'Attiva' if rule.Enabled else 'Disattivata'}",
            f"Tipo: {'Locale' if rule.IsLocalRule else 'Server'}",
            f"Ordine esecuzione: {rule.ExecutionOrder}",
            "",
            "CONDIZIONI:",
        ]

        # Detailed conditions
        try:
            conditions = rule.Conditions
            condition_count = 0

            if hasattr(conditions, "From") and conditions.From.Enabled:
                recipients = conditions.From.Recipients
                if recipients and recipients.Count > 0:
                    names = [recipients.Item(i).Name for i in range(1, recipients.Count + 1)]
                    lines.append(f"  - Da: {', '.join(names)}")
                    condition_count += 1

            if hasattr(conditions, "Subject") and conditions.Subject.Enabled:
                text = conditions.Subject.Text
                if hasattr(text, "__iter__") and not isinstance(text, str):
                    keywords = ", ".join(str(t) for t in text)
                else:
                    keywords = str(text)
                lines.append(f"  - Oggetto contiene: {keywords}")
                condition_count += 1

            if hasattr(conditions, "SentTo") and conditions.SentTo.Enabled:
                recipients = conditions.SentTo.Recipients
                if recipients and recipients.Count > 0:
                    names = [recipients.Item(i).Name for i in range(1, recipients.Count + 1)]
                    lines.append(f"  - Inviato a: {', '.join(names)}")
                    condition_count += 1

            if hasattr(conditions, "Body") and conditions.Body.Enabled:
                text = conditions.Body.Text
                if hasattr(text, "__iter__") and not isinstance(text, str):
                    keywords = ", ".join(str(t) for t in text)
                else:
                    keywords = str(text)
                lines.append(f"  - Corpo contiene: {keywords}")
                condition_count += 1

            if hasattr(conditions, "MessageSize") and conditions.MessageSize.Enabled:
                lines.append("  - Dimensione messaggio (filtro attivo)")
                condition_count += 1

            if hasattr(conditions, "Importance") and conditions.Importance.Enabled:
                importance_map = {0: "Bassa", 1: "Normale", 2: "Alta"}
                imp = importance_map.get(conditions.Importance.Importance, "Sconosciuta")
                lines.append(f"  - Importanza: {imp}")
                condition_count += 1

            if condition_count == 0:
                lines.append("  (Nessuna condizione configurata)")

        except Exception as exc:
            lines.append(f"  (Errore nel recupero delle condizioni: {exc})")

        lines.append("")
        lines.append("AZIONI:")

        # Detailed actions
        try:
            actions = rule.Actions
            action_count = 0

            if hasattr(actions, "MoveToFolder") and actions.MoveToFolder.Enabled:
                try:
                    folder = actions.MoveToFolder.Folder
                    folder_name = getattr(folder, "Name", "cartella sconosciuta")
                    folder_path = getattr(folder, "FolderPath", "")
                    lines.append(f"  - Sposta in cartella: {folder_name} ({folder_path})")
                    action_count += 1
                except Exception as exc:
                    lines.append(f"  - Sposta in cartella (errore: {exc})")
                    action_count += 1

            if hasattr(actions, "CopyToFolder") and actions.CopyToFolder.Enabled:
                try:
                    folder = actions.CopyToFolder.Folder
                    folder_name = getattr(folder, "Name", "cartella sconosciuta")
                    lines.append(f"  - Copia in cartella: {folder_name}")
                    action_count += 1
                except Exception:
                    lines.append("  - Copia in cartella")
                    action_count += 1

            if hasattr(actions, "MarkAsRead") and actions.MarkAsRead.Enabled:
                lines.append("  - Segna come letto")
                action_count += 1

            if hasattr(actions, "Delete") and actions.Delete.Enabled:
                lines.append("  - Elimina")
                action_count += 1

            if hasattr(actions, "AssignToCategory") and actions.AssignToCategory.Enabled:
                try:
                    categories = actions.AssignToCategory.Categories
                    if hasattr(categories, "__iter__") and not isinstance(categories, str):
                        cat_list = ", ".join(str(c) for c in categories)
                    else:
                        cat_list = str(categories)
                    lines.append(f"  - Assegna categorie: {cat_list}")
                    action_count += 1
                except Exception:
                    lines.append("  - Assegna categorie")
                    action_count += 1

            if hasattr(actions, "Forward") and actions.Forward.Enabled:
                try:
                    recipients = actions.Forward.Recipients
                    if recipients and recipients.Count > 0:
                        names = [recipients.Item(i).Name for i in range(1, recipients.Count + 1)]
                        lines.append(f"  - Inoltra a: {', '.join(names)}")
                        action_count += 1
                except Exception:
                    lines.append("  - Inoltra")
                    action_count += 1

            if hasattr(actions, "Stop") and actions.Stop.Enabled:
                lines.append("  - Interrompi elaborazione ulteriori regole")
                action_count += 1

            if action_count == 0:
                lines.append("  (Nessuna azione configurata)")

        except Exception as exc:
            lines.append(f"  (Errore nel recupero delle azioni: {exc})")

        return "\n".join(lines)

    except Exception as exc:
        logger.exception("Errore durante get_rule_details per regola '%s'.", rule_name)
        return f"Errore durante il recupero dei dettagli della regola: {exc}"


@mcp_tool()
@feature_gate(group="rules")
def enable_disable_rule(
    rule_name: str,
    enabled: bool = True,
) -> str:
    """Abilita o disabilita una regola esistente."""
    try:
        if not rule_name.strip():
            return "Errore: specifica il nome della regola"

        enabled_bool = coerce_bool(enabled)
        logger.info("enable_disable_rule chiamato per regola '%s' (enabled=%s)", rule_name, enabled_bool)

        _, namespace = _connect()

        # Get rules
        try:
            default_store = namespace.DefaultStore
            rules = default_store.GetRules()
        except Exception as exc:
            logger.exception("Impossibile recuperare le regole.")
            return f"Errore: impossibile accedere alle regole ({exc})"

        # Find and update rule
        rule_found = False
        for idx in range(1, rules.Count + 1):
            rule = rules.Item(idx)
            if rule.Name.lower() == rule_name.lower():
                rule.Enabled = enabled_bool
                rule_found = True
                break

        if not rule_found:
            return f"Errore: regola '{rule_name}' non trovata"

        # Save changes
        try:
            rules.Save()
        except Exception as exc:
            logger.exception("Impossibile salvare le modifiche alle regole.")
            return f"Errore: impossibile salvare le modifiche ({exc})"

        status = "abilitata" if enabled_bool else "disabilitata"
        return f"Regola '{rule_name}' {status} con successo."

    except Exception as exc:
        logger.exception("Errore durante enable_disable_rule per regola '%s'.", rule_name)
        return f"Errore durante l'abilitazione/disabilitazione della regola: {exc}"


@mcp_tool()
@feature_gate(group="rules")
def delete_rule(rule_name: str) -> str:
    """Elimina una regola esistente."""
    try:
        if not rule_name.strip():
            return "Errore: specifica il nome della regola"

        logger.info("delete_rule chiamato per regola '%s'", rule_name)
        _, namespace = _connect()

        # Get rules
        try:
            default_store = namespace.DefaultStore
            rules = default_store.GetRules()
        except Exception as exc:
            logger.exception("Impossibile recuperare le regole.")
            return f"Errore: impossibile accedere alle regole ({exc})"

        # Find and delete rule
        rule_index = None
        for idx in range(1, rules.Count + 1):
            rule = rules.Item(idx)
            if rule.Name.lower() == rule_name.lower():
                rule_index = idx
                break

        if rule_index is None:
            return f"Errore: regola '{rule_name}' non trovata"

        # Remove rule
        try:
            rules.Remove(rule_index)
            rules.Save()
        except Exception as exc:
            logger.exception("Impossibile eliminare la regola.")
            return f"Errore: impossibile eliminare la regola ({exc})"

        return f"Regola '{rule_name}' eliminata con successo."

    except Exception as exc:
        logger.exception("Errore durante delete_rule per regola '%s'.", rule_name)
        return f"Errore durante l'eliminazione della regola: {exc}"


@mcp_tool()
@feature_gate(group="rules")
def create_move_rule(
    rule_name: str,
    from_address: Optional[str] = None,
    subject_contains: Optional[str] = None,
    body_contains: Optional[str] = None,
    target_folder_name: Optional[str] = None,
    target_folder_path: Optional[str] = None,
    mark_as_read: bool = False,
    enabled: bool = True,
    stop_processing: bool = True,
) -> str:
    """Crea una nuova regola per spostare email in una cartella specifica.

    Questa è una funzione semplificata per il caso d'uso più comune.
    Per regole più complesse, usa l'interfaccia di Outlook direttamente.
    """
    try:
        if not rule_name.strip():
            return "Errore: specifica un nome per la regola"

        if not from_address and not subject_contains and not body_contains:
            return "Errore: specifica almeno una condizione (from_address, subject_contains, o body_contains)"

        if not target_folder_name and not target_folder_path:
            return "Errore: specifica la cartella di destinazione (target_folder_name o target_folder_path)"

        enabled_bool = coerce_bool(enabled)
        mark_read_bool = coerce_bool(mark_as_read)
        stop_bool = coerce_bool(stop_processing)

        logger.info(
            "create_move_rule chiamato con nome='%s' da=%s oggetto=%s cartella=%s",
            rule_name,
            from_address,
            subject_contains,
            target_folder_name or target_folder_path,
        )

        _, namespace = _connect()

        # Resolve target folder
        target_folder = None
        if target_folder_path:
            target_folder = folder_service.get_folder_by_path(namespace, target_folder_path)
        elif target_folder_name:
            target_folder = folder_service.get_folder_by_name(namespace, target_folder_name)

        if not target_folder:
            return f"Errore: cartella di destinazione non trovata"

        # Get rules collection
        try:
            default_store = namespace.DefaultStore
            rules = default_store.GetRules()
        except Exception as exc:
            logger.exception("Impossibile recuperare le regole.")
            return f"Errore: impossibile accedere alle regole ({exc})"

        # Create new rule
        try:
            rule = rules.Create(rule_name, 0)  # 0 = olRuleReceive (server-side rule)
            rule.Enabled = enabled_bool
        except Exception as exc:
            logger.exception("Impossibile creare la nuova regola.")
            return f"Errore: impossibile creare la regola ({exc}). Potrebbe già esistere una regola con questo nome."

        # Set conditions
        conditions_set = []
        try:
            conditions = rule.Conditions

            if from_address:
                try:
                    conditions.From.Enabled = True
                    recipient = conditions.From.Recipients.Add(from_address)
                    recipient.Resolve()
                    conditions_set.append(f"Da: {from_address}")
                except Exception as exc:
                    logger.warning("Impossibile impostare condizione 'From': %s", exc)

            if subject_contains:
                try:
                    conditions.Subject.Enabled = True
                    conditions.Subject.Text = [subject_contains]
                    conditions_set.append(f"Oggetto contiene: {subject_contains}")
                except Exception as exc:
                    logger.warning("Impossibile impostare condizione 'Subject': %s", exc)

            if body_contains:
                try:
                    conditions.Body.Enabled = True
                    conditions.Body.Text = [body_contains]
                    conditions_set.append(f"Corpo contiene: {body_contains}")
                except Exception as exc:
                    logger.warning("Impossibile impostare condizione 'Body': %s", exc)

        except Exception as exc:
            logger.exception("Errore nell'impostazione delle condizioni.")
            # Try to clean up
            try:
                rules.Remove(rules.Count)
            except Exception:
                pass
            return f"Errore: impossibile configurare le condizioni della regola ({exc})"

        # Set actions
        actions_set = []
        try:
            actions = rule.Actions

            # Move to folder
            try:
                actions.MoveToFolder.Enabled = True
                actions.MoveToFolder.Folder = target_folder
                folder_name = getattr(target_folder, "Name", "cartella")
                actions_set.append(f"Sposta in: {folder_name}")
            except Exception as exc:
                logger.exception("Impossibile impostare azione 'MoveToFolder'.")
                # Clean up
                try:
                    rules.Remove(rules.Count)
                except Exception:
                    pass
                return f"Errore: impossibile configurare l'azione di spostamento ({exc})"

            # Optional: Mark as read
            if mark_read_bool:
                try:
                    actions.MarkAsRead.Enabled = True
                    actions_set.append("Segna come letto")
                except Exception as exc:
                    logger.warning("Impossibile impostare azione 'MarkAsRead': %s", exc)

            # Optional: Stop processing more rules
            if stop_bool:
                try:
                    actions.Stop.Enabled = True
                    actions_set.append("Interrompi elaborazione")
                except Exception as exc:
                    logger.warning("Impossibile impostare azione 'Stop': %s", exc)

        except Exception as exc:
            logger.exception("Errore nell'impostazione delle azioni.")
            # Clean up
            try:
                rules.Remove(rules.Count)
            except Exception:
                pass
            return f"Errore: impossibile configurare le azioni della regola ({exc})"

        # Save the rule
        try:
            rules.Save()
        except Exception as exc:
            logger.exception("Impossibile salvare la regola.")
            return f"Errore: impossibile salvare la regola ({exc})"

        summary = [
            f"Regola '{rule_name}' creata con successo!",
            "",
            "Condizioni:",
        ]
        summary.extend(f"  - {c}" for c in conditions_set)
        summary.append("")
        summary.append("Azioni:")
        summary.extend(f"  - {a}" for a in actions_set)
        summary.append("")
        summary.append(f"Stato: {'Abilitata' if enabled_bool else 'Disabilitata'}")

        return "\n".join(summary)

    except Exception as exc:
        logger.exception("Errore durante create_move_rule.")
        return f"Errore durante la creazione della regola: {exc}"
