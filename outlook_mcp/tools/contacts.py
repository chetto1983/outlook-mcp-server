from __future__ import annotations

from typing import Optional, List

from ..features import feature_gate
from outlook_mcp.toolkit import mcp_tool
from outlook_mcp import logger

def _connect():
    from outlook_mcp import connect_to_outlook

    return connect_to_outlook()


@mcp_tool()
@feature_gate(group="contacts")
def search_contacts(
    search_term: Optional[str] = None,
    max_results: int = 50,
) -> str:
    """Ricerca contatti di Outlook, con filtro opzionale."""
    try:
        if not isinstance(max_results, int) or max_results < 1 or max_results > 200:
            return "Errore: 'max_results' deve essere un intero tra 1 e 200."

        if search_term is not None:
            search_display = str(search_term).strip()
            normalized_term = search_display.lower()
        else:
            search_display = None
            normalized_term = ""

        logger.info(
            "search_contacts chiamato (termine=%s max_results=%s).",
            search_display,
            max_results,
        )

        _, namespace = _connect()
        try:
            contacts_folder = namespace.GetDefaultFolder(10)  # Contacts default folder
        except Exception as exc:
            logger.exception("Impossibile accedere alla cartella Contatti.")
            return f"Errore: impossibile accedere alla cartella dei contatti ({exc})."

        items = getattr(contacts_folder, "Items", None)
        if not items:
            return "Nessun contatto disponibile."

        matches: List[dict] = []
        total_count = getattr(items, "Count", None)

        # Iteration strategy: prefer direct indexing if available
        if hasattr(items, "Count") and hasattr(items, "__call__"):
            for index in range(1, items.Count + 1):
                contact = items(index)
                if not contact:
                    continue

                name_candidates = [
                    getattr(contact, "FullName", None),
                    getattr(contact, "FileAs", None),
                    getattr(contact, "CompanyName", None),
                ]
                display_name = next((value for value in name_candidates if value), "Senza nome")

                email_candidates = [
                    getattr(contact, "Email1Address", None),
                    getattr(contact, "Email2Address", None),
                    getattr(contact, "Email3Address", None),
                ]
                primary_email = next((value for value in email_candidates if value), "")

                phone_candidates = [
                    getattr(contact, "MobileTelephoneNumber", None),
                    getattr(contact, "BusinessTelephoneNumber", None),
                    getattr(contact, "HomeTelephoneNumber", None),
                    getattr(contact, "PrimaryTelephoneNumber", None),
                ]
                phone_number = next((value for value in phone_candidates if value), "")

                company = getattr(contact, "CompanyName", "") or ""
                categories = getattr(contact, "Categories", "") or ""

                if normalized_term:
                    haystack_parts = [
                        str(display_name),
                        str(primary_email or ""),
                        company,
                        str(phone_number or ""),
                        categories,
                    ]
                    haystack = " ".join(part.lower() for part in haystack_parts if part)
                    if normalized_term not in haystack:
                        continue

                matches.append(
                    {
                        "name": str(display_name),
                        "email": str(primary_email).strip() if primary_email else "",
                        "company": company.strip(),
                        "phone": str(phone_number).strip() if phone_number else "",
                    }
                )

                if len(matches) >= max_results:
                    break

        else:
            # Fallback: ad-hoc iterable patterns (GetFirst/GetNext) or attribute list
            seen = 0
            first = getattr(items, "GetFirst", None)
            getnext = getattr(items, "GetNext", None)
            current = first() if callable(first) else None
            if current is None and hasattr(items, "_items"):
                # Some test doubles expose a simple list
                for contact in getattr(items, "_items"):
                    current = contact
                    break

            while current is not None:
                seen += 1
                name = getattr(current, "FullName", None) or getattr(current, "FileAs", None) or "Senza nome"
                email = getattr(current, "Email1Address", None) or ""
                if not normalized_term or normalized_term in f"{name} {email}".lower():
                    matches.append(
                        {
                            "name": str(name),
                            "email": str(email).strip(),
                            "company": (getattr(current, "CompanyName", "") or "").strip(),
                            "phone": (getattr(current, "MobileTelephoneNumber", "") or "").strip(),
                        }
                    )
                    if len(matches) >= max_results:
                        break
                current = getnext() if callable(getnext) else None

        if not matches:
            return "Nessun contatto corrisponde ai criteri richiesti."

        header_suffix = ""
        if total_count is not None:
            header_suffix = f" su {total_count}"
        lines = [f"Trovati {len(matches)} contatti{header_suffix}.", ""]
        for index, info in enumerate(matches, 1):
            parts = [f"{index}. {info['name']}"]
            if info["email"]:
                parts.append(f"<{info['email']}>")
            details: List[str] = []
            if info["company"]:
                details.append(info["company"])
            if info["phone"]:
                details.append(info["phone"])
            if details:
                parts.append(f"({'; '.join(details)})")
            lines.append(" ".join(parts))

        if normalized_term:
            lines.append("")
            lines.append(f"Filtro applicato: '{search_display}'.")

        return "\n".join(lines)
    except Exception as exc:
        logger.exception("Errore durante search_contacts.")
        return f"Errore durante la ricerca dei contatti: {exc}"
