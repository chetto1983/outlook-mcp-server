"""Manual script to inspect Outlook calendar items directly via COM.

Run with: python tests/manual_calendar_probe.py
"""

import datetime
try:
    from tabulate import tabulate  # type: ignore
except ImportError:
    tabulate = None

try:
    from win32com.client import Dispatch  # type: ignore
except ImportError as exc:
    raise SystemExit("win32com è necessario. Installa pywin32.") from exc


OUTLOOK_FORMAT = "%m/%d/%Y %H:%M"
LOOKAHEAD_DAYS = 14


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode("cp1252", errors="replace").decode("cp1252"))


def main():
    outlook = Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    appointments = namespace.GetDefaultFolder(9).Items
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = True

    begin = datetime.datetime.now() - datetime.timedelta(days=1)
    end = begin + datetime.timedelta(days=LOOKAHEAD_DAYS)

    restriction = (
        f"[End] >= '{begin.strftime('%m/%d/%Y %I:%M %p')}' AND "
        f"[Start] <= '{end.strftime('%m/%d/%Y %I:%M %p')}'"
    )
    safe_print(f"Restriction usata: {restriction}")
    restricted_items = appointments.Restrict(restriction)
    restricted_items.Sort("[Start]")
    restricted_items.IncludeRecurrences = True
    try:
        safe_print(f"Count dopo Restrict: {restricted_items.Count}")
    except Exception:
        safe_print("Impossibile leggere Count dopo Restrict")

    try:
        probe_first = restricted_items.GetFirst()
        safe_print(f"Primo elemento via GetFirst: {getattr(probe_first, 'Subject', None)}")
    except Exception:
        safe_print("GetFirst ha generato eccezione")

    rows = []
    for item in restricted_items:
        try:
            rows.append(
                [
                    getattr(item, "Subject", ""),
                    getattr(item, "Organizer", ""),
                    item.Start.Format(OUTLOOK_FORMAT) if hasattr(item, "Start") else "",
                    getattr(item, "Duration", ""),
                ]
            )
        except Exception:
            continue

    if rows:
        if tabulate:
            headers = ["Titolo", "Organizzatore", "Inizio", "Durata (minuti)"]
            print(tabulate(rows, headers=headers))
        else:
            for row in rows:
                safe_print(f"{row[0]} | {row[1]} | {row[2]} | {row[3]}")
    else:
        safe_print("Nessun evento trovato.")


if __name__ == "__main__":
    main()
