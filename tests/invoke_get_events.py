"""Invoke get_events_from_folder against live Outlook data."""

import sys
from pathlib import Path

from win32com.client import Dispatch  # type: ignore

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from outlook_mcp.services.calendar import get_events_from_folder

def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode("cp1252", errors="replace").decode("cp1252"))


def main():
    outlook = Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folder = namespace.GetDefaultFolder(9)
    print("Folder:", folder.Name)
    events = get_events_from_folder(folder, days=14)
    print("Events returned:", len(events))
    for event in events:
        safe_print(f"- {event.get('subject')} {event.get('start_time')}")


if __name__ == "__main__":
    main()
