"""Inspect the type of Outlook appointment Start/End values."""

from win32com.client import Dispatch  # type: ignore


def main():
    outlook = Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    items = namespace.GetDefaultFolder(9).Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    for item in items:
        start = getattr(item, "Start", None)
        end = getattr(item, "End", None)
        print("type(start):", type(start))
        print("repr(start):", repr(start))
        if hasattr(start, "Format"):
            print("format(start):", start.Format("%Y-%m-%d %H:%M"))
        print("str(start):", str(start))
        try:
            print("float(start):", float(start))
        except Exception as exc:
            print("float(start) error:", exc)
        print("type(end):", type(end))
        print("repr(end):", repr(end))
        if hasattr(end, "Format"):
            print("format(end):", end.Format("%Y-%m-%d %H:%M"))
        print("str(end):", str(end))
        try:
            print("float(end):", float(end))
        except Exception as exc:
            print("float(end) error:", exc)
        break


if __name__ == "__main__":
    main()
