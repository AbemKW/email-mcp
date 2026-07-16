"""Folder-walking helpers over live COM folder objects.

Pure with respect to the session — every function takes COM folder objects (or a
session) and returns COM folder objects. Ported from the folder logic in
``buildQueryScript`` and the calendar resolution in the ``list_calendar`` handler.
"""

from __future__ import annotations

from typing import Any, Iterator

from email_mcp.outlook.session import (
    OL_CALENDAR_DEFAULT_TYPE,
    OL_FOLDER_CALENDAR,
    OL_MAIL_ITEM,
    OutlookSession,
)


def walk_mail_folders(parent: Any) -> Iterator[Any]:
    """Yield every descendant folder of ``parent`` whose DefaultItemType is mail.

    Depth-first, defensive against folders that raise on access (mirrors the
    ``Get-MailFolders`` recursion which swallows per-folder errors).
    """
    try:
        subfolders = parent.Folders
    except Exception:
        return
    for f in subfolders:
        try:
            if f.DefaultItemType == OL_MAIL_ITEM:
                yield f
        except Exception:
            pass
        yield from walk_mail_folders(f)


def select_stores(session: OutlookSession, account_substring: str = "") -> list[Any]:
    """Store roots to search: substring-matched by name, else all; never empty.

    Mirrors buildQueryScript: if a filter is given, keep stores whose Name matches;
    if nothing matches (or no filter), fall back to all stores, then to the first.
    """
    roots = session.store_roots()
    if account_substring:
        needle = account_substring.lower()
        matched = [s for s in roots if needle in str(getattr(s, "Name", "")).lower()]
        if matched:
            return matched
    if roots:
        return roots
    # Absolute fallback: first store via 1-based COM index.
    return [session.ns.Folders.Item(1)]


def all_mail_folders(session: OutlookSession, account_substring: str = "") -> list[Any]:
    """Every mail folder across the selected stores."""
    folders: list[Any] = []
    for store in select_stores(session, account_substring):
        folders.extend(walk_mail_folders(store))
    return folders


def calendar_folder(session: OutlookSession, account_substring: str = "") -> Any:
    """Resolve the calendar folder to list events from.

    With an account substring: find that store, then its first folder with the
    calendar DefaultItemType, else the store's default calendar. Without: the
    namespace default calendar (``GetDefaultFolder(9)``).
    """
    if account_substring:
        needle = account_substring.lower()
        store = next(
            (s for s in session.store_roots() if needle in str(getattr(s, "Name", "")).lower()),
            None,
        )
        if store is None:
            store = session.ns.Folders.Item(1)
        for f in store.Folders:
            try:
                if f.DefaultItemType == OL_CALENDAR_DEFAULT_TYPE:
                    return f
            except Exception:
                continue
    return session.ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
