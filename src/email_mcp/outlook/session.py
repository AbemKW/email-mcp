"""OutlookSession — the single COM entry point.

A session owns one ``Outlook.Application`` handle and its MAPI namespace. It does
NOT manage threading: whoever enters the session (the CLI main thread, or the
server's STA worker thread) is responsible for having called ``CoInitialize`` for
that thread. Use it as a context manager, which handles CoInitialize/CoUninitialize
for the entering thread:

    with OutlookSession() as s:
        s.list_accounts()

COM object model reference (same objects the old PowerShell drove):
- ``Application`` -> ``.GetNamespace("MAPI")`` -> namespace
- ``namespace.Folders`` -> store roots (one per configured account)
- ``application.Session.Accounts`` -> send accounts (carry ``.SmtpAddress``)
"""

from __future__ import annotations

from typing import Any, Iterator

from email_mcp.errors import (
    AccountNotFound,
    EmailNotFound,
    NoCOMAvailable,
    OutlookNotRunning,
)

# --- Outlook COM enum constants (avoid a win32com makepy dependency) ---
OL_MAIL_ITEM = 0            # olMailItem (CreateItem arg; also DefaultItemType for mail folders)
OL_CALENDAR_DEFAULT_TYPE = 1  # DefaultItemType for calendar folders
OL_FOLDER_CALENDAR = 9      # olFolderCalendar (GetDefaultFolder arg)

OL_CLASS_MAIL = 43                  # olMail (item.Class for a standard mail item)
OL_CLASS_REPORT = 46                # olReport (NDR, delivery/read receipts)
OL_CLASS_MEETING_REQUEST = 53       # olMeetingRequest (meeting invite)
OL_CLASS_MEETING_CANCELLATION = 54  # olMeetingCancellation
OL_CLASS_MEETING_DECLINE = 55       # olMeetingResponseNegative
OL_CLASS_MEETING_ACCEPT = 56        # olMeetingResponsePositive
OL_CLASS_MEETING_TENTATIVE = 57     # olMeetingResponseTentative
OL_CLASS_SHARING = 181              # olSharing (calendar/folder sharing invite)

OL_MAIL_CLASSES: frozenset[int] = frozenset(
    {
        OL_CLASS_MAIL,
        OL_CLASS_REPORT,
        OL_CLASS_MEETING_REQUEST,
        OL_CLASS_MEETING_CANCELLATION,
        OL_CLASS_MEETING_DECLINE,
        OL_CLASS_MEETING_ACCEPT,
        OL_CLASS_MEETING_TENTATIVE,
        OL_CLASS_SHARING,
    }
)


class OutlookSession:
    """Live COM handle to classic Outlook. Lazy-attaches on first ``.app`` access."""

    def __init__(self) -> None:
        self._app: Any = None
        self._ns: Any = None
        self._co_initialized = False

    # -- lifecycle -------------------------------------------------------
    def open(self) -> "OutlookSession":
        """CoInitialize the current thread and attach to Outlook."""
        import pythoncom

        if not self._co_initialized:
            pythoncom.CoInitialize()
            self._co_initialized = True
        self._attach()
        return self

    def close(self) -> None:
        self._app = None
        self._ns = None
        if self._co_initialized:
            import pythoncom

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._co_initialized = False

    def __enter__(self) -> "OutlookSession":
        return self.open()

    def __exit__(self, *exc: Any) -> None:
        self.close()

    def _attach(self) -> None:
        if self._app is not None:
            return
        import pythoncom
        import win32com.client

        # Prefer an already-running Outlook; else launch one. Mirrors INIT in index.js.
        app = None
        try:
            app = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            pass
        if app is None:
            try:
                app = win32com.client.Dispatch("Outlook.Application")
            except pythoncom.com_error as e:  # type: ignore[attr-defined]
                raise NoCOMAvailable(
                    "Could not reach Outlook via COM. This requires classic Outlook "
                    "desktop (the 'new' Outlook for Windows does not expose COM). "
                    f"Underlying error: {e}"
                ) from e
        self._app = app
        try:
            self._ns = app.GetNamespace("MAPI")
        except Exception as e:
            raise OutlookNotRunning(f"Attached to Outlook but MAPI namespace failed: {e}") from e

    # -- accessors -------------------------------------------------------
    @property
    def app(self) -> Any:
        if self._app is None:
            self._attach()
        return self._app

    @property
    def ns(self) -> Any:
        if self._ns is None:
            self._attach()
        return self._ns

    # -- convenience over the object model -------------------------------
    def store_roots(self) -> list[Any]:
        """Store root folders (``namespace.Folders``), one per configured account."""
        return [f for f in self.ns.Folders]

    def list_accounts(self) -> list[dict[str, str]]:
        """Configured accounts as ``[{name, entry_id}]`` — from store roots.

        This is the list account validation matches against (mirrors
        ``fetchOutlookAccounts`` in index.js), distinct from send accounts.
        """
        out: list[dict[str, str]] = []
        for store in self.ns.Folders:
            try:
                out.append({"name": str(store.Name), "entry_id": str(store.EntryID)})
            except Exception:
                continue
        return out

    def get_item(self, entry_id: str) -> Any:
        """Resolve a MAPI item by EntryID or raise :class:`EmailNotFound`."""
        try:
            item = self.ns.GetItemFromID(entry_id)
        except Exception as e:
            raise EmailNotFound(f"Email not found for EntryID: {entry_id} ({e})") from e
        if item is None:
            raise EmailNotFound(f"Email not found for EntryID: {entry_id}")
        return item

    def send_accounts(self) -> Iterator[Any]:
        """The ``Session.Accounts`` collection (send accounts, carry SmtpAddress)."""
        for acct in self.app.Session.Accounts:
            yield acct

    def find_send_account(self, substring: str) -> Any | None:
        """Return the first send account whose SmtpAddress contains ``substring``
        (case-insensitive), or None. This drives ``SendUsingAccount`` — note it is a
        DIFFERENT list from :meth:`list_accounts` (store names)."""
        needle = (substring or "").lower()
        for acct in self.send_accounts():
            try:
                smtp = str(acct.SmtpAddress or "")
            except Exception:
                smtp = ""
            if needle in smtp.lower():
                return acct
        return None

    def current_user_address(self) -> str:
        try:
            return str(self.app.Session.CurrentUser.Address)
        except Exception:
            return ""


def validate_account_selection(supplied: str | None, accounts: list[dict[str, str]]) -> dict[str, str]:
    """Pure port of ``validateAccountSelection`` (index.js).

    Confirm ``supplied`` is a non-empty string that case-insensitively substring-matches
    at least one configured account ``name``. Returns the matched account dict or raises
    :class:`AccountNotFound` (whose message lists available accounts).
    """
    names = [str(a.get("name", "")) for a in accounts]
    if not isinstance(supplied, str) or not supplied.strip():
        raise AccountNotFound(None, names)
    needle = supplied.lower()
    for a in accounts:
        if needle in str(a.get("name", "")).lower():
            return a
    raise AccountNotFound(supplied, names)
