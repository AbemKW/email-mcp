"""Typed exceptions for email-mcp.

Every user-facing failure mode is one of these. CLI renders them to stderr with a
non-zero exit; the MCP server wraps them as tool errors carrying the same message.
Nothing below imports COM — these are safe to raise from pure code and tests.
"""

from __future__ import annotations


class EmailMcpError(Exception):
    """Base for all email-mcp errors. `message` is safe to show the user."""


class NoCOMAvailable(EmailMcpError):
    """Outlook COM could not be reached (e.g. only the 'new' Outlook is installed).

    The 'new' Outlook for Windows does not expose COM automation. Classic Outlook
    desktop must be installed and configured with at least one account.
    """


class OutlookNotRunning(EmailMcpError):
    """Outlook is installed but a live Application/MAPI handle could not be created."""


class AccountNotFound(EmailMcpError):
    """The supplied `account` substring matched no configured Outlook account.

    Constructed with the list of available account names so the message can guide
    the caller, mirroring the JS `validateAccountSelection` behavior.
    """

    def __init__(self, supplied: str | None, available: list[str]):
        self.supplied = supplied
        self.available = available
        avail = ", ".join(available) if available else "(none configured)"
        if not supplied:
            super().__init__(f"account is required. Available accounts: {avail}")
        else:
            super().__init__(f"Unknown account: '{supplied}'. Available accounts: {avail}")


class InvalidFilter(EmailMcpError):
    """The query_emails filter tree was malformed (bad field, operator, date, shape)."""


class InvalidAttachmentPath(EmailMcpError):
    """An attachment path was not a non-empty, absolute path to an existing file."""


class EmailNotFound(EmailMcpError):
    """No item resolved for the given EntryID."""


class SendAsUnresolved(EmailMcpError):
    """The send_as address could not be resolved by Exchange (typically not same-tenant)."""
