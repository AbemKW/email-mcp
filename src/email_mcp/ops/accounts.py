"""Account operations — list and validate configured Outlook accounts.

Thin wrappers over :class:`email_mcp.outlook.session.OutlookSession`. The session
owns all COM; this module only delegates and applies the pure validation logic
ported from the old JS ``validateAccountSelection``.
"""

from __future__ import annotations

from typing import Any

from email_mcp.outlook.session import validate_account_selection


def list_accounts(session: Any) -> list[dict[str, str]]:
    """Return configured Outlook accounts as ``[{name, entry_id}]``.

    Delegates straight to :meth:`OutlookSession.list_accounts`, which reads the
    store roots (mirrors ``fetchOutlookAccounts`` in index.js).
    """
    return session.list_accounts()


def resolve_validated_account(session: Any, account: str) -> dict[str, str]:
    """Fetch configured accounts and return the one matching ``account``.

    Case-insensitively substring-matches ``account`` against configured account
    names via :func:`validate_account_selection`, which raises
    :class:`email_mcp.errors.AccountNotFound` (listing available accounts) when
    ``account`` is empty or matches nothing.
    """
    accounts = session.list_accounts()
    return validate_account_selection(account, accounts)
