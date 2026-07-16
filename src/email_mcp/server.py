"""FastMCP server exposing the 10 Outlook email tools over stdio.

This is a thin pass-through layer. It owns a single :class:`OutlookWorker` (an
STA thread that holds the live Outlook COM handle) and marshals every tool call
onto it via :func:`_run`. All real logic — clamps, validation, DASL compilation,
COM manipulation — lives in the ``email_mcp.ops`` modules; this module only maps
MCP tool signatures onto those ops calls and serializes the result to JSON.

Tool names, descriptions, parameter names, and defaults mirror the old Node
implementation's ``TOOLS`` table (index.js) exactly. User-facing failures raised
by the ops layer are :class:`email_mcp.errors.EmailMcpError` instances; those are
caught and returned as ``{"error": <message>}`` JSON (matching the JS behavior of
emitting error JSON for not-found cases). Any other exception propagates so
FastMCP surfaces it as an internal error.
"""

from __future__ import annotations

import asyncio
import json
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from email_mcp.errors import EmailMcpError
from email_mcp.ops import accounts as accounts_ops
from email_mcp.ops import attachments as attachments_ops
from email_mcp.ops import calendar as calendar_ops
from email_mcp.ops import messages as messages_ops
from email_mcp.ops import sync as sync_ops
from email_mcp.outlook.session import OutlookSession
from email_mcp.outlook.worker import OutlookWorker

mcp = FastMCP("email")

# Module-level worker; NOT started at import — main() owns its lifecycle.
_worker = OutlookWorker()


async def _run(fn: Callable[[OutlookSession], Any]) -> Any:
    """Marshal ``fn(session)`` onto the STA worker thread and await its result."""
    return await asyncio.wrap_future(_worker.submit(fn))


async def _call(fn: Callable[[OutlookSession], Any]) -> str:
    """Run an ops closure on the worker and serialize the result to JSON.

    :class:`EmailMcpError` (the user-facing failure modes) is caught and rendered
    as ``{"error": <message>}`` to mirror the JS server's not-found handling. Any
    other exception propagates for FastMCP to surface as an internal error.
    """
    try:
        result = await _run(fn)
    except EmailMcpError as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)
    return json.dumps(result, default=str, ensure_ascii=False)


# ---------- tool descriptions (verbatim from index.js TOOLS 789-939) ----------

_QUERY_EMAILS_DESC = """Query emails using a MongoDB-style filter tree. Searches across ALL mail folders by default (Inbox, Sent, Archive, custom, etc.).

FIELDS (queryable): subject, from, from_name, to, cc, body (slow, requires allow_slow=true), received, sent, unread, has_attachments, importance, size.

OPERATORS: $eq $ne $in $nin $contains $not_contains $starts_with $ends_with $gte $lte $gt $lt $exists.
COMBINATORS: $and $or $not.

DATES: ISO format ("2026-01-25" or "2026-01-25T14:30").

DEFAULT FIELDS in response: small queries (limit <= 20) include entry_id; larger scans omit it for token efficiency. Override with the 'fields' param.

PATTERN: For broad searches, scan first with limit > 20 (no entry_ids), then re-query narrowly to get entry_ids for the specific items you want to act on (read/reply/mark).

EXAMPLE filter:
  { "$and": [
      { "from": { "$in": ["a@x.com","b@x.com"] } },
      { "subject": { "$contains": "report" } },
      { "received": { "$gte": "2026-01-01" } }
  ]}"""

_SEND_EMAIL_DESC = (
    "Send a new email. The `account` parameter is REQUIRED — pass a substring of the account name "
    "(e.g. SMTP address) to disambiguate. This prevents accidentally sending personal mail from a work "
    "account (or vice versa) when multiple accounts are configured in Outlook. Use the optional `send_as` "
    'parameter to send AS a different mailbox (true Send As — recipient sees the target address as the '
    'sender, no "on behalf of"). Requires Exchange "Send As" permission granted server-side on the target '
    "mailbox."
)

_REPLY_EMAIL_DESC = (
    "Reply to an email. The `account` parameter is REQUIRED to prevent accidental cross-account replies. "
    "Use the optional `send_as` parameter for true Send As (see send_email for the full constraints)."
)

_FORWARD_EMAIL_DESC = (
    "Forward an email. Preserves the original subject (FW:), quoted body, and attachments. The `account` "
    "parameter is REQUIRED to prevent accidental cross-account forwards. Use the optional `send_as` "
    "parameter for true Send As (see send_email for the full constraints)."
)

_DOWNLOAD_ATTACHMENTS_DESC = (
    "Save real (non-inline) attachments from an email to a descriptive subfolder under the user's "
    "Downloads folder (~/Downloads/email-attachments/YYYY-MM-DD_sender_subject/). Returns the folder path "
    "and the list of saved files. Inline images (signatures, embedded screenshots) are filtered out by "
    "default — set include_inline=true to save them too."
)

_FORCE_SYNC_DESC = (
    "Trigger Outlook Send/Receive and block until all configured sync groups finish (or timeout). Useful "
    "before querying to make sure recent server-side mail is local. Returns per-group elapsed time and a "
    "timed_out flag."
)


# ---------- tools ----------


@mcp.tool(description="List all email accounts configured in Outlook")
async def list_accounts() -> str:
    return await _call(lambda s: accounts_ops.list_accounts(s))


@mcp.tool(description=_QUERY_EMAILS_DESC)
async def query_emails(
    filter: dict = {},
    fields: list[str] | None = None,
    account: str = "",
    limit: int = 20,
    offset: int = 0,
    order_by: str = "received_desc",
    allow_slow: bool = False,
) -> str:
    f = dict(filter)
    return await _call(
        lambda s: messages_ops.query_emails(
            s,
            filter=f,
            fields=fields,
            account=account,
            limit=limit,
            offset=offset,
            order_by=order_by,
            allow_slow=allow_slow,
        )
    )


@mcp.tool(description="Read the full content of an email by its EntryID")
async def read_email(entry_id: str) -> str:
    return await _call(lambda s: messages_ops.read_email(s, entry_id))


@mcp.tool(description=_SEND_EMAIL_DESC)
async def send_email(
    to: str,
    subject: str,
    body: str,
    account: str,
    cc: str = "",
    attachments: list[str] | None = None,
    send_as: str = "",
) -> str:
    return await _call(
        lambda s: messages_ops.send_email(
            s,
            to=to,
            subject=subject,
            body=body,
            account=account,
            cc=cc,
            attachments=attachments,
            send_as=send_as,
        )
    )


@mcp.tool(description=_REPLY_EMAIL_DESC)
async def reply_email(
    entry_id: str,
    body: str,
    account: str,
    reply_all: bool = False,
    send_as: str = "",
) -> str:
    return await _call(
        lambda s: messages_ops.reply_email(
            s,
            entry_id=entry_id,
            body=body,
            account=account,
            reply_all=reply_all,
            send_as=send_as,
        )
    )


@mcp.tool(description=_FORWARD_EMAIL_DESC)
async def forward_email(
    entry_id: str,
    to: str,
    account: str,
    cc: str = "",
    body: str = "",
    send_as: str = "",
) -> str:
    return await _call(
        lambda s: messages_ops.forward_email(
            s,
            entry_id=entry_id,
            to=to,
            account=account,
            cc=cc,
            body=body,
            send_as=send_as,
        )
    )


@mcp.tool(description=_DOWNLOAD_ATTACHMENTS_DESC)
async def download_attachments(entry_id: str, include_inline: bool = False) -> str:
    return await _call(
        lambda s: attachments_ops.download_attachments(s, entry_id, include_inline=include_inline)
    )


@mcp.tool(description="Mark an email as read or unread")
async def mark_as_read(entry_id: str, read: bool = True) -> str:
    return await _call(lambda s: messages_ops.mark_as_read(s, entry_id, read=read))


@mcp.tool(description=_FORCE_SYNC_DESC)
async def force_sync(timeout_sec: int = 60, account: str = "") -> str:
    return await _call(lambda s: sync_ops.force_sync(s, timeout_sec=timeout_sec, account=account))


@mcp.tool(description="List upcoming calendar events")
async def list_calendar(days: int = 7, count: int = 20, account: str = "") -> str:
    return await _call(
        lambda s: calendar_ops.list_calendar(s, days=days, count=count, account=account)
    )


def main() -> None:
    """Start the STA worker, run the MCP server over stdio, then tear the worker down."""
    _worker.start()
    try:
        mcp.run()
    finally:
        _worker.shutdown()


if __name__ == "__main__":
    main()
