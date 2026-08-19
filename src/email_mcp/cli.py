"""Typer CLI adapter for email-mcp.

One subcommand per tool (11: list-accounts, query, read, send, draft, reply,
forward, download, mark-read, sync, calendar) plus ``mcp`` (which hands off to
the MCP server). Each
tool command opens a single :class:`email_mcp.outlook.session.OutlookSession`,
delegates to the matching function in :mod:`email_mcp.ops`, and prints the result
as pretty JSON.

This is a *thin* adapter: it parses flags and passes raw values straight through
to the ops layer. All clamps, defaults, and business logic live in the ops
functions — never here.
"""

from __future__ import annotations

import json
from typing import Any, Callable, Optional

import typer

from email_mcp.errors import EmailMcpError
from email_mcp.ops import accounts as ops_accounts
from email_mcp.ops import attachments as ops_attachments
from email_mcp.ops import calendar as ops_calendar
from email_mcp.ops import messages as ops_messages
from email_mcp.ops import sync as ops_sync
from email_mcp.outlook.session import OutlookSession

app = typer.Typer(
    name="email-mcp",
    help="Outlook (classic, Windows COM) email + calendar as a CLI and MCP server.",
    no_args_is_help=True,
    add_completion=False,
)


def _emit(result: Any) -> None:
    """Print ``result`` as pretty JSON on stdout (stable options across commands)."""
    print(json.dumps(result, indent=2, default=str, ensure_ascii=False))


def _run(op: Callable[[OutlookSession], Any]) -> None:
    """Open a session, run ``op`` against it, and emit the result.

    Catches :class:`email_mcp.errors.EmailMcpError`, writes its message to stderr,
    and exits non-zero. Every tool command routes through here so JSON options and
    error handling cannot drift between commands.
    """
    try:
        with OutlookSession() as s:
            result = op(s)
    except EmailMcpError as e:
        typer.echo(str(e), err=True)
        raise typer.Exit(1)
    _emit(result)


@app.command("list-accounts")
def list_accounts_cmd() -> None:
    """List configured Outlook accounts."""
    _run(lambda s: ops_accounts.list_accounts(s))


@app.command("query")
def query_cmd(
    filter_json: str = typer.Option(
        "{}", "--filter", help="Filter tree as a JSON string (default match-all)."
    ),
    account: str = typer.Option("", "--account", help="Account name substring."),
    limit: int = typer.Option(20, "--limit", help="Max results to return."),
    offset: int = typer.Option(0, "--offset", help="Result offset for paging."),
    order_by: str = typer.Option(
        "received_desc", "--order-by", help="Sort order (e.g. received_desc)."
    ),
    allow_slow: bool = typer.Option(
        False, "--allow-slow", help="Permit slow filter fields (e.g. body)."
    ),
    fields: Optional[str] = typer.Option(
        None, "--fields", help="Comma-separated output fields to project."
    ),
) -> None:
    """Query emails with an optional filter."""
    try:
        parsed = json.loads(filter_json)
    except json.JSONDecodeError as e:
        typer.echo(f"Invalid --filter JSON: {e}", err=True)
        raise typer.Exit(1)

    field_list: Optional[list[str]] = None
    if fields is not None:
        field_list = [f.strip() for f in fields.split(",") if f.strip()]

    _run(
        lambda s: ops_messages.query_emails(
            s,
            filter=parsed,
            fields=field_list,
            account=account,
            limit=limit,
            offset=offset,
            order_by=order_by,
            allow_slow=allow_slow,
        )
    )


@app.command("read")
def read_cmd(entry_id: str = typer.Argument(..., help="EntryID of the email.")) -> None:
    """Read a single email by EntryID."""
    _run(lambda s: ops_messages.read_email(s, entry_id))


@app.command("send")
def send_cmd(
    to: str = typer.Option(..., "--to", help="Recipient address(es)."),
    subject: str = typer.Option(..., "--subject", help="Email subject."),
    body: str = typer.Option(..., "--body", help="Plain-text body."),
    account: str = typer.Option(..., "--account", help="Sending account substring."),
    cc: str = typer.Option("", "--cc", help="CC address(es)."),
    send_as: str = typer.Option("", "--send-as", help="SMTP address to send AS."),
    attach: Optional[list[str]] = typer.Option(
        None, "--attach", help="Attachment file path (repeatable)."
    ),
) -> None:
    """Send a new email."""
    _run(
        lambda s: ops_messages.send_email(
            s,
            to=to,
            subject=subject,
            body=body,
            account=account,
            cc=cc,
            attachments=attach,
            send_as=send_as,
        )
    )


@app.command("draft")
def draft_cmd(
    to: str = typer.Option(..., "--to", help="Recipient address(es)."),
    subject: str = typer.Option(..., "--subject", help="Email subject."),
    body: str = typer.Option(..., "--body", help="Plain-text body."),
    account: str = typer.Option(..., "--account", help="Sending account substring."),
    cc: str = typer.Option("", "--cc", help="CC address(es)."),
    attach: Optional[list[str]] = typer.Option(
        None, "--attach", help="Attachment file path (repeatable)."
    ),
) -> None:
    """Save a new email as a draft (does not send)."""
    _run(
        lambda s: ops_messages.draft_email(
            s,
            to=to,
            subject=subject,
            body=body,
            account=account,
            cc=cc,
            attachments=attach,
        )
    )


@app.command("reply")
def reply_cmd(
    entry_id: str = typer.Argument(..., help="EntryID of the email to reply to."),
    body: str = typer.Option(..., "--body", help="Plain-text reply body."),
    account: str = typer.Option(..., "--account", help="Sending account substring."),
    reply_all: bool = typer.Option(False, "--reply-all", help="Reply to all recipients."),
    send_as: str = typer.Option("", "--send-as", help="SMTP address to send AS."),
) -> None:
    """Reply to an email."""
    _run(
        lambda s: ops_messages.reply_email(
            s,
            entry_id=entry_id,
            body=body,
            account=account,
            reply_all=reply_all,
            send_as=send_as,
        )
    )


@app.command("forward")
def forward_cmd(
    entry_id: str = typer.Argument(..., help="EntryID of the email to forward."),
    to: str = typer.Option(..., "--to", help="Recipient address(es)."),
    account: str = typer.Option(..., "--account", help="Sending account substring."),
    cc: str = typer.Option("", "--cc", help="CC address(es)."),
    body: str = typer.Option("", "--body", help="Optional plain-text body to prepend."),
    send_as: str = typer.Option("", "--send-as", help="SMTP address to send AS."),
) -> None:
    """Forward an email."""
    _run(
        lambda s: ops_messages.forward_email(
            s,
            entry_id=entry_id,
            to=to,
            account=account,
            cc=cc,
            body=body,
            send_as=send_as,
        )
    )


@app.command("download")
def download_cmd(
    entry_id: str = typer.Argument(..., help="EntryID of the email."),
    include_inline: bool = typer.Option(
        False, "--include-inline", help="Also save inline attachments."
    ),
) -> None:
    """Download an email's attachments."""
    _run(
        lambda s: ops_attachments.download_attachments(
            s, entry_id, include_inline=include_inline
        )
    )


@app.command("mark-read")
def mark_read_cmd(
    entry_id: str = typer.Argument(..., help="EntryID of the email."),
    unread: bool = typer.Option(
        False, "--unread", help="Mark as unread instead of read."
    ),
) -> None:
    """Mark an email as read (default) or unread (--unread)."""
    _run(lambda s: ops_messages.mark_as_read(s, entry_id, read=not unread))


@app.command("sync")
def sync_cmd(
    timeout_sec: int = typer.Option(60, "--timeout-sec", help="Sync timeout in seconds."),
    account: str = typer.Option("", "--account", help="Account name substring."),
) -> None:
    """Force an Outlook Send/Receive and wait for it to finish."""
    _run(lambda s: ops_sync.force_sync(s, timeout_sec=timeout_sec, account=account))


@app.command("calendar")
def calendar_cmd(
    days: int = typer.Option(7, "--days", help="Number of days to look ahead."),
    count: int = typer.Option(20, "--count", help="Max events to return."),
    account: str = typer.Option("", "--account", help="Account name substring."),
) -> None:
    """List upcoming calendar events."""
    _run(lambda s: ops_calendar.list_calendar(s, days=days, count=count, account=account))


@app.command("mcp")
def mcp_cmd() -> None:
    """Run the MCP server over stdio (owns its own COM session via an STA worker)."""
    # Deferred import: the server pulls in the MCP/stdio stack that should not load
    # for ordinary CLI invocations, and it owns its own session — do not open one here.
    import email_mcp.server

    email_mcp.server.main()


if __name__ == "__main__":
    app()
