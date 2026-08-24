"""Message operations — the core query/read/send/draft/reply/forward/mark ops.

Native pywin32 COM port of the corresponding handlers in the old ``index.js``
(the ``query_emails``/``read_email``/``send_email``/``reply_email``/
``forward_email``/``mark_as_read`` cases, plus the folder/sort/slice logic that
lived in ``buildQueryScript``). Behavior matches the JS tool-for-tool: field
names, defaults, clamps, date formats, and JSON output keys are identical.
``draft_email`` (below) is new — not part of the original JS tool set — and
mirrors ``send_email``'s composition with ``.Save()`` instead of ``.Send()``.

All COM access goes through the injected :class:`~email_mcp.outlook.session.OutlookSession`
— this module never imports ``win32com`` and never spawns PowerShell. Every COM
read is wrapped in a guarded try/except returning the same safe default the JS
``Get-Field`` produced, so a hostile item cannot abort a query.
"""

from __future__ import annotations

from typing import Any, Callable

from email_mcp.errors import AccountNotFound, InvalidFilter
from email_mcp.models import ORDER_MAP, VALID_OUTPUT_FIELDS
from email_mcp.ops.accounts import resolve_validated_account
from email_mcp.ops.attachments import validate_attachments
from email_mcp.ops.sendas import apply_send_as
from email_mcp.outlook.folders import all_mail_folders
from email_mcp.outlook.session import OL_MAIL_CLASSES
from email_mcp.query import compile_filter
from email_mcp.text import text_to_html

# Default sort when order_by is unknown (mirrors ORDER_MAP.received_desc fallback).
_DEFAULT_SORT: tuple[str, bool] = ORDER_MAP["received_desc"]


def _is_valid_com_date(value: Any) -> bool:
    """Return True if value is a datetime with a realistic year (not an Outlook sentinel)."""
    try:
        year = getattr(value, "year", 0)
        return 1900 <= year < 4500
    except Exception:
        return False


def _fmt_dt(value: Any, fmt: str) -> str:
    """Format a COM datetime like PowerShell's ``.ToString(fmt)``; ``''`` on failure.

    Guards against null/absent values and Outlook sentinel dates (4501-01-01 / 1601-01-01).
    """
    try:
        if not value or not _is_valid_com_date(value):
            return ""
        return value.strftime(fmt)
    except Exception:
        return ""


def _get_sender_email(item: Any) -> str:
    """Return sender email, falling back to SenderName (e.g. for ReportItems/MeetingItems)."""
    try:
        addr = str(getattr(item, "SenderEmailAddress", "") or "")
        if addr:
            return addr
    except Exception:
        pass
    try:
        return str(getattr(item, "SenderName", "") or "")
    except Exception:
        return ""


def _project_field(item: Any, field: str) -> Any:
    """Return one output field for ``item`` — a direct port of PS ``Get-Field``.

    Each read is individually guarded, returning the same default the JS produced:
    ``''`` for string/date/preview fields, ``False`` for the boolean fields,
    ``1`` for importance, and ``0`` for size.
    """
    if field == "entry_id":
        try:
            return str(item.EntryID)
        except Exception:
            return ""
    if field == "subject":
        try:
            return str(item.Subject)
        except Exception:
            return ""
    if field == "from":
        return _get_sender_email(item)
    if field == "from_name":
        try:
            return str(item.SenderName)
        except Exception:
            return ""
    if field == "to":
        try:
            return str(item.To)
        except Exception:
            return ""
    if field == "cc":
        try:
            return str(item.CC)
        except Exception:
            return ""
    if field == "received":
        return _fmt_dt(_safe_get(item, "ReceivedTime"), "%Y-%m-%d %H:%M")
    if field == "sent":
        return _fmt_dt(_safe_get(item, "SentOn"), "%Y-%m-%d %H:%M")
    if field == "unread":
        try:
            return bool(item.UnRead)
        except Exception:
            return False
    if field == "has_attachments":
        try:
            return item.Attachments.Count > 0
        except Exception:
            return False
    if field == "preview":
        try:
            body = item.Body
            if body:
                return str(body)[:150].strip()
        except Exception:
            pass
        return ""
    if field == "importance":
        try:
            return int(item.Importance)
        except Exception:
            return 1
    if field == "size":
        try:
            return int(item.Size)
        except Exception:
            return 0
    # Unreachable for validated fields, but keep parity with PS switch (no match -> null).
    return None


def _safe_get(item: Any, prop: str) -> Any:
    """Read a COM property, returning ``None`` instead of raising."""
    try:
        return getattr(item, prop)
    except Exception:
        return None


def _sort_key_for(prop: str) -> Callable[[Any], Any]:
    """Build a never-raising sort key for the cross-folder re-sort.

    Returns a primitive (float timestamp for date props, lowercased str for Subject) so
    ``sorted`` is stable and immune to None/sentinel dates/tz-aware comparisons.
    """
    if prop in ("ReceivedTime", "SentOn"):

        def date_key(item: Any) -> float:
            try:
                val = getattr(item, prop, None)
                if not val or not _is_valid_com_date(val):
                    return 0.0
                return float(val.timestamp())
            except Exception:
                return 0.0

        return date_key

    def subject_key(item: Any) -> str:
        try:
            val = getattr(item, "Subject", "")
            return "" if val is None else str(val).lower()
        except Exception:
            return ""

    return subject_key


def query_emails(
    session: Any,
    filter: dict | None = None,
    fields: list[str] | None = None,
    account: str = "",
    limit: int = 20,
    offset: int = 0,
    order_by: str = "received_desc",
    allow_slow: bool = False,
) -> dict:
    """Search mail folders and project a page of results.

    Ports ``buildQueryScript`` + the ``query_emails`` handler. Compiles ``filter``
    to DASL, resolves the mail folders for ``account`` (substring over store names,
    else all), per-folder ``Sort`` then optional ``Restrict("@SQL="+expr)``,
    accumulates ``total_matched`` from each restricted collection's ``Count``, takes
    up to ``offset+limit`` mail items (``Class == 43``) per folder, cross-folder
    re-sorts when more than one folder contributed, slices ``[offset:offset+limit]``,
    and projects the requested fields.

    Returns ``{results, total_returned, total_matched, has_more, next_offset}``.
    Raises :class:`~email_mcp.errors.InvalidFilter` on a malformed filter.
    """
    filter = filter or {}

    # Field selection uses the RAW limit (matches JS handler, which checks before clamp).
    if fields is not None and len(fields) > 0:
        bad = [f for f in fields if f not in VALID_OUTPUT_FIELDS]
        if bad:
            valid = ", ".join(sorted(VALID_OUTPUT_FIELDS))
            raise InvalidFilter(
                f"Unknown output field(s): {', '.join(bad)}. Valid: {valid}"
            )
        selected_fields = list(fields)
    elif limit <= 20:
        selected_fields = ["entry_id", "subject", "from", "received", "unread", "has_attachments"]
    else:
        selected_fields = ["subject", "from", "received"]

    expr = compile_filter(filter, allow_slow=allow_slow)

    sort_prop, sort_desc = ORDER_MAP.get(order_by, _DEFAULT_SORT)

    # Clamp AFTER field selection; clamped values drive paging/slicing.
    limit = max(1, min(500, limit))
    offset = max(0, offset)
    per_folder_cap = offset + limit

    mail_folders = all_mail_folders(session, account)

    all_items: list[Any] = []
    total_matched = 0

    for folder in mail_folders:
        try:
            items = folder.Items
        except Exception:
            continue

        # Restrict first, then sort the restricted collection in-place.
        # In Outlook COM, Restrict returns a new collection with uninitialized sort order.
        if expr:
            try:
                items = items.Restrict("@SQL=" + expr)
            except Exception:
                continue

        try:
            items.Sort(f"[{sort_prop}]", sort_desc)
        except Exception:
            pass

        try:
            total_matched += int(items.Count)
        except Exception:
            pass

        taken = 0
        try:
            for item in items:
                if taken >= per_folder_cap:
                    break
                try:
                    if getattr(item, "Class", 0) not in OL_MAIL_CLASSES:
                        continue
                    all_items.append(item)
                    taken += 1
                except Exception:
                    pass
        except Exception:
            pass

    # Cross-folder final sort: run unconditionally for stable sentinel date,
    # timezone, and case-insensitive subject handling.
    try:
        sorted_items = sorted(
            all_items, key=_sort_key_for(sort_prop), reverse=sort_desc
        )
    except Exception:
        sorted_items = all_items

    sliced = sorted_items[offset : offset + limit]

    results: list[dict[str, Any]] = []
    for item in sliced:
        obj: dict[str, Any] = {}
        for f in selected_fields:
            obj[f] = _project_field(item, f)
        results.append(obj)

    has_more = total_matched > (offset + limit)
    next_offset = (offset + limit) if has_more else None

    return {
        "results": results,
        "total_returned": len(results),
        "total_matched": total_matched,
        "has_more": has_more,
        "next_offset": next_offset,
    }


def read_email(session: Any, entry_id: str) -> dict:
    """Return the full detail of one email by EntryID.

    Ports the ``read_email`` handler. Every COM read is guarded (returning the JS
    default) so a partially-broken item still yields a well-formed record. Dates use
    the ``yyyy-MM-dd HH:mm:ss`` format. Raises
    :class:`~email_mcp.errors.EmailNotFound` if the item cannot be resolved.
    """
    item = session.get_item(entry_id)

    def _s(prop: str) -> str:
        try:
            return str(getattr(item, prop))
        except Exception:
            return ""

    try:
        attachments = [str(a.FileName) for a in item.Attachments]
    except Exception:
        attachments = []

    try:
        has_attachments = item.Attachments.Count > 0
    except Exception:
        has_attachments = False

    try:
        unread = bool(item.UnRead)
    except Exception:
        unread = False

    return {
        "entry_id": _s("EntryID"),
        "subject": _s("Subject"),
        "from": _get_sender_email(item),
        "from_name": _s("SenderName"),
        "to": _s("To"),
        "cc": _s("CC"),
        "received": _fmt_dt(_safe_get(item, "ReceivedTime"), "%Y-%m-%d %H:%M:%S"),
        "sent": _fmt_dt(_safe_get(item, "SentOn"), "%Y-%m-%d %H:%M:%S"),
        "unread": unread,
        "has_attachments": has_attachments,
        "attachments": attachments,
        "body": _s("Body"),
    }


def _resolve_send_account(session: Any, account: str) -> Any:
    """Return the send account COM object for ``account`` or raise AccountNotFound.

    Distinct from :func:`resolve_validated_account` (which matches STORE names): this
    matches send-account SmtpAddress and yields the COM object assigned to
    ``SendUsingAccount``.
    """
    acct = session.find_send_account(account)
    if acct is None:
        raise AccountNotFound(account, [a["name"] for a in session.list_accounts()])
    return acct


def send_email(
    session: Any,
    to: str,
    subject: str,
    body: str,
    account: str,
    cc: str = "",
    attachments: list[str] | None = None,
    send_as: str = "",
) -> dict:
    """Compose and send a new mail. Ports the ``send_email`` handler.

    Validates ``account`` against configured store names (fail-fast) and validates
    all attachment paths before touching COM. Builds a mail via ``CreateItem(0)``,
    sets the div-wrapped HTML body, recipients, and attachments, assigns
    ``SendUsingAccount`` from the matching send account, applies ``send_as`` (if any)
    after the account is set and before ``.Send()``, then sends.

    Returns ``{status: 'sent', to, from[, sent_as]}`` where ``from`` is the sending
    account's SmtpAddress (falling back to the current user's address).
    """
    resolve_validated_account(session, account)
    validated_attachments = validate_attachments(attachments)

    item = session.app.CreateItem(0)
    item.Subject = subject
    item.HTMLBody = f"<div>{text_to_html(body)}</div>"
    item.To = to
    if cc:
        item.CC = cc
    for path in validated_attachments:
        item.Attachments.Add(path)

    send_acct = _resolve_send_account(session, account)
    item.SendUsingAccount = send_acct

    try:
        sent_from = str(send_acct.SmtpAddress)
    except Exception:
        sent_from = session.current_user_address()

    apply_send_as(session, item, send_as)
    item.Send()

    result: dict[str, Any] = {"status": "sent", "to": to, "from": sent_from}
    if send_as:
        result["sent_as"] = send_as
    return result


def draft_email(
    session: Any,
    to: str,
    subject: str,
    body: str,
    account: str,
    cc: str = "",
    attachments: list[str] | None = None,
) -> dict:
    """Compose a new mail and save it as a draft instead of sending. Sibling of
    :func:`send_email` — identical composition, but calls ``item.Save()`` instead
    of ``item.Send()``, so the item lands in the account's Drafts folder for later
    review/editing/sending from Outlook itself.

    Validates ``account`` and attachment paths exactly like ``send_email``. Still
    assigns ``SendUsingAccount`` so the draft remembers which account it will send
    from. Does NOT support ``send_as``: Send As rewrites sender identity props at
    submit time, so applying it to an item that is never submitted would silently
    misrepresent the draft's actual sender.

    Returns ``{status: 'draft', entry_id, to, from}`` where ``entry_id`` lets a
    caller later ``read``/edit/send the same item, and ``from`` is the assigned
    account's SmtpAddress (falling back to the current user's address).
    """
    resolve_validated_account(session, account)
    validated_attachments = validate_attachments(attachments)

    item = session.app.CreateItem(0)
    item.Subject = subject
    item.HTMLBody = f"<div>{text_to_html(body)}</div>"
    item.To = to
    if cc:
        item.CC = cc
    for path in validated_attachments:
        item.Attachments.Add(path)

    send_acct = _resolve_send_account(session, account)
    item.SendUsingAccount = send_acct

    try:
        drafted_from = str(send_acct.SmtpAddress)
    except Exception:
        drafted_from = session.current_user_address()

    item.Save()

    try:
        entry_id = str(item.EntryID)
    except Exception:
        entry_id = ""

    return {"status": "draft", "entry_id": entry_id, "to": to, "from": drafted_from}


def reply_email(
    session: Any,
    entry_id: str,
    body: str,
    account: str,
    reply_all: bool = False,
    attachments: list[str] | None = None,
    send_as: str = "",
) -> dict:
    """Reply (or reply-all) to an email. Ports the ``reply_email`` handler.

    Validates ``account`` (store names) and validates all attachment paths before
    touching COM. Creates the reply, PREPENDS the div-wrapped HTML body to the reply's
    existing HTMLBody, sets attachments, assigns ``SendUsingAccount``, applies
    ``send_as`` (if any) before sending.

    Returns ``{status: 'sent', from[, sent_as]}``.
    """
    resolve_validated_account(session, account)
    validated_attachments = validate_attachments(attachments)

    item = session.get_item(entry_id)
    reply = item.ReplyAll() if reply_all else item.Reply()
    reply.HTMLBody = f"<div>{text_to_html(body)}</div>" + reply.HTMLBody

    send_acct = _resolve_send_account(session, account)
    reply.SendUsingAccount = send_acct

    try:
        sent_from = str(send_acct.SmtpAddress)
    except Exception:
        sent_from = session.current_user_address()

    for path in validated_attachments:
        reply.Attachments.Add(path)

    apply_send_as(session, reply, send_as)
    reply.Send()

    result: dict[str, Any] = {"status": "sent", "from": sent_from}
    if send_as:
        result["sent_as"] = send_as
    return result


def forward_email(
    session: Any,
    entry_id: str,
    to: str,
    account: str,
    cc: str = "",
    body: str = "",
    attachments: list[str] | None = None,
    send_as: str = "",
) -> dict:
    """Forward an email. Ports the ``forward_email`` handler.

    Validates ``account`` (store names) and validates all attachment paths before
    touching COM. Creates the forward, sets recipients (CC only when truthy), sets
    additional attachments, and — only if ``body`` is non-empty — PREPENDS the
    div-wrapped HTML body to the forward's existing HTMLBody. Assigns
    ``SendUsingAccount``, applies ``send_as`` (if any) before sending.

    Returns ``{status: 'sent', to, from[, sent_as]}``.
    """
    resolve_validated_account(session, account)
    validated_attachments = validate_attachments(attachments)

    item = session.get_item(entry_id)
    fwd = item.Forward()
    fwd.To = to
    if cc:
        fwd.CC = cc
    if body:
        fwd.HTMLBody = f"<div>{text_to_html(body)}</div>" + fwd.HTMLBody

    send_acct = _resolve_send_account(session, account)
    fwd.SendUsingAccount = send_acct

    try:
        sent_from = str(send_acct.SmtpAddress)
    except Exception:
        sent_from = session.current_user_address()

    for path in validated_attachments:
        fwd.Attachments.Add(path)

    apply_send_as(session, fwd, send_as)
    fwd.Send()

    result: dict[str, Any] = {"status": "sent", "to": to, "from": sent_from}
    if send_as:
        result["sent_as"] = send_as
    return result


def mark_as_read(session: Any, entry_id: str, read: bool = True) -> dict:
    """Mark an email read (``read=True``) or unread. Ports the ``mark_as_read`` handler.

    Sets ``item.UnRead = not read`` and saves. Returns ``{status: 'ok', unread: not read}``.
    """
    item = session.get_item(entry_id)
    item.UnRead = not read
    item.Save()
    return {"status": "ok", "unread": (not read)}
