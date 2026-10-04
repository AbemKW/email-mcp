"""Draft lifecycle operations — list, edit, and delete unsent drafts.

Siblings of ``draft_email`` / ``reply_email(draft=True)`` / ``forward_email(draft=True)``,
which only ever create drafts. Without these, every revision meant a fresh draft and the
Drafts folder filled up. Every mutating op goes through :func:`_resolve_draft`, which
refuses anything that is not an unsent item in a Drafts folder, so a stray EntryID can
never edit or delete real sent/received mail.

``delete_draft`` calls ``item.Delete()``, which moves the item to Deleted Items — it is
recoverable, never a permanent purge.
"""

from __future__ import annotations

from typing import Any

from email_mcp.errors import EmailMcpError, NotADraft
from email_mcp.ops.attachments import validate_attachments
from email_mcp.ops.messages import _fmt_dt, _safe_get
from email_mcp.outlook.folders import select_stores
from email_mcp.outlook.session import OL_CLASS_MAIL, OL_FOLDER_DRAFTS
from email_mcp.text import has_quoted_thread, marked_body_html, replace_marked_body


def _is_in_drafts_folder(item: Any) -> bool:
    """True when ``item`` lives in its own store's default Drafts folder."""
    parent = item.Parent
    drafts = parent.Store.GetDefaultFolder(OL_FOLDER_DRAFTS)
    return str(parent.EntryID) == str(drafts.EntryID)


def _resolve_draft(session: Any, entry_id: str) -> Any:
    """Return the mail item for ``entry_id`` or raise :class:`NotADraft`."""
    item = session.get_item(entry_id)
    try:
        if item.Class != OL_CLASS_MAIL:
            raise NotADraft(f"EntryID {entry_id} is not a mail item.")
        if item.Sent:
            raise NotADraft(f"EntryID {entry_id} is a sent or received message, not a draft.")
        in_drafts = _is_in_drafts_folder(item)
    except NotADraft:
        raise
    except Exception as e:
        raise NotADraft(f"Could not confirm EntryID {entry_id} is a draft ({e}).") from e
    if not in_drafts:
        raise NotADraft(f"EntryID {entry_id} is not in a Drafts folder; refusing to touch it.")
    return item


def _attachment_names(item: Any) -> list[str]:
    try:
        return [str(a.FileName) for a in item.Attachments]
    except Exception:
        return []


def list_drafts(session: Any, account: str = "", limit: int = 20) -> dict:
    """List unsent drafts (default Drafts folder of each matching store), newest first.

    Returns ``{drafts, total, has_more}``; each draft carries ``entry_id``, ``account``
    (store name), ``to``, ``cc``, ``subject``, ``modified``, ``has_attachments`` and a
    150-char ``preview``.
    """
    limit = max(1, min(200, limit))
    rows: list[tuple[float, dict[str, Any]]] = []

    for store in select_stores(session, account):
        try:
            folder = store.Store.GetDefaultFolder(OL_FOLDER_DRAFTS)
            items = folder.Items
        except Exception:
            continue
        for item in items:
            try:
                if item.Class != OL_CLASS_MAIL:
                    continue
            except Exception:
                continue
            modified = _safe_get(item, "LastModificationTime")
            try:
                ts = modified.timestamp()
            except Exception:
                ts = 0.0
            body = _safe_get(item, "Body")
            rows.append(
                (
                    ts,
                    {
                        "entry_id": str(_safe_get(item, "EntryID") or ""),
                        "account": str(_safe_get(store, "Name") or ""),
                        "to": str(_safe_get(item, "To") or ""),
                        "cc": str(_safe_get(item, "CC") or ""),
                        "subject": str(_safe_get(item, "Subject") or ""),
                        "modified": _fmt_dt(modified, "%Y-%m-%d %H:%M"),
                        "has_attachments": bool(_attachment_names(item)),
                        "preview": str(body)[:150].strip() if body else "",
                    },
                )
            )

    rows.sort(key=lambda r: r[0], reverse=True)
    return {
        "drafts": [r[1] for r in rows[:limit]],
        "total": len(rows),
        "has_more": len(rows) > limit,
    }


def edit_draft(
    session: Any,
    entry_id: str,
    to: str | None = None,
    cc: str | None = None,
    subject: str | None = None,
    body: str | None = None,
    attachments_add: list[str] | None = None,
    attachments_remove: list[str] | None = None,
) -> dict:
    """Edit an unsent draft in place. Only the fields passed are changed.

    ``to``/``cc``/``subject`` replace the current value (``cc=""`` clears CC). ``body``
    replaces the text this server wrote; on a reply/forward draft the quoted thread below
    it is preserved. A body Outlook authored (no marker, quoted thread present) cannot be
    edited safely and raises. ``attachments_remove`` matches by filename; unmatched names
    are reported, not fatal.

    Returns ``{status: 'draft', entry_id, changed, attachments[, attachments_not_found]}``.
    """
    item = _resolve_draft(session, entry_id)
    # Validate before mutating anything so a bad path leaves the draft untouched.
    new_paths = validate_attachments(attachments_add)

    changed: list[str] = []

    if body is not None:
        html = str(_safe_get(item, "HTMLBody") or "")
        swapped = replace_marked_body(html, body)
        if swapped is None:
            if has_quoted_thread(html):
                raise EmailMcpError(
                    "This reply/forward draft has no email-mcp body marker, so its quoted "
                    "thread cannot be separated from the text. Delete it and re-create it "
                    "with reply_email/forward_email(draft=true), or edit it in Outlook."
                )
            swapped = marked_body_html(body)
        item.HTMLBody = swapped
        changed.append("body")
    if to is not None:
        item.To = to
        changed.append("to")
    if cc is not None:
        item.CC = cc
        changed.append("cc")
    if subject is not None:
        item.Subject = subject
        changed.append("subject")

    not_found: list[str] = []
    if attachments_remove:
        wanted = {n.lower() for n in attachments_remove}
        found: set[str] = set()
        # COM collections are 1-based; remove back-to-front so indexes stay valid.
        for i in range(item.Attachments.Count, 0, -1):
            att = item.Attachments.Item(i)
            name = str(att.FileName)
            if name.lower() in wanted:
                found.add(name.lower())
                att.Delete()
        not_found = [n for n in attachments_remove if n.lower() not in found]
        if found:
            changed.append("attachments_remove")
    for path in new_paths:
        item.Attachments.Add(path)
    if new_paths:
        changed.append("attachments_add")

    if changed:
        item.Save()

    result: dict[str, Any] = {
        "status": "draft",
        # Re-read: a store may reassign the EntryID on save.
        "entry_id": str(_safe_get(item, "EntryID") or entry_id),
        "changed": changed,
        "attachments": _attachment_names(item),
    }
    if not_found:
        result["attachments_not_found"] = not_found
    return result


def delete_draft(session: Any, entry_id: str) -> dict:
    """Delete an unsent draft (moved to Deleted Items — recoverable).

    Returns ``{status: 'deleted', entry_id, subject}``.
    """
    item = _resolve_draft(session, entry_id)
    subject = str(_safe_get(item, "Subject") or "")
    item.Delete()
    return {"status": "deleted", "entry_id": entry_id, "subject": subject}
