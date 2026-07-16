"""Attachment operations — path validation, slugging, and download.

Ports the attachment logic from the old ``index.js``:
- ``normalizeAttachmentPath`` / ``validateAttachments`` (the send-path validators)
- the ``Slug`` helper and download flow from ``buildDownloadScript``
- the ``download_attachments`` handler

Windows COM is reached only through the provided :class:`OutlookSession` — this
module never imports win32com. Path handling uses ``ntpath`` explicitly so that
absolute/relative decisions match the old ``path.win32`` behavior deterministically.
"""

from __future__ import annotations

import hashlib
import ntpath
import os
import re
import stat as stat_module
from pathlib import Path
from typing import Any

from email_mcp.errors import InvalidAttachmentPath

# MAPI PidTagAttachContentId — carries an inline attachment's Content-ID.
_PROP_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

# Windows-reserved characters, stripped from folder/file names.
_RESERVED_CHARS_RE = re.compile(r'[<>:"/\\|?*]')

# Sentinel Outlook uses for "no received time".
_NULL_DATE = "4501-01-01"


def _is_absolute_win32(p: str) -> bool:
    """Port of Node's ``path.win32.isAbsolute`` (NOT Python's ``ntpath.isabs``).

    Python 3.13 changed ``ntpath.isabs`` so a rooted path without a drive
    (``\\foo\\bar``) is no longer considered absolute; Node has always treated it
    as absolute. Since the JS relied on ``path.win32.isAbsolute``, replicate its
    rule exactly: a path is absolute if it starts with a separator, or is a drive
    letter followed by a separator (``C:\\``). Bare ``C:`` (drive-relative) is not.
    """
    if not p:
        return False
    if p[0] in ("\\", "/"):
        return True
    if (
        len(p) > 2
        and p[0].isalpha()
        and p[1] == ":"
        and p[2] in ("\\", "/")
    ):
        return True
    return False


def normalize_attachment_path(p: str) -> str:
    """Normalize a user-supplied attachment path to an absolute Windows path.

    Accepts forward or back slashes; requires the result to be absolute. Performs
    NO ``~`` or environment-variable expansion and does NOT touch the filesystem
    (callers stat-check separately). Port of ``normalizeAttachmentPath`` in index.js.

    Raises :class:`InvalidAttachmentPath` for empty/non-string or relative paths.
    """
    if not isinstance(p, str) or not p.strip():
        raise InvalidAttachmentPath("Attachment path must be a non-empty string")
    normalized = ntpath.normpath(p.replace("/", "\\"))
    if not _is_absolute_win32(normalized):
        raise InvalidAttachmentPath(f"Attachment path must be absolute: {p}")
    return normalized


def validate_attachments(paths: list[str] | None) -> list[str]:
    """Normalize each path and confirm it is an existing regular file.

    Port of ``validateAttachments`` in index.js. Normalization runs first over all
    paths and raises immediately on the first bad shape (relative/empty). Only the
    filesystem checks aggregate: every offender is collected and reported together
    as ``"Invalid attachment path(s): <p> (not found); <p> (not a regular file)"``.

    Returns the list of normalized absolute paths on success.
    """
    if not paths:
        return []
    normalized = [normalize_attachment_path(p) for p in paths]
    problems: list[str] = []
    for p in normalized:
        try:
            st = os.stat(p)
        except OSError:
            problems.append(f"{p} (not found)")
            continue
        if not stat_module.S_ISREG(st.st_mode):
            problems.append(f"{p} (not a regular file)")
    if problems:
        raise InvalidAttachmentPath(
            "Invalid attachment path(s): " + "; ".join(problems)
        )
    return normalized


def slug(s: str, maxlen: int) -> str:
    """Slugify ``s``: lowercase, collapse non-alphanumerics to ``-``, trim, cap.

    Port of the ``Slug`` PowerShell function in index.js. Empty input yields ``''``.
    After capping to ``maxlen`` the trailing/leading ``-`` are stripped a second
    time so a truncation landing on a separator does not leave a dangling dash.
    """
    if not s:
        return ""
    t = s.lower()
    t = re.sub(r"[^a-z0-9]+", "-", t)
    t = t.strip("-")
    if len(t) > maxlen:
        t = t[:maxlen].strip("-")
    return t


def _received_date_str(item: Any) -> str:
    """Format ``item.ReceivedTime`` as ``YYYY-MM-DD``, falling back to today.

    Mirrors the JS: an empty value or the ``4501-01-01`` null-date sentinel (which
    Outlook uses for "no time") falls back to the current date.
    """
    date_str = ""
    try:
        date_str = item.ReceivedTime.strftime("%Y-%m-%d")
    except Exception:
        date_str = ""
    if not date_str or date_str == _NULL_DATE:
        from datetime import datetime

        date_str = datetime.now().strftime("%Y-%m-%d")
    return date_str


def _sender_slug(item: Any) -> str:
    """Slug the sender (30 chars), falling back to the email local-part then
    ``'unknown-sender'`` — mirrors the JS sender-name resolution."""
    sender_raw = ""
    try:
        sender_raw = str(item.SenderName or "")
    except Exception:
        sender_raw = ""
    if not sender_raw:
        addr = ""
        try:
            addr = str(item.SenderEmailAddress or "")
        except Exception:
            addr = ""
        if addr and "@" in addr:
            sender_raw = addr.split("@")[0]
        else:
            sender_raw = addr
    s = slug(sender_raw, 30)
    return s or "unknown-sender"


def _subject_slug(item: Any) -> str:
    """Slug the subject (50 chars), falling back to ``'no-subject'``."""
    subject_raw = ""
    try:
        subject_raw = str(item.Subject or "")
    except Exception:
        subject_raw = ""
    s = slug(subject_raw, 50)
    return s or "no-subject"


def _folder_name(item: Any) -> str:
    """Build the ``YYYY-MM-DD_<sender>_<subject>`` folder name, sanitized.

    Strips Windows-reserved characters and trailing dots/spaces defensively (the
    sluggers already remove them, but the date portion and joiners are raw).
    """
    name = f"{_received_date_str(item)}_{_sender_slug(item)}_{_subject_slug(item)}"
    name = _RESERVED_CHARS_RE.sub("_", name)
    name = name.rstrip(" .")
    return name


def _sanitize_filename(raw: str) -> str:
    """Sanitize an attachment filename: reserved chars -> ``_``, trim trailing
    dots/spaces, reject empty / all-dots names. Port of the JS per-file cleanup."""
    name = str(raw or "")
    if not name:
        name = "attachment.bin"
    name = _RESERVED_CHARS_RE.sub("_", name).rstrip(" .")
    if not name or re.fullmatch(r"\.+", name):
        name = "attachment.bin"
    return name


def download_attachments(
    session: Any, entry_id: str, include_inline: bool = False
) -> dict:
    """Download an email's attachments to ``~/Downloads/email-attachments/<folder>``.

    Port of ``buildDownloadScript`` + the ``download_attachments`` handler. Inline
    (CID) attachments referenced by the HTML body are skipped unless
    ``include_inline`` is true; a plaintext-only body treats any CID-bearing
    attachment as inline (best effort).

    Returns ``{folder, saved, skipped_inline[, note]}`` matching the JS shape:
    - ``folder`` is the absolute destination folder, or ``None`` if nothing was kept.
    - ``saved`` is a list of ``{filename, size}`` (with ``error`` on failure).
    - ``note`` is present only when nothing was kept but inline images were skipped.
    """
    item = session.get_item(entry_id)

    # --- 1. Decide which attachments are inline ---
    try:
        html_body = str(item.HTMLBody or "")
    except Exception:
        html_body = ""

    keep: list[Any] = []
    skipped_inline = 0
    for att in item.Attachments:
        cid = ""
        try:
            cid = str(att.PropertyAccessor.GetProperty(_PROP_ATTACH_CONTENT_ID) or "")
        except Exception:
            cid = ""
        is_inline = False
        if cid and cid.strip():
            if html_body and ("cid:" + cid.lower()) in html_body.lower():
                is_inline = True
            elif not html_body:
                is_inline = True
        if is_inline and not include_inline:
            skipped_inline += 1
            continue
        keep.append(att)

    if not keep:
        payload: dict = {"folder": None, "saved": [], "skipped_inline": skipped_inline}
        if skipped_inline > 0:
            payload["note"] = (
                "No real attachments — only inline images. "
                "Pass include_inline=true to save them."
            )
        return payload

    # --- 2. Build subfolder name ---
    downloads_root = Path.home() / "Downloads" / "email-attachments"
    folder_name = _folder_name(item)
    folder = downloads_root / folder_name

    # --- 3. Folder reuse / collision handling via .entry_id marker ---
    marker_file = folder / ".entry_id"
    if folder.exists() and marker_file.exists():
        try:
            existing = marker_file.read_text(encoding="utf-8").strip()
        except OSError:
            existing = ""
        if existing != entry_id:
            digest = hashlib.sha1(entry_id.encode("utf-8")).hexdigest()[:8]
            folder_name = folder_name + "_" + digest
            folder = downloads_root / folder_name
            marker_file = folder / ".entry_id"

    folder.mkdir(parents=True, exist_ok=True)
    if not marker_file.exists():
        marker_file.write_text(entry_id, encoding="utf-8", newline="")

    # --- 4. Save each kept attachment, handling filename collisions ---
    saved: list[dict] = []
    used_names: dict[str, bool] = {}
    folder_full = os.path.abspath(str(folder))

    for att in keep:
        try:
            raw_name = str(att.FileName or "")
        except Exception:
            raw_name = ""
        name = _sanitize_filename(raw_name)

        final_name = name
        if final_name.lower() in used_names:
            base, ext = os.path.splitext(name)
            i = 2
            while f"{base} ({i}){ext}".lower() in used_names:
                i += 1
            final_name = f"{base} ({i}){ext}"
        used_names[final_name.lower()] = True

        dest = folder / final_name
        dest_full = os.path.abspath(str(dest))
        # Defense-in-depth: ensure the resolved path stays inside the folder.
        if not os.path.normcase(dest_full).startswith(os.path.normcase(folder_full)):
            saved.append(
                {"filename": final_name, "size": 0, "error": "path traversal blocked"}
            )
            continue
        try:
            att.SaveAsFile(dest_full)
            size = os.path.getsize(dest_full)
            saved.append({"filename": final_name, "size": size})
        except Exception as e:  # pywintypes.com_error or OSError — best-effort message
            saved.append({"filename": final_name, "size": 0, "error": str(e)})

    return {"folder": folder_full, "saved": saved, "skipped_inline": skipped_inline}
