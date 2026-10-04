"""Small text helpers shared by the send/reply/forward ops."""

from __future__ import annotations

import re


def text_to_html(s: str) -> str:
    """Escape a plain-text string and convert newlines to <br> (port of textToHtml).

    Matches index.js: escape ``&``, ``<``, ``>``, ``"`` then turn CRLF/CR/LF into ``<br>``.
    """
    escaped = (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
    # Normalize CRLF and CR to LF first, then replace LF.
    escaped = escaped.replace("\r\n", "\n").replace("\r", "\n")
    return escaped.replace("\n", "<br>")


# Id on the div wrapping the body text we author in a draft. edit_draft swaps only
# this div's contents, so a reply/forward draft keeps its quoted thread intact.
BODY_MARKER_ID = "email-mcp-body"

_MARKER_OPEN_RE = re.compile(
    r"<div\b[^>]*\bid=[\"']?" + BODY_MARKER_ID + r"[\"']?[^>]*>", re.IGNORECASE
)
_DIV_TAG_RE = re.compile(r"<(/?)div\b[^>]*>", re.IGNORECASE)
# Outlook's quoted-thread separators on reply/forward items.
_QUOTED_THREAD_RE = re.compile(r"divRplyFwdMsg|border-top:\s*solid|<hr\b", re.IGNORECASE)


def marked_body_html(body: str) -> str:
    """Div-wrap ``body`` as HTML, tagged with the marker edit_draft looks for."""
    return f'<div id="{BODY_MARKER_ID}">{text_to_html(body)}</div>'


def replace_marked_body(html: str, body: str) -> str | None:
    """Return ``html`` with the marked div's contents replaced by ``body``.

    Walks div nesting to find the matching close, so markup Outlook nests inside the
    div on save cannot truncate the match. ``None`` when no marker is present.
    """
    m = _MARKER_OPEN_RE.search(html)
    if not m:
        return None
    depth = 1
    for tag in _DIV_TAG_RE.finditer(html, m.end()):
        depth += -1 if tag.group(1) else 1
        if depth == 0:
            return html[: m.end()] + text_to_html(body) + html[tag.start() :]
    return None


def has_quoted_thread(html: str) -> bool:
    """True when ``html`` looks like a reply/forward with a quoted original below."""
    return bool(_QUOTED_THREAD_RE.search(html))
