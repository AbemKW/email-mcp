"""Small text helpers shared by the send/reply/forward ops."""

from __future__ import annotations


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
