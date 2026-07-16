"""Return-shape definitions for the ops layer.

These document the JSON that each tool returns (identical across CLI and MCP).
Ops functions return plain ``dict``/``list`` matching these shapes — the TypedDicts
exist for clarity and editor help, not runtime validation.
"""

from __future__ import annotations

from typing import Any, Optional, TypedDict


class Account(TypedDict):
    name: str
    entry_id: str


class QueryResult(TypedDict):
    results: list[dict[str, Any]]
    total_returned: int
    total_matched: int
    has_more: bool
    next_offset: Optional[int]


# Output fields query_emails may project (mirrors VALID_FIELDS_OUT in index.js).
VALID_OUTPUT_FIELDS: frozenset[str] = frozenset(
    {
        "entry_id", "subject", "from", "from_name", "to", "cc", "received",
        "sent", "unread", "has_attachments", "preview", "importance", "size",
    }
)

# order_by -> (COM item property, descending?)
ORDER_MAP: dict[str, tuple[str, bool]] = {
    "received_desc": ("ReceivedTime", True),
    "received_asc": ("ReceivedTime", False),
    "sent_desc": ("SentOn", True),
    "sent_asc": ("SentOn", False),
    "subject_asc": ("Subject", False),
}


class ReadEmail(TypedDict):
    entry_id: str
    subject: str
    from_: str  # serialized as "from"
    from_name: str
    to: str
    cc: str
    received: str
    sent: str
    unread: bool
    has_attachments: bool
    attachments: list[str]
    body: str


class SavedAttachment(TypedDict, total=False):
    filename: str
    size: int
    error: str


class DownloadResult(TypedDict, total=False):
    folder: Optional[str]
    saved: list[SavedAttachment]
    skipped_inline: int
    note: str


class CalendarEvent(TypedDict):
    subject: str
    start: str
    end: str
    location: str
    organizer: str
    all_day: bool
    body: str
