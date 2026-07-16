"""Calendar listing operation.

Port of the ``list_calendar`` handler from the old ``index.js`` (the PowerShell it
emitted). Resolves a calendar folder via
:func:`email_mcp.outlook.folders.calendar_folder`, restricts its items to the
``[Start]`` window ``now .. now + days``, and projects up to ``count`` events.

All COM is reached through the passed session / the folder objects it yields; this
module never imports ``win32com`` directly. Per-field COM reads are guarded so a
single event that raises on a property does not abort the whole listing (mirrors
the tolerance of the original PowerShell projection).
"""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

from email_mcp.models import CalendarEvent
from email_mcp.outlook.folders import calendar_folder

# Restrict-filter timestamp format is the US/.NET 'MM/dd/yyyy HH:mm' the DASL
# calendar Restrict expects; output fields use ISO-ish 'yyyy-MM-dd HH:mm'.
_RESTRICT_FMT = "%m/%d/%Y %H:%M"
_OUTPUT_FMT = "%Y-%m-%d %H:%M"

_COUNT_CAP = 50


def _fmt(value: Any, fmt: str) -> str:
    """Format a COM date value with ``fmt``, guarding against read failures."""
    try:
        return datetime(
            value.year, value.month, value.day, value.hour, value.minute, value.second
        ).strftime(fmt)
    except Exception:
        try:
            return str(value)
        except Exception:
            return ""


def _read_str(item: Any, prop: str) -> str:
    """Read a string COM property, returning ``""`` on any failure or None."""
    try:
        val = getattr(item, prop)
    except Exception:
        return ""
    return "" if val is None else str(val)


def _read_bool(item: Any, prop: str) -> bool:
    """Read a boolean COM property, returning ``False`` on any failure."""
    try:
        return bool(getattr(item, prop))
    except Exception:
        return False


def _body_preview(item: Any) -> str:
    """First 200 characters of the event body, stripped (mirrors the JS Substring/Trim)."""
    try:
        body = getattr(item, "Body")
    except Exception:
        return ""
    if body is None:
        return ""
    text = str(body)
    return text[: min(200, len(text))].strip()


def list_calendar(
    session: Any,
    days: int = 7,
    count: int = 20,
    account: str = "",
) -> list[CalendarEvent]:
    """List upcoming calendar events in the window ``now .. now + days``.

    Args:
        session: The live :class:`OutlookSession` (owns all COM).
        days: Size of the look-ahead window in days (falsy -> 7, per the JS default).
        count: Maximum number of events to return (falsy -> 20, capped at 50).
        account: Optional account-name substring selecting which store's calendar to
            read; empty means the namespace default calendar.

    Returns:
        A list of event dicts, each with ``subject``, ``start`` and ``end``
        (``'yyyy-MM-dd HH:mm'``), ``location``, ``organizer``, ``all_day`` and a
        200-char stripped ``body`` preview.
    """
    days = days or 7
    max_items = min(count or 20, _COUNT_CAP)

    start = datetime.now()
    end = start + timedelta(days=days)

    cal = calendar_folder(session, account)
    items = cal.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    restrict = (
        f"[Start] >= '{start.strftime(_RESTRICT_FMT)}' "
        f"AND [Start] <= '{end.strftime(_RESTRICT_FMT)}'"
    )
    events = items.Restrict(restrict)

    results: list[CalendarEvent] = []
    for event in events:
        if len(results) >= max_items:
            break
        try:
            start_val = getattr(event, "Start")
            end_val = getattr(event, "End")
        except Exception:
            start_val = None
            end_val = None
        results.append(
            {
                "subject": _read_str(event, "Subject"),
                "start": _fmt(start_val, _OUTPUT_FMT) if start_val is not None else "",
                "end": _fmt(end_val, _OUTPUT_FMT) if end_val is not None else "",
                "location": _read_str(event, "Location"),
                "organizer": _read_str(event, "Organizer"),
                "all_day": _read_bool(event, "AllDayEvent"),
                "body": _body_preview(event),
            }
        )

    return results
