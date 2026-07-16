"""force_sync — best-effort Send/Receive over Outlook's SyncObjects.

Port of ``buildForceSyncScript`` (index.js) and the ``force_sync`` handler. Outlook
exposes no synchronous "is send/receive done" barrier and its ``SyncObject`` COM
events cannot be bound reliably, so this fires each matched sync group asynchronously,
sleeps a brief grace period, and returns a best-effort result — never a true barrier.

COM is reached only through the provided :class:`~email_mcp.outlook.session.OutlookSession`.
"""

from __future__ import annotations

import time
from typing import Any


def force_sync(session: Any, timeout_sec: int = 60, account: str = "") -> dict:
    """Fire a best-effort Send/Receive and return immediately after a grace period.

    Mirrors the JS tool exactly:

    - ``timeout_sec`` is clamped to ``[1, 600]``.
    - Sync groups are ``namespace.SyncObjects`` filtered to those whose ``Name``
      contains ``account`` (case-insensitive substring). Empty ``account`` matches all.
    - If the filter matches nothing but other SyncObjects exist, returns
      ``ok=False`` with an ``error`` naming the filter.
    - If there are no SyncObjects at all, calls ``namespace.SendAndReceive(False)``
      (best-effort) and returns ``ok=True`` with a ``note``.
    - Otherwise ``.Start()`` each matched group, sleeps ``min(timeout_sec, 5)`` seconds,
      and returns ``ok=True`` with the per-group ``started`` list and a timing ``note``.

    :param session: live :class:`OutlookSession`.
    :param timeout_sec: caller timeout in seconds; clamped to ``[1, 600]``.
    :param account: case-insensitive substring matched against SyncObject names.
    :returns: a dict with keys ``ok``, ``started``, ``timed_out`` and (situationally)
        ``error`` or ``note``.
    """
    timeout_sec = max(1, min(600, timeout_sec))
    acct_filter = str(account or "")

    all_groups: list[Any] = [so for so in session.ns.SyncObjects]
    if acct_filter:
        needle = acct_filter.lower()
        groups = [so for so in all_groups if needle in _group_name(so, -1).lower()]
    else:
        groups = all_groups

    if not groups:
        if acct_filter and len(all_groups) > 0:
            return {
                "ok": False,
                "started": [],
                "timed_out": False,
                "error": f"No SyncObject matches account filter '{acct_filter}'",
            }
        try:
            session.ns.SendAndReceive(False)
        except Exception:
            pass
        return {
            "ok": True,
            "started": [],
            "timed_out": False,
            "note": "No SyncObjects configured; fired SendAndReceive (non-blocking)",
        }

    results: list[dict] = []
    for i, so in enumerate(groups):
        name = _group_name(so, i)
        try:
            so.Start()
            results.append({"name": name, "status": "started"})
        except Exception as e:
            results.append({"name": name, "status": "error", "error": str(e)})

    # Brief grace period to let async sync make headway (capped to keep tool responsive).
    grace_sec = min(timeout_sec, 5)
    if grace_sec > 0:
        time.sleep(grace_sec)
    grace_ms = grace_sec * 1000

    return {
        "ok": True,
        "started": results,
        "timed_out": False,
        "note": (
            "Send/Receive fired asynchronously. Outlook does not expose a sync "
            f"barrier; tool waited {grace_ms}ms then returned."
        ),
    }


def _group_name(so: Any, index: int) -> str:
    """SyncObject ``Name`` as a string, falling back to ``group_<index>`` on failure.

    Matches the JS ``try { [string]$so.Name } catch { "group_$i" }``. When ``index``
    is negative (filter phase, where the JS uses ``-like`` and never falls back) an
    empty string is returned instead, so a name-less group simply fails to match.
    """
    try:
        return str(so.Name)
    except Exception:
        return "" if index < 0 else f"group_{index}"
