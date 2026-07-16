"""Pure DASL filter compiler for query_emails.

Compiles a nested filter tree (the same JSON shape the JS tool accepted) into a
DASL boolean expression string suitable for ``Items.Restrict("@SQL=" + expr)``.

This module is a faithful port of the ``compileFilter`` family from the old
Node script (``index.js`` lines 62-219). It contains NO COM access and no
side effects — it is exercised directly by the test suite. The public entry
point is :func:`compile_filter`.

Behavioral contract (matched exactly against the JS):

* An empty object ``{}`` compiles to ``''`` (match-all). The returned string
  never carries the ``@SQL=`` prefix — the caller prepends that.
* A bare scalar value on a field is treated as ``$eq``; a bare array as ``$in``.
* Multiple operators on one field are AND'd together.
* ``$and`` / ``$or`` combinators: empty arrays short-circuit to ``1 = 1`` /
  ``1 = 0``; a single surviving child is emitted unwrapped; multiple children
  are each parenthesized and joined.
* ``unread`` maps to the DASL ``read`` property with a *negated* value
  (``unread=true`` => ``read=0``).
* ``body`` is a slow (full-text) field: compiling any predicate on it raises
  unless ``allow_slow=True``.

All malformed input raises :class:`email_mcp.errors.InvalidFilter`.
"""

from __future__ import annotations

import math
import re
from typing import Any

from email_mcp.errors import InvalidFilter

# ---------- field metadata ----------

FIELD_DASL: dict[str, str] = {
    "subject": "urn:schemas:httpmail:subject",
    "from": "urn:schemas:httpmail:fromemail",
    "from_name": "urn:schemas:httpmail:from",
    "to": "urn:schemas:httpmail:to",
    "cc": "urn:schemas:httpmail:cc",
    "body": "urn:schemas:httpmail:textdescription",
    "received": "urn:schemas:httpmail:datereceived",
    "sent": "urn:schemas:mailheader:date",
    "unread": "urn:schemas:httpmail:read",  # negated; unread=true => read=0
    "has_attachments": "urn:schemas:httpmail:hasattachment",
    "importance": "urn:schemas:httpmail:importance",
    "size": "urn:schemas:httpmail:size",
}

STRING_FIELDS: frozenset[str] = frozenset(
    {"subject", "from", "from_name", "to", "cc", "body"}
)
DATE_FIELDS: frozenset[str] = frozenset({"received", "sent"})
BOOL_FIELDS: frozenset[str] = frozenset({"unread", "has_attachments"})
NUM_FIELDS: frozenset[str] = frozenset({"importance", "size"})

SLOW_FIELDS: frozenset[str] = frozenset({"body"})

# Matches 'YYYY-MM-DD' optionally followed by '[T ]HH:mm[:ss]'. Anchored at the
# start only (mirrors the JS regex, which has no end anchor), so any trailing
# characters are ignored. The seconds group is captured but discarded.
_DATE_RE = re.compile(
    r"^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?"
)


# ---------- scalar formatters ----------


def _esc_sql_string(s: Any) -> str:
    """Escape a value for a single-quoted DASL string literal (double the quotes)."""
    return str(s).replace("'", "''")


def _esc_like_arg(s: Any) -> str:
    """Escape a value for use inside a DASL LIKE pattern.

    In addition to SQL-string escaping, the LIKE wildcards ``%`` and ``_`` are
    escaped as ``[%]`` and ``[_]`` so they match literally.
    """
    return _esc_sql_string(s).replace("%", "[%]").replace("_", "[_]")


def _fmt_date(v: Any) -> str:
    """Normalize a date value to ``'YYYY-MM-DD HH:mm'``.

    Accepts ``'YYYY-MM-DD'`` or ``'YYYY-MM-DD[T ]HH:mm[:ss]'``. Missing time
    components default to ``00``; seconds are ignored.
    """
    m = _DATE_RE.match(str(v))
    if not m:
        raise InvalidFilter(f"Invalid date: {v} (use YYYY-MM-DD or YYYY-MM-DDTHH:mm)")
    y, mo, d = m.group(1), m.group(2), m.group(3)
    hh = m.group(4) or "00"
    mm = m.group(5) or "00"
    return f"{y}-{mo}-{d} {hh}:{mm}"


def _fmt_bool(v: Any) -> str:
    """Format a boolean value as DASL ``'1'`` (truthy) or ``'0'`` (falsy)."""
    return "1" if v else "0"


def _fmt_num(v: Any) -> str:
    """Format a numeric value, mirroring JS ``String(Number(v))`` output.

    Integers print without a decimal point (``2`` -> ``'2'``, not ``'2.0'``).
    Non-finite or non-numeric input raises :class:`InvalidFilter`.
    """
    try:
        n = float(v)
    except (TypeError, ValueError):
        raise InvalidFilter(f"Invalid number: {v}")
    if not math.isfinite(n):
        raise InvalidFilter(f"Invalid number: {v}")
    if n.is_integer():
        return str(int(n))
    return repr(n)


def _fmt_value(field: str, v: Any) -> str:
    """Format ``v`` for ``field`` per that field's type (string/date/bool/num)."""
    if field in STRING_FIELDS:
        return f"'{_esc_sql_string(v)}'"
    if field in DATE_FIELDS:
        return f"'{_fmt_date(v)}'"
    if field in BOOL_FIELDS:
        return _fmt_bool(v)
    if field in NUM_FIELDS:
        return _fmt_num(v)
    raise InvalidFilter(f"Unsupported field: {field}")


def _negate_bool_value(v: Any) -> str:
    """Emit the negated DASL bool for the ``unread`` -> ``read`` mapping."""
    return "0" if v else "1"


# ---------- predicate / operator compilation ----------


def _compile_op(prop: str, field: str, op: str, val: Any, is_unread: bool) -> str:
    """Compile a single ``op``/``val`` pair on ``prop`` into a DASL fragment."""

    def v(raw: Any) -> str:
        return _fmt_value(field, raw)

    def v_unread(raw: Any) -> str:
        return _negate_bool_value(raw)

    if op == "$eq":
        if is_unread:
            return f"{prop} = {v_unread(val)}"
        return f"{prop} = {v(val)}"
    if op == "$ne":
        if is_unread:
            return f"{prop} <> {v_unread(val)}"
        return f"{prop} <> {v(val)}"
    if op == "$in":
        if not isinstance(val, list) or len(val) == 0:
            return "1 = 0"  # empty $in matches nothing
        parts = [f"{prop} = {v_unread(x) if is_unread else v(x)}" for x in val]
        return f"({' OR '.join(parts)})"
    if op == "$nin":
        if not isinstance(val, list) or len(val) == 0:
            return "1 = 1"
        parts = [f"{prop} <> {v_unread(x) if is_unread else v(x)}" for x in val]
        return f"({' AND '.join(parts)})"
    if op == "$contains":
        if field not in STRING_FIELDS:
            raise InvalidFilter(f"$contains requires string field: {field}")
        return f"{prop} LIKE '%{_esc_like_arg(val)}%'"
    if op == "$not_contains":
        if field not in STRING_FIELDS:
            raise InvalidFilter(f"$not_contains requires string field: {field}")
        return f"NOT ({prop} LIKE '%{_esc_like_arg(val)}%')"
    if op == "$starts_with":
        if field not in STRING_FIELDS:
            raise InvalidFilter(f"$starts_with requires string field: {field}")
        return f"{prop} LIKE '{_esc_like_arg(val)}%'"
    if op == "$ends_with":
        if field not in STRING_FIELDS:
            raise InvalidFilter(f"$ends_with requires string field: {field}")
        return f"{prop} LIKE '%{_esc_like_arg(val)}'"
    if op == "$gte":
        return f"{prop} >= {v(val)}"
    if op == "$lte":
        return f"{prop} <= {v(val)}"
    if op == "$gt":
        return f"{prop} > {v(val)}"
    if op == "$lt":
        return f"{prop} < {v(val)}"
    if op == "$exists":
        return f"{prop} IS NOT NULL" if val else f"{prop} IS NULL"
    raise InvalidFilter(f"Unknown operator: {op}")


def _compile_predicate(field: str, op_or_val: Any, allow_slow: bool) -> str:
    """Compile a field predicate (bare value, bare array, or operator object)."""
    dasl = FIELD_DASL.get(field)
    if not dasl:
        raise InvalidFilter(f"Unknown field: {field}")
    if field in SLOW_FIELDS and not allow_slow:
        raise InvalidFilter(
            f"Field '{field}' is slow (full-text scan). Pass allow_slow: true to opt in."
        )

    prop = f'"{dasl}"'
    is_unread = field == "unread"

    # Bare array -> $in; bare non-dict scalar (incl. None) -> $eq.
    if isinstance(op_or_val, list):
        return _compile_op(prop, field, "$in", op_or_val, is_unread)
    if not isinstance(op_or_val, dict):
        return _compile_op(prop, field, "$eq", op_or_val, is_unread)

    keys = list(op_or_val.keys())
    if len(keys) == 0:
        raise InvalidFilter(f"Empty operator object for field: {field}")
    # multiple operators on same field => AND them
    parts = [_compile_op(prop, field, op, op_or_val[op], is_unread) for op in keys]
    return parts[0] if len(parts) == 1 else f"({' AND '.join(parts)})"


def _compile_filter(node: Any, allow_slow: bool) -> str:
    """Recursive compiler over combinators (``$and``/``$or``/``$not``) and fields."""
    if not isinstance(node, dict):
        raise InvalidFilter("Filter must be an object")
    keys = list(node.keys())
    if len(keys) == 0:
        return ""  # empty filter = match all

    parts: list[str] = []
    for k in keys:
        if k == "$and" or k == "$or":
            branch = node[k]
            if not isinstance(branch, list):
                raise InvalidFilter(f"{k} requires an array")
            if len(branch) == 0:
                parts.append("1 = 1" if k == "$and" else "1 = 0")
                continue
            sub = [
                s for s in (_compile_filter(c, allow_slow) for c in branch) if len(s) > 0
            ]
            if len(sub) == 0:
                continue
            joiner = " AND " if k == "$and" else " OR "
            if len(sub) == 1:
                parts.append(sub[0])
            else:
                parts.append(f"({joiner.join(f'({s})' for s in sub)})")
        elif k == "$not":
            inner = _compile_filter(node[k], allow_slow)
            if not inner:
                continue
            parts.append(f"NOT ({inner})")
        else:
            # field predicate
            frag = _compile_predicate(k, node[k], allow_slow)
            if len(frag) > 0:
                parts.append(frag)

    if len(parts) == 0:
        return ""
    if len(parts) == 1:
        return parts[0]
    return " AND ".join(f"({p})" for p in parts)


def compile_filter(node: dict, allow_slow: bool = False) -> str:
    """Compile a filter tree into a DASL boolean expression string.

    Args:
        node: The filter tree (a JSON-style ``dict``). An empty dict matches all.
        allow_slow: When ``True``, permits predicates on slow (full-text) fields
            such as ``body``; otherwise those raise :class:`InvalidFilter`.

    Returns:
        A DASL boolean expression, or ``''`` for a match-all filter. The result
        does NOT include the ``@SQL=`` prefix — the caller prepends that.

    Raises:
        InvalidFilter: For a malformed tree (bad field, operator, date, number,
            or shape).
    """
    return _compile_filter(node, allow_slow)
