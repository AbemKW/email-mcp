"""Exhaustive, table-driven pins for the DASL filter compiler.

Every expected string here was derived by hand-tracing the original JS
(``index.js`` lines 62-219). If the Python port ever drifts from the JS
behavior, one of these assertions will catch it.
"""

from __future__ import annotations

import pytest

from email_mcp.errors import InvalidFilter
from email_mcp.query import compile_filter

# DASL property strings (as emitted, including the wrapping double-quotes).
SUBJECT = '"urn:schemas:httpmail:subject"'
FROM = '"urn:schemas:httpmail:fromemail"'
FROM_NAME = '"urn:schemas:httpmail:from"'
TO = '"urn:schemas:httpmail:to"'
CC = '"urn:schemas:httpmail:cc"'
BODY = '"urn:schemas:httpmail:textdescription"'
RECEIVED = '"urn:schemas:httpmail:datereceived"'
SENT = '"urn:schemas:mailheader:date"'
READ = '"urn:schemas:httpmail:read"'
HAS_ATT = '"urn:schemas:httpmail:hasattachment"'
IMPORTANCE = '"urn:schemas:httpmail:importance"'
SIZE = '"urn:schemas:httpmail:size"'


# --- match-all / empty shapes -------------------------------------------------


def test_empty_object_matches_all():
    assert compile_filter({}) == ""


def test_empty_and_is_true():
    assert compile_filter({"$and": []}) == "1 = 1"


def test_empty_or_is_false():
    assert compile_filter({"$or": []}) == "1 = 0"


# --- bare value / bare array --------------------------------------------------

BARE_CASES = [
    ({"subject": "hello"}, f"{SUBJECT} = 'hello'"),
    ({"from": "a@x.com"}, f"{FROM} = 'a@x.com'"),
    # bare array -> $in
    (
        {"from": ["a@x.com", "b@x.com"]},
        f"({FROM} = 'a@x.com' OR {FROM} = 'b@x.com')",
    ),
    # single-element bare array
    ({"to": ["a@x.com"]}, f"({TO} = 'a@x.com')"),
    # bare empty array -> $in -> matches nothing
    ({"cc": []}, "1 = 0"),
]


@pytest.mark.parametrize("node,expected", BARE_CASES)
def test_bare_values(node, expected):
    assert compile_filter(node) == expected


# --- string operators ---------------------------------------------------------

STRING_OP_CASES = [
    ({"subject": {"$eq": "hi"}}, f"{SUBJECT} = 'hi'"),
    ({"subject": {"$ne": "hi"}}, f"{SUBJECT} <> 'hi'"),
    ({"subject": {"$contains": "report"}}, f"{SUBJECT} LIKE '%report%'"),
    (
        {"subject": {"$not_contains": "spam"}},
        f"NOT ({SUBJECT} LIKE '%spam%')",
    ),
    ({"subject": {"$starts_with": "RE:"}}, f"{SUBJECT} LIKE 'RE:%'"),
    ({"subject": {"$ends_with": "!"}}, f"{SUBJECT} LIKE '%!'"),
    (
        {"from": {"$in": ["a@x.com", "b@x.com"]}},
        f"({FROM} = 'a@x.com' OR {FROM} = 'b@x.com')",
    ),
    (
        {"from": {"$nin": ["a@x.com", "b@x.com"]}},
        f"({FROM} <> 'a@x.com' AND {FROM} <> 'b@x.com')",
    ),
    ({"from": {"$in": []}}, "1 = 0"),
    ({"from": {"$nin": []}}, "1 = 1"),
    # non-array $in / $nin values short-circuit like empty
    ({"from": {"$in": "a@x.com"}}, "1 = 0"),
    ({"from": {"$nin": "a@x.com"}}, "1 = 1"),
    ({"subject": {"$exists": True}}, f"{SUBJECT} IS NOT NULL"),
    ({"subject": {"$exists": False}}, f"{SUBJECT} IS NULL"),
]


@pytest.mark.parametrize("node,expected", STRING_OP_CASES)
def test_string_operators(node, expected):
    assert compile_filter(node) == expected


# --- SQL quote escaping -------------------------------------------------------

def test_single_quote_escaping():
    assert compile_filter({"subject": "O'Brien"}) == f"{SUBJECT} = 'O''Brien'"


# --- LIKE wildcard escaping ---------------------------------------------------

LIKE_ESCAPE_CASES = [
    ({"subject": {"$contains": "50%"}}, f"{SUBJECT} LIKE '%50[%]%'"),
    ({"subject": {"$contains": "a_b"}}, f"{SUBJECT} LIKE '%a[_]b%'"),
    (
        {"subject": {"$contains": "a%b_c"}},
        f"{SUBJECT} LIKE '%a[%]b[_]c%'",
    ),
    # quote + wildcard together (quote escaped first, then wildcards)
    (
        {"subject": {"$starts_with": "it's 50%"}},
        f"{SUBJECT} LIKE 'it''s 50[%]%'",
    ),
]


@pytest.mark.parametrize("node,expected", LIKE_ESCAPE_CASES)
def test_like_escaping(node, expected):
    assert compile_filter(node) == expected


# --- date formatting ----------------------------------------------------------

DATE_CASES = [
    ({"received": {"$gte": "2026-01-01"}}, f"{RECEIVED} >= '2026-01-01 00:00'"),
    (
        {"received": {"$gte": "2026-01-01T09:30"}},
        f"{RECEIVED} >= '2026-01-01 09:30'",
    ),
    # seconds are parsed but dropped
    (
        {"received": {"$gte": "2026-01-01T09:30:45"}},
        f"{RECEIVED} >= '2026-01-01 09:30'",
    ),
    # space separator instead of 'T'
    (
        {"received": {"$lte": "2026-12-31 23:59"}},
        f"{RECEIVED} <= '2026-12-31 23:59'",
    ),
    # trailing junk after a valid prefix is ignored (no end-anchor in the regex)
    (
        {"received": {"$gt": "2026-01-01xyz"}},
        f"{RECEIVED} > '2026-01-01 00:00'",
    ),
    ({"sent": {"$lt": "2025-06-15"}}, f"{SENT} < '2025-06-15 00:00'"),
    ({"received": "2026-01-01"}, f"{RECEIVED} = '2026-01-01 00:00'"),
]


@pytest.mark.parametrize("node,expected", DATE_CASES)
def test_date_formatting(node, expected):
    assert compile_filter(node) == expected


def test_invalid_date_raises():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter({"received": {"$gte": "not-a-date"}})
    assert "Invalid date: not-a-date" in str(ei.value)
    assert "YYYY-MM-DD" in str(ei.value)


# --- boolean formatting (has_attachments) ------------------------------------

BOOL_CASES = [
    ({"has_attachments": True}, f"{HAS_ATT} = 1"),
    ({"has_attachments": False}, f"{HAS_ATT} = 0"),
    ({"has_attachments": {"$eq": True}}, f"{HAS_ATT} = 1"),
    ({"has_attachments": {"$ne": True}}, f"{HAS_ATT} <> 1"),
]


@pytest.mark.parametrize("node,expected", BOOL_CASES)
def test_bool_formatting(node, expected):
    assert compile_filter(node) == expected


# --- unread negation ----------------------------------------------------------

UNREAD_CASES = [
    # unread=true => read=0
    ({"unread": True}, f"{READ} = 0"),
    # unread=false => read=1
    ({"unread": False}, f"{READ} = 1"),
    ({"unread": {"$eq": True}}, f"{READ} = 0"),
    ({"unread": {"$ne": True}}, f"{READ} <> 0"),
    ({"unread": {"$ne": False}}, f"{READ} <> 1"),
    # $in over unread values also negates each element
    (
        {"unread": {"$in": [True, False]}},
        f"({READ} = 0 OR {READ} = 1)",
    ),
    (
        {"unread": {"$nin": [True]}},
        f"({READ} <> 0)",
    ),
]


@pytest.mark.parametrize("node,expected", UNREAD_CASES)
def test_unread_negation(node, expected):
    assert compile_filter(node) == expected


# --- numeric formatting -------------------------------------------------------

NUM_CASES = [
    # integer prints without decimal point (JS String(Number(x)) parity)
    ({"importance": 2}, f"{IMPORTANCE} = 2"),
    ({"importance": {"$gte": 1}}, f"{IMPORTANCE} >= 1"),
    ({"size": {"$gt": 1000000}}, f"{SIZE} > 1000000"),
    # numeric string coerces the same way
    ({"size": "1024"}, f"{SIZE} = 1024"),
    ({"importance": {"$lte": 2}}, f"{IMPORTANCE} <= 2"),
]


@pytest.mark.parametrize("node,expected", NUM_CASES)
def test_num_formatting(node, expected):
    assert compile_filter(node) == expected


def test_invalid_number_raises():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter({"size": {"$gt": "abc"}})
    assert "Invalid number: abc" in str(ei.value)


# --- multiple operators on one field are AND'd --------------------------------

def test_multiple_ops_on_one_field():
    node = {"received": {"$gte": "2026-01-01", "$lt": "2026-02-01"}}
    expected = (
        f"({RECEIVED} >= '2026-01-01 00:00' AND {RECEIVED} < '2026-02-01 00:00')"
    )
    assert compile_filter(node) == expected


# --- multiple top-level field keys wrap each & AND ----------------------------

def test_multiple_top_level_fields():
    node = {"subject": "hi", "unread": False}
    expected = f"({SUBJECT} = 'hi') AND ({READ} = 1)"
    assert compile_filter(node) == expected


# --- combinators: $and / $or --------------------------------------------------

def test_single_child_and_unwraps():
    node = {"$and": [{"subject": "hi"}]}
    assert compile_filter(node) == f"{SUBJECT} = 'hi'"


def test_single_child_or_unwraps():
    node = {"$or": [{"subject": "hi"}]}
    assert compile_filter(node) == f"{SUBJECT} = 'hi'"


def test_multi_child_and_wraps():
    node = {"$and": [{"subject": "hi"}, {"from": "a@x.com"}]}
    expected = f"(({SUBJECT} = 'hi') AND ({FROM} = 'a@x.com'))"
    assert compile_filter(node) == expected


def test_multi_child_or_wraps():
    node = {"$or": [{"subject": "hi"}, {"from": "a@x.com"}]}
    expected = f"(({SUBJECT} = 'hi') OR ({FROM} = 'a@x.com'))"
    assert compile_filter(node) == expected


def test_and_with_empty_children_collapses():
    # children that compile to '' (e.g. {}) are filtered out
    node = {"$and": [{}, {}]}
    assert compile_filter(node) == ""


def test_and_with_one_empty_child_unwraps_survivor():
    node = {"$and": [{}, {"subject": "hi"}]}
    assert compile_filter(node) == f"{SUBJECT} = 'hi'"


# --- $not ---------------------------------------------------------------------

def test_not_wraps_inner():
    node = {"$not": {"subject": "hi"}}
    assert compile_filter(node) == f"NOT ({SUBJECT} = 'hi')"


def test_not_of_empty_is_empty():
    assert compile_filter({"$not": {}}) == ""


# --- nesting ------------------------------------------------------------------

def test_nested_and_or():
    node = {
        "$and": [
            {"$or": [{"from": "a@x.com"}, {"from": "b@x.com"}]},
            {"subject": {"$contains": "report"}},
        ]
    }
    # NB: the $or child is itself parenthesized, and the enclosing $and wraps
    # *every* child in its own parens — hence the doubled parens around the $or.
    # Paren accounting (JS-faithful): the inner $or emits
    #   (("...a") OR ("...b"))
    # then the enclosing $and wraps EACH child in its own parens, so the $or
    # child becomes ((("...a") OR ("...b"))) and the whole thing is wrapped once
    # more by the $and.
    expected = (
        f"(((({FROM} = 'a@x.com') OR ({FROM} = 'b@x.com'))) "
        f"AND ({SUBJECT} LIKE '%report%'))"
    )
    assert compile_filter(node) == expected


# --- slow-field guard (body) --------------------------------------------------

def test_body_slow_field_guard_raises_by_default():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter({"body": {"$contains": "invoice"}})
    assert "is slow (full-text scan)" in str(ei.value)
    assert "allow_slow: true" in str(ei.value)


def test_body_allowed_when_allow_slow():
    node = {"body": {"$contains": "invoice"}}
    assert compile_filter(node, allow_slow=True) == f"{BODY} LIKE '%invoice%'"


def test_allow_slow_threads_through_nesting():
    node = {"$and": [{"body": {"$contains": "invoice"}}]}
    # single-child $and unwraps
    assert compile_filter(node, allow_slow=True) == f"{BODY} LIKE '%invoice%'"
    with pytest.raises(InvalidFilter):
        compile_filter(node, allow_slow=False)


# --- error shapes -------------------------------------------------------------

def test_unknown_field_raises():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter({"nope": "x"})
    assert "Unknown field: nope" in str(ei.value)


def test_unknown_operator_raises():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter({"subject": {"$bogus": "x"}})
    assert "Unknown operator: $bogus" in str(ei.value)


def test_empty_operator_object_raises():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter({"subject": {}})
    assert "Empty operator object for field: subject" in str(ei.value)


def test_non_object_filter_raises():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter([])  # type: ignore[arg-type]
    assert "Filter must be an object" in str(ei.value)


def test_and_requires_array():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter({"$and": {"subject": "hi"}})
    assert "$and requires an array" in str(ei.value)


def test_or_requires_array():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter({"$or": "nope"})
    assert "$or requires an array" in str(ei.value)


def test_contains_requires_string_field():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter({"importance": {"$contains": "x"}})
    assert "$contains requires string field: importance" in str(ei.value)


def test_starts_with_requires_string_field():
    with pytest.raises(InvalidFilter) as ei:
        compile_filter({"size": {"$starts_with": "x"}})
    assert "$starts_with requires string field: size" in str(ei.value)


# --- README example (the canonical end-to-end trace) --------------------------

def test_readme_example_filter():
    node = {
        "$and": [
            {"from": {"$in": ["a@x.com", "b@x.com"]}},
            {"subject": {"$contains": "report"}},
            {"received": {"$gte": "2026-01-01"}},
        ]
    }
    # The $in child compiles to a single parenthesized OR-group; the enclosing
    # $and then wraps every child in its own parens, doubling parens on the $in.
    expected = (
        f"((({FROM} = 'a@x.com' OR {FROM} = 'b@x.com')) "
        f"AND ({SUBJECT} LIKE '%report%') "
        f"AND ({RECEIVED} >= '2026-01-01 00:00'))"
    )
    assert compile_filter(node) == expected
