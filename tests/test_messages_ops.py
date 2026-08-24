"""Unit tests for query_emails, read_email, item class whitelist, and date guards."""

from __future__ import annotations

from datetime import datetime
from unittest.mock import MagicMock

import pytest

from email_mcp.ops.messages import (
    _fmt_dt,
    _get_sender_email,
    _sort_key_for,
    query_emails,
    read_email,
)
from email_mcp.outlook.session import (
    OL_CLASS_MAIL,
    OL_CLASS_MEETING_REQUEST,
    OL_CLASS_REPORT,
)


def make_mock_item(
    entry_id="ID123",
    subject="Test Subject",
    sender_email="alice@example.com",
    sender_name="Alice",
    to="bob@example.com",
    cc="",
    received=None,
    sent=None,
    unread=False,
    body="Body text",
    item_class=43,
    attachments_count=0,
):
    item = MagicMock()
    item.EntryID = entry_id
    item.Subject = subject
    item.SenderEmailAddress = sender_email
    item.SenderName = sender_name
    item.To = to
    item.CC = cc
    item.ReceivedTime = received
    item.SentOn = sent
    item.UnRead = unread
    item.Body = body
    item.Class = item_class
    item.Importance = 1
    item.Size = 1024
    item.Attachments = MagicMock()
    item.Attachments.Count = attachments_count
    item.Attachments.__iter__.return_value = []
    return item


class TestMeetingAndReportItemHandling:
    def test_query_emails_includes_meeting_and_report_items(self):
        mail_item = make_mock_item(entry_id="M1", subject="Mail Item", item_class=OL_CLASS_MAIL)
        meeting_item = make_mock_item(entry_id="MTG1", subject="Sync Invite", item_class=OL_CLASS_MEETING_REQUEST)
        report_item = make_mock_item(entry_id="RPT1", subject="Delivery Report", item_class=OL_CLASS_REPORT)
        task_item = make_mock_item(entry_id="T1", subject="Task Item", item_class=48)  # olTask

        folder = MagicMock()
        items_collection = MagicMock()
        items_collection.__iter__.return_value = [mail_item, meeting_item, report_item, task_item]
        items_collection.Count = 4
        items_collection.Restrict.return_value = items_collection
        folder.Items = items_collection
        folder.DefaultItemType = 0

        session = MagicMock()
        session.store_roots.return_value = [MagicMock(Folders=[folder], Name="Store1")]

        res = query_emails(session, filter={}, limit=10)
        ids = [r["entry_id"] for r in res["results"]]
        assert "M1" in ids
        assert "MTG1" in ids
        assert "RPT1" in ids
        assert "T1" not in ids

    def test_read_email_reads_meeting_item(self):
        meeting_item = make_mock_item(
            entry_id="MTG1",
            subject="Architecture Sync",
            sender_email="lead@example.com",
            sender_name="Lead Engineer",
            item_class=OL_CLASS_MEETING_REQUEST,
            received=datetime(2026, 8, 20, 10, 0),
        )
        session = MagicMock()
        session.get_item.return_value = meeting_item

        res = read_email(session, "MTG1")
        assert res["entry_id"] == "MTG1"
        assert res["subject"] == "Architecture Sync"
        assert res["from"] == "lead@example.com"
        assert res["from_name"] == "Lead Engineer"
        assert res["received"] == "2026-08-20 10:00:00"

    def test_sender_fallback_on_report_item_missing_sender_email(self):
        report_item = make_mock_item(
            entry_id="RPT1",
            subject="Delivery Status Notification",
            sender_email="",
            sender_name="Mail Delivery Subsystem",
            item_class=OL_CLASS_REPORT,
        )
        del report_item.SenderEmailAddress  # simulate COM attribute error

        assert _get_sender_email(report_item) == "Mail Delivery Subsystem"


class TestSortAfterRestrict:
    def test_sort_is_called_on_restricted_items(self):
        old_item = make_mock_item(entry_id="OLD", received=datetime(2026, 1, 1), item_class=43)
        new_item = make_mock_item(entry_id="NEW", received=datetime(2026, 8, 1), item_class=43)

        initial_items = MagicMock()
        restricted_items = MagicMock()
        initial_items.Restrict.return_value = restricted_items
        restricted_items.__iter__.return_value = [new_item, old_item]
        restricted_items.Count = 2

        folder = MagicMock()
        folder.Items = initial_items
        folder.DefaultItemType = 0

        session = MagicMock()
        session.store_roots.return_value = [MagicMock(Folders=[folder], Name="Store1")]

        res = query_emails(session, filter={"subject": "test"}, order_by="received_desc", limit=1)

        # Assert Sort was called on the restricted collection
        restricted_items.Sort.assert_called_once_with("[ReceivedTime]", True)
        assert res["results"][0]["entry_id"] == "NEW"


class TestSentinelDate4501:
    def test_fmt_dt_excludes_sentinel_year(self):
        sentinel_dt = datetime(4501, 1, 1, 0, 0, 0)
        assert _fmt_dt(sentinel_dt, "%Y-%m-%d %H:%M") == ""
        assert _fmt_dt(sentinel_dt, "%Y-%m-%d %H:%M:%S") == ""

    def test_sort_key_treats_sentinel_as_zero(self):
        sentinel_item = make_mock_item(received=datetime(4501, 1, 1, 0, 0, 0))
        real_item = make_mock_item(received=datetime(2026, 8, 20, 12, 0, 0))

        key_fn = _sort_key_for("ReceivedTime")
        assert key_fn(sentinel_item) == 0.0
        assert key_fn(real_item) > 0.0
