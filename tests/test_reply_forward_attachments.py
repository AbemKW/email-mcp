"""Unit tests for attachment handling in reply_email and forward_email."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest

from email_mcp.errors import InvalidAttachmentPath
from email_mcp.ops.messages import forward_email, reply_email


class TestReplyEmailAttachments:
    def test_reply_email_validates_and_adds_attachments(self, tmp_path: Path) -> None:
        file1 = tmp_path / "report.pdf"
        file1.write_text("dummy content")
        file2 = tmp_path / "data.csv"
        file2.write_text("a,b,c")

        session = MagicMock()
        item = MagicMock()
        reply = MagicMock()
        send_acct = MagicMock()
        send_acct.SmtpAddress = "user@example.com"

        session.list_accounts.return_value = [{"name": "MyStore"}]
        session.find_send_account.return_value = send_acct
        session.get_item.return_value = item
        item.Reply.return_value = reply
        reply.HTMLBody = "<blockquote>Original</blockquote>"

        result = reply_email(
            session=session,
            entry_id="ENTRY123",
            body="Replying with attachment",
            account="MyStore",
            reply_all=False,
            attachments=[str(file1), str(file2).replace("\\", "/")],
        )

        assert result["status"] == "sent"
        assert result["from"] == "user@example.com"
        assert reply.Attachments.Add.call_count == 2
        reply.Attachments.Add.assert_any_call(str(file1))
        reply.Attachments.Add.assert_any_call(str(file2))
        reply.Send.assert_called_once()

    def test_reply_email_invalid_attachment_raises(self) -> None:
        session = MagicMock()
        session.list_accounts.return_value = [{"name": "MyStore"}]

        with pytest.raises(InvalidAttachmentPath, match="must be absolute"):
            reply_email(
                session=session,
                entry_id="ENTRY123",
                body="Replying with attachment",
                account="MyStore",
                attachments=["relative/file.pdf"],
            )

        # Ensure COM was not touched
        session.get_item.assert_not_called()


class TestForwardEmailAttachments:
    def test_forward_email_validates_and_adds_attachments(self, tmp_path: Path) -> None:
        file1 = tmp_path / "attachment.docx"
        file1.write_text("docx dummy")

        session = MagicMock()
        item = MagicMock()
        fwd = MagicMock()
        send_acct = MagicMock()
        send_acct.SmtpAddress = "user@example.com"

        session.list_accounts.return_value = [{"name": "MyStore"}]
        session.find_send_account.return_value = send_acct
        session.get_item.return_value = item
        item.Forward.return_value = fwd
        fwd.HTMLBody = "<blockquote>Original</blockquote>"

        result = forward_email(
            session=session,
            entry_id="ENTRY123",
            to="recipient@example.com",
            account="MyStore",
            cc="cc@example.com",
            body="Forwarding with extra attachment",
            attachments=[str(file1)],
        )

        assert result["status"] == "sent"
        assert result["to"] == "recipient@example.com"
        assert result["from"] == "user@example.com"
        fwd.Attachments.Add.assert_called_once_with(str(file1))
        fwd.Send.assert_called_once()

    def test_forward_email_invalid_attachment_raises(self) -> None:
        session = MagicMock()
        session.list_accounts.return_value = [{"name": "MyStore"}]

        with pytest.raises(InvalidAttachmentPath, match="non-empty string"):
            forward_email(
                session=session,
                entry_id="ENTRY123",
                to="recipient@example.com",
                account="MyStore",
                attachments=[""],
            )

        # Ensure COM was not touched
        session.get_item.assert_not_called()
