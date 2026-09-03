"""Unit tests for the draft-vs-send split on reply_email and forward_email.

Pure unit tests — no COM runtime or pywin32 dependencies required.

Regression coverage for the 2026-09-03 incident: a request to "draft a reply"
had no correct tool to reach for — reply_email always called .Send(), and
draft_email only composes a brand-new (unthreaded) message — so the reply went
out live instead of landing in Drafts. reply_email/forward_email now accept
draft=True and call .Save() instead of .Send(), mirroring draft_email's
existing send/draft split but preserving reply/forward threading.
"""

from __future__ import annotations

from typing import Any

from email_mcp.ops.messages import forward_email, reply_email


class FakeReplyItem:
    def __init__(self, entry_id: str = "reply-entry-1") -> None:
        self.HTMLBody: str = "<original body>"
        self.To: str = ""
        self.CC: str = ""
        self.SendUsingAccount: Any = None
        self.EntryID: str = entry_id
        self.saved: bool = False
        self.sent: bool = False

    def Save(self) -> None:
        self.saved = True

    def Send(self) -> None:
        self.sent = True


class FakeSourceItem:
    """The original mail item that .Reply()/.ReplyAll()/.Forward() is called on."""

    def __init__(self, reply_item: FakeReplyItem, forward_item: FakeReplyItem) -> None:
        self._reply_item = reply_item
        self._forward_item = forward_item
        self.reply_calls = 0
        self.reply_all_calls = 0
        self.forward_calls = 0

    def Reply(self) -> FakeReplyItem:
        self.reply_calls += 1
        return self._reply_item

    def ReplyAll(self) -> FakeReplyItem:
        self.reply_all_calls += 1
        return self._reply_item

    def Forward(self) -> FakeReplyItem:
        self.forward_calls += 1
        return self._forward_item


class FakeAccount:
    def __init__(self, name: str = "Work", smtp: str = "work@example.com") -> None:
        self.DisplayName = name
        self.SmtpAddress = smtp


class FakeSession:
    def __init__(self) -> None:
        self.reply_item = FakeReplyItem(entry_id="reply-entry-1")
        self.forward_item = FakeReplyItem(entry_id="forward-entry-1")
        self.source_item = FakeSourceItem(self.reply_item, self.forward_item)
        self.send_acct = FakeAccount()
        self._accounts_list = [{"name": "Work", "entry_id": "store-entry-1"}]

    def get_item(self, entry_id: str) -> FakeSourceItem:
        return self.source_item

    def list_accounts(self) -> list[dict[str, str]]:
        return self._accounts_list

    def find_send_account(self, substring: str) -> FakeAccount | None:
        if (
            substring.lower() in self.send_acct.DisplayName.lower()
            or substring.lower() in self.send_acct.SmtpAddress.lower()
        ):
            return self.send_acct
        return None

    def current_user_address(self) -> str:
        return "current_user@example.com"


class TestReplyEmailDraft:
    def test_draft_true_saves_instead_of_sends(self) -> None:
        session = FakeSession()

        result = reply_email(
            session=session,
            entry_id="original-entry-1",
            body="Nov 4 works for me.",
            account="Work",
            draft=True,
        )

        assert result["status"] == "draft"
        assert result["entry_id"] == "reply-entry-1"
        assert result["from"] == "work@example.com"
        assert "sent_as" not in result

        assert session.reply_item.saved is True
        assert session.reply_item.sent is False
        assert "Nov 4 works for me." in session.reply_item.HTMLBody
        assert session.reply_item.SendUsingAccount is session.send_acct

    def test_draft_default_false_still_sends(self) -> None:
        """Backward compatibility: omitting draft preserves prior send-on-call behavior."""
        session = FakeSession()

        result = reply_email(
            session=session,
            entry_id="original-entry-1",
            body="Nov 4 works for me.",
            account="Work",
        )

        assert result["status"] == "sent"
        assert session.reply_item.sent is True
        assert session.reply_item.saved is False

    def test_draft_true_uses_reply_all_when_requested(self) -> None:
        session = FakeSession()

        reply_email(
            session=session,
            entry_id="original-entry-1",
            body="body",
            account="Work",
            reply_all=True,
            draft=True,
        )

        assert session.source_item.reply_all_calls == 1
        assert session.source_item.reply_calls == 0
        assert session.reply_item.saved is True


class TestForwardEmailDraft:
    def test_draft_true_saves_instead_of_sends(self) -> None:
        session = FakeSession()

        result = forward_email(
            session=session,
            entry_id="original-entry-1",
            to="alice@example.com",
            account="Work",
            draft=True,
        )

        assert result["status"] == "draft"
        assert result["entry_id"] == "forward-entry-1"
        assert result["to"] == "alice@example.com"
        assert result["from"] == "work@example.com"

        assert session.forward_item.saved is True
        assert session.forward_item.sent is False
        assert session.forward_item.To == "alice@example.com"

    def test_draft_default_false_still_sends(self) -> None:
        session = FakeSession()

        result = forward_email(
            session=session,
            entry_id="original-entry-1",
            to="alice@example.com",
            account="Work",
        )

        assert result["status"] == "sent"
        assert session.forward_item.sent is True
        assert session.forward_item.saved is False
