"""Unit tests for account-scoped mail item creation in send_email and draft_email.

Pure unit tests — no COM runtime or pywin32 dependencies required.
Verifies that items are created via the target account's DeliveryStore Drafts folder
rather than Application.CreateItem(0), preventing store reparenting / transport issues
in multi-account Outlook profiles, while preserving CreateItem(0) as a fallback.
"""

from __future__ import annotations

from typing import Any
import pytest

from email_mcp.ops.messages import draft_email, send_email
from email_mcp.outlook.session import OL_FOLDER_DRAFTS


class FakeAttachments:
    def __init__(self) -> None:
        self.items: list[str] = []

    def Add(self, path: str) -> None:
        self.items.append(path)


class FakeMailItem:
    def __init__(self, via: str) -> None:
        self._via = via
        self.Subject: str = ""
        self.HTMLBody: str = ""
        self.To: str = ""
        self.CC: str = ""
        self.SendUsingAccount: Any = None
        self.Attachments = FakeAttachments()
        self.EntryID: str = f"entry-id-{via}"
        self.saved: bool = False
        self.sent: bool = False

    def Save(self) -> None:
        self.saved = True

    def Send(self) -> None:
        self.sent = True


class FakeItems:
    def __init__(self, item: FakeMailItem) -> None:
        self.item = item
        self.add_calls: list[str] = []

    def Add(self, item_type: str = "IPM.Note") -> FakeMailItem:
        self.add_calls.append(item_type)
        return self.item


class FakeDraftsFolder:
    def __init__(self, item: FakeMailItem) -> None:
        self.Items = FakeItems(item)


class FakeDeliveryStore:
    def __init__(self, drafts_item: FakeMailItem) -> None:
        self.drafts_folder = FakeDraftsFolder(drafts_item)
        self.get_default_folder_calls: list[int] = []

    def GetDefaultFolder(self, folder_type: int) -> FakeDraftsFolder:
        self.get_default_folder_calls.append(folder_type)
        return self.drafts_folder


class FakeAccount:
    def __init__(
        self,
        name: str,
        smtp: str,
        store: FakeDeliveryStore | None = None,
    ) -> None:
        self.DisplayName = name
        self.SmtpAddress = smtp
        if store is not None:
            self.DeliveryStore = store


class FakeApp:
    def __init__(self, default_item: FakeMailItem) -> None:
        self.default_item = default_item
        self.create_item_calls: list[int] = []

    def CreateItem(self, item_type: int) -> FakeMailItem:
        self.create_item_calls.append(item_type)
        return self.default_item


class FakeSession:
    def __init__(
        self,
        account_name: str = "Work",
        account_smtp: str = "work@example.com",
        has_delivery_store: bool = True,
    ) -> None:
        self.delivery_drafts_item = FakeMailItem(via="delivery_store_drafts")
        self.default_store_item = FakeMailItem(via="default_store_createitem")
        self.app = FakeApp(self.default_store_item)

        if has_delivery_store:
            self.delivery_store: FakeDeliveryStore | None = FakeDeliveryStore(self.delivery_drafts_item)
            self.send_acct = FakeAccount(account_name, account_smtp, self.delivery_store)
        else:
            self.delivery_store = None
            self.send_acct = FakeAccount(account_name, account_smtp, None)

        self._accounts_list = [{"name": account_name, "entry_id": "store-entry-1"}]

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


class TestAccountScopedDraftEmail:
    def test_draft_created_via_delivery_store_drafts(self) -> None:
        session = FakeSession()

        result = draft_email(
            session=session,
            to="alice@example.com",
            subject="Test Draft",
            body="Hello from draft",
            account="Work",
            cc="carol@example.com",
        )

        # Assert returned result dict
        assert result["status"] == "draft"
        assert result["to"] == "alice@example.com"
        assert result["from"] == "work@example.com"
        assert result["entry_id"] == "entry-id-delivery_store_drafts"

        # Assert item used is the DeliveryStore one, not CreateItem(0)
        assert session.delivery_drafts_item.saved is True
        assert session.delivery_drafts_item._via == "delivery_store_drafts"
        assert session.delivery_drafts_item.Subject == "Test Draft"
        assert session.delivery_drafts_item.To == "alice@example.com"
        assert session.delivery_drafts_item.CC == "carol@example.com"
        assert "Hello from draft" in session.delivery_drafts_item.HTMLBody
        assert session.delivery_drafts_item.SendUsingAccount is session.send_acct

        # Assert GetDefaultFolder was requested with OL_FOLDER_DRAFTS (16)
        assert session.delivery_store is not None
        assert session.delivery_store.get_default_folder_calls == [OL_FOLDER_DRAFTS]
        assert session.delivery_store.drafts_folder.Items.add_calls == ["IPM.Note"]

        # Default store CreateItem(0) was NOT called
        assert session.app.create_item_calls == []
        assert session.default_store_item.saved is False

    def test_draft_fallback_when_account_lacks_delivery_store(self) -> None:
        session = FakeSession(has_delivery_store=False)

        result = draft_email(
            session=session,
            to="alice@example.com",
            subject="Test Fallback Draft",
            body="Hello from fallback",
            account="Work",
        )

        assert result["status"] == "draft"
        assert result["entry_id"] == "entry-id-default_store_createitem"

        # Default store CreateItem(0) was used
        assert session.app.create_item_calls == [0]
        assert session.default_store_item.saved is True
        assert session.default_store_item._via == "default_store_createitem"
        assert session.default_store_item.Subject == "Test Fallback Draft"
        assert session.default_store_item.SendUsingAccount is session.send_acct

    def test_draft_fallback_when_get_default_folder_raises(self) -> None:
        session = FakeSession()
        assert session.delivery_store is not None

        def _exploding_get_default_folder(folder_type: int) -> Any:
            raise RuntimeError("COM error accessing Drafts folder")

        session.delivery_store.GetDefaultFolder = _exploding_get_default_folder  # type: ignore[method-assign]

        result = draft_email(
            session=session,
            to="alice@example.com",
            subject="Exploding Drafts",
            body="Fallback test",
            account="Work",
        )

        assert result["status"] == "draft"
        assert result["entry_id"] == "entry-id-default_store_createitem"
        assert session.app.create_item_calls == [0]
        assert session.default_store_item.saved is True


class TestAccountScopedSendEmail:
    def test_send_created_via_delivery_store_drafts(self) -> None:
        session = FakeSession()

        result = send_email(
            session=session,
            to="bob@example.com",
            subject="Test Send",
            body="Hello from send",
            account="Work",
            cc="dave@example.com",
        )

        # Assert returned result dict
        assert result["status"] == "sent"
        assert result["to"] == "bob@example.com"
        assert result["from"] == "work@example.com"

        # Assert item used is the DeliveryStore one, not CreateItem(0)
        assert session.delivery_drafts_item.sent is True
        assert session.delivery_drafts_item._via == "delivery_store_drafts"
        assert session.delivery_drafts_item.Subject == "Test Send"
        assert session.delivery_drafts_item.To == "bob@example.com"
        assert session.delivery_drafts_item.CC == "dave@example.com"
        assert "Hello from send" in session.delivery_drafts_item.HTMLBody
        assert session.delivery_drafts_item.SendUsingAccount is session.send_acct

        # Assert GetDefaultFolder was requested with OL_FOLDER_DRAFTS (16)
        assert session.delivery_store is not None
        assert session.delivery_store.get_default_folder_calls == [OL_FOLDER_DRAFTS]
        assert session.delivery_store.drafts_folder.Items.add_calls == ["IPM.Note"]

        # Default store CreateItem(0) was NOT called
        assert session.app.create_item_calls == []
        assert session.default_store_item.sent is False

    def test_send_fallback_when_account_lacks_delivery_store(self) -> None:
        session = FakeSession(has_delivery_store=False)

        result = send_email(
            session=session,
            to="bob@example.com",
            subject="Test Fallback Send",
            body="Hello from fallback send",
            account="Work",
        )

        assert result["status"] == "sent"

        # Default store CreateItem(0) was used
        assert session.app.create_item_calls == [0]
        assert session.default_store_item.sent is True
        assert session.default_store_item._via == "default_store_createitem"
        assert session.default_store_item.Subject == "Test Fallback Send"
        assert session.default_store_item.SendUsingAccount is session.send_acct

    def test_send_fallback_when_items_add_raises(self) -> None:
        session = FakeSession()
        assert session.delivery_store is not None

        def _exploding_add(item_type: str = "IPM.Note") -> Any:
            raise RuntimeError("COM error adding item")

        session.delivery_store.drafts_folder.Items.Add = _exploding_add  # type: ignore[method-assign]

        result = send_email(
            session=session,
            to="bob@example.com",
            subject="Exploding Add",
            body="Fallback test",
            account="Work",
        )

        assert result["status"] == "sent"
        assert session.app.create_item_calls == [0]
        assert session.default_store_item.sent is True
