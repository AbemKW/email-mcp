"""Unit tests for list_drafts / edit_draft / delete_draft and the body marker.

Pure unit tests — no COM runtime or pywin32 required. The key invariants:
- mutating ops refuse anything that is not an unsent item in a Drafts folder;
- edit_draft changes only the fields passed;
- on a reply/forward draft, edit_draft swaps only the authored text and keeps the
  quoted thread;
- an Outlook-authored reply (no marker, quoted thread) is refused, not clobbered.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

import pytest

from email_mcp.errors import EmailMcpError, EmailNotFound, InvalidAttachmentPath, NotADraft
from email_mcp.ops.drafts import delete_draft, edit_draft, list_drafts
from email_mcp.ops.messages import draft_email
from email_mcp.text import marked_body_html, replace_marked_body

QUOTED = '<div id="divRplyFwdMsg"><b>From:</b> Laura</div><p>original thread</p>'


class FakeAttachment:
    def __init__(self, name: str, owner: "FakeAttachments") -> None:
        self.FileName = name
        self._owner = owner

    def Delete(self) -> None:
        self._owner.items.remove(self)


class FakeAttachments:
    def __init__(self, names: list[str] | None = None) -> None:
        self.items = [FakeAttachment(n, self) for n in (names or [])]

    @property
    def Count(self) -> int:
        return len(self.items)

    def Item(self, i: int) -> FakeAttachment:
        return self.items[i - 1]

    def Add(self, path: str) -> None:
        self.items.append(FakeAttachment(path.rsplit("\\", 1)[-1], self))

    def __iter__(self):
        return iter(list(self.items))


class FakeFolder:
    def __init__(self, entry_id: str, store: "FakeStore") -> None:
        self.EntryID = entry_id
        self.Store = store
        self.Items: list[Any] = []


class FakeStore:
    def __init__(self, name: str) -> None:
        self.Name = name
        self.drafts = FakeFolder(f"{name}-drafts", self)
        self.inbox = FakeFolder(f"{name}-inbox", self)
        self.Store = self  # store roots expose .Store, like Outlook's root folder

    def GetDefaultFolder(self, kind: int) -> FakeFolder:
        assert kind == 16
        return self.drafts


class FakeItem:
    Class = 43

    def __init__(self, entry_id: str, folder: FakeFolder, *, sent: bool = False, html: str = "") -> None:
        self.EntryID = entry_id
        self.Parent = folder
        self.Sent = sent
        self.HTMLBody = html
        self.Body = "plain"
        self.To = "a@x.com"
        self.CC = ""
        self.Subject = "Subj"
        self.LastModificationTime = datetime(2026, 10, 3, 9, 0)
        self.Attachments = FakeAttachments()
        self.saved = 0
        self.deleted = False
        folder.Items.append(self)

    def Save(self) -> None:
        self.saved += 1

    def Delete(self) -> None:
        self.deleted = True


class FakeSession:
    def __init__(self) -> None:
        self.store = FakeStore("Work")
        self.items: dict[str, FakeItem] = {}

    def add(self, item: FakeItem) -> FakeItem:
        self.items[item.EntryID] = item
        return item

    def get_item(self, entry_id: str) -> FakeItem:
        if entry_id not in self.items:
            raise EmailNotFound(entry_id)
        return self.items[entry_id]

    def store_roots(self) -> list[FakeStore]:
        return [self.store]


@pytest.fixture
def session() -> FakeSession:
    return FakeSession()


def _draft(session: FakeSession, entry_id: str = "d1", html: str | None = None) -> FakeItem:
    return session.add(
        FakeItem(entry_id, session.store.drafts, html=marked_body_html("old") if html is None else html)
    )


# ---- guard ---------------------------------------------------------------


def test_refuses_sent_item(session: FakeSession) -> None:
    session.add(FakeItem("s1", session.store.drafts, sent=True))
    with pytest.raises(NotADraft):
        edit_draft(session, "s1", subject="x")
    with pytest.raises(NotADraft):
        delete_draft(session, "s1")


def test_refuses_item_outside_drafts_folder(session: FakeSession) -> None:
    received = session.add(FakeItem("r1", session.store.inbox))
    with pytest.raises(NotADraft):
        delete_draft(session, "r1")
    assert received.deleted is False


# ---- edit ----------------------------------------------------------------


def test_edit_changes_only_passed_fields(session: FakeSession) -> None:
    item = _draft(session)
    result = edit_draft(session, "d1", subject="New subject")
    assert item.Subject == "New subject"
    assert item.To == "a@x.com"
    assert "old" in item.HTMLBody
    assert result["changed"] == ["subject"]
    assert item.saved == 1


def test_edit_cc_empty_string_clears(session: FakeSession) -> None:
    item = _draft(session)
    item.CC = "b@x.com"
    edit_draft(session, "d1", cc="")
    assert item.CC == ""


def test_edit_no_fields_does_not_save(session: FakeSession) -> None:
    item = _draft(session)
    result = edit_draft(session, "d1")
    assert result["changed"] == []
    assert item.saved == 0


def test_edit_body_plain_draft_replaces_marked_div(session: FakeSession) -> None:
    item = _draft(session)
    edit_draft(session, "d1", body="line1\nline2 <b>")
    assert item.HTMLBody == '<div id="email-mcp-body">line1<br>line2 &lt;b&gt;</div>'


def test_edit_body_on_reply_keeps_quoted_thread(session: FakeSession) -> None:
    item = _draft(session, html=marked_body_html("old reply") + QUOTED)
    edit_draft(session, "d1", body="new reply")
    assert "new reply" in item.HTMLBody
    assert "old reply" not in item.HTMLBody
    assert item.HTMLBody.endswith(QUOTED)


def test_edit_body_refuses_outlook_authored_reply(session: FakeSession) -> None:
    item = _draft(session, html="<div>typed in Outlook</div>" + QUOTED)
    with pytest.raises(EmailMcpError):
        edit_draft(session, "d1", body="x")
    assert "typed in Outlook" in item.HTMLBody
    assert item.saved == 0


def test_edit_body_unmarked_without_thread_is_replaced(session: FakeSession) -> None:
    item = _draft(session, html="<div>typed in Outlook</div>")
    edit_draft(session, "d1", body="fresh")
    assert item.HTMLBody == marked_body_html("fresh")


def test_edit_attachments_add_remove_and_report_missing(session: FakeSession, tmp_path) -> None:
    item = _draft(session)
    item.Attachments = FakeAttachments(["a.pdf", "b.pdf"])
    new = tmp_path / "c.txt"
    new.write_text("hi")
    result = edit_draft(
        session, "d1", attachments_add=[str(new)], attachments_remove=["A.PDF", "ghost.pdf"]
    )
    assert result["attachments"] == ["b.pdf", "c.txt"]
    assert result["attachments_not_found"] == ["ghost.pdf"]


def test_bad_attachment_path_leaves_draft_untouched(session: FakeSession) -> None:
    item = _draft(session)
    with pytest.raises(InvalidAttachmentPath):
        edit_draft(session, "d1", subject="changed", attachments_add=["C:\\nope\\missing.pdf"])
    assert item.Subject == "Subj"
    assert item.saved == 0


# ---- delete / list ---------------------------------------------------------


def test_delete_draft(session: FakeSession) -> None:
    item = _draft(session)
    result = delete_draft(session, "d1")
    assert item.deleted is True
    assert result == {"status": "deleted", "entry_id": "d1", "subject": "Subj"}


def test_list_drafts_newest_first_and_limit(session: FakeSession) -> None:
    older = _draft(session, "d1")
    older.LastModificationTime = datetime(2026, 10, 1)
    newer = _draft(session, "d2")
    newer.LastModificationTime = datetime(2026, 10, 3)
    session.add(FakeItem("r1", session.store.inbox))  # not a draft, must not appear

    result = list_drafts(session)
    assert [d["entry_id"] for d in result["drafts"]] == ["d2", "d1"]
    assert result["total"] == 2
    assert result["drafts"][0]["account"] == "Work"

    limited = list_drafts(session, limit=1)
    assert len(limited["drafts"]) == 1
    assert limited["has_more"] is True


# ---- marker helper ---------------------------------------------------------


def test_replace_marked_body_handles_nested_divs() -> None:
    html = '<div id="email-mcp-body">old<div>inner</div> tail</div><p>quoted</p>'
    assert replace_marked_body(html, "new") == '<div id="email-mcp-body">new</div><p>quoted</p>'


def test_replace_marked_body_tolerates_unquoted_attr() -> None:
    assert replace_marked_body("<div id=email-mcp-body>old</div>", "new") == (
        "<div id=email-mcp-body>new</div>"
    )


def test_replace_marked_body_none_without_marker() -> None:
    assert replace_marked_body("<div>plain</div>", "new") is None
