"""Unit tests for attachment path normalization and slugging.

Pure — no COM, no filesystem writes. Exercises ``normalize_attachment_path`` and
``slug`` against the behavior ported from index.js.
"""

from __future__ import annotations

import pytest

from email_mcp.errors import InvalidAttachmentPath
from email_mcp.ops.attachments import normalize_attachment_path, slug


class TestNormalizeAttachmentPath:
    def test_absolute_drive_path_passes(self) -> None:
        assert normalize_attachment_path(r"C:\Users\me\file.txt") == r"C:\Users\me\file.txt"

    def test_forward_slashes_normalized_to_backslashes(self) -> None:
        assert normalize_attachment_path("C:/Users/me/file.txt") == r"C:\Users\me\file.txt"

    def test_mixed_slashes_and_dotdot_collapsed(self) -> None:
        assert normalize_attachment_path("C:/Users/me/../me/file.txt") == r"C:\Users\me\file.txt"

    def test_rooted_no_drive_is_absolute(self) -> None:
        # ntpath.isabs treats a rooted (no-drive) path as absolute — matches path.win32.
        assert normalize_attachment_path("/foo/bar") == r"\foo\bar"

    def test_relative_path_raises(self) -> None:
        with pytest.raises(InvalidAttachmentPath, match="must be absolute"):
            normalize_attachment_path("relative/file.txt")

    def test_drive_relative_path_raises(self) -> None:
        # 'C:foo' is drive-relative, not absolute.
        with pytest.raises(InvalidAttachmentPath, match="must be absolute"):
            normalize_attachment_path("C:foo.txt")

    def test_relative_error_uses_original_input(self) -> None:
        with pytest.raises(InvalidAttachmentPath) as exc:
            normalize_attachment_path("foo/bar.txt")
        # Message quotes the ORIGINAL path, not the normalized one.
        assert "foo/bar.txt" in str(exc.value)

    def test_empty_string_raises_nonempty(self) -> None:
        with pytest.raises(InvalidAttachmentPath, match="non-empty string"):
            normalize_attachment_path("")

    def test_whitespace_only_raises_nonempty(self) -> None:
        with pytest.raises(InvalidAttachmentPath, match="non-empty string"):
            normalize_attachment_path("   ")

    def test_non_string_raises_nonempty(self) -> None:
        with pytest.raises(InvalidAttachmentPath, match="non-empty string"):
            normalize_attachment_path(None)  # type: ignore[arg-type]


class TestSlug:
    def test_lowercases(self) -> None:
        assert slug("HELLO", 50) == "hello"

    def test_punctuation_collapsed_to_single_dash(self) -> None:
        assert slug("hello, world!!!", 50) == "hello-world"

    def test_spaces_and_symbols_collapse(self) -> None:
        assert slug("Re:  Meeting @ 3pm", 50) == "re-meeting-3pm"

    def test_leading_trailing_dashes_trimmed(self) -> None:
        assert slug("!!!edge!!!", 50) == "edge"

    def test_empty_returns_empty(self) -> None:
        assert slug("", 50) == ""

    def test_all_punctuation_returns_empty(self) -> None:
        assert slug("!!!???", 50) == ""

    def test_capped_to_maxlen(self) -> None:
        assert slug("abcdefghij", 5) == "abcde"

    def test_cap_retrims_trailing_dash(self) -> None:
        # Cap lands right after 'abc' where the next char is a separator -> retrim.
        assert slug("abc def ghij", 4) == "abc"

    def test_alphanumerics_preserved(self) -> None:
        assert slug("file123name", 50) == "file123name"

    def test_unicode_treated_as_separator(self) -> None:
        # Non-ASCII alphanumerics are outside [a-z0-9] -> collapse to dash.
        assert slug("cafeémenu", 50) == "cafe-menu"
