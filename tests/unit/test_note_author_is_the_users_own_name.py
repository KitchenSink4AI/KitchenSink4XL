"""Who a note is signed by: the user's own name by default, the assistant
label only when the server is started that way.

A note added without an author used to be signed with a fixed assistant
name. Like the Word editor, a note is now signed with the name the user's
own Office signs with (the Office user name under HKCU\\Software\\Microsoft\\Office\\Common\\
UserInfo), else the Windows sign-in name, else "Author". The launch setting
KS4XL_AUTHOR=assistant signs "AI Assistant via KS4XL" instead. An author named
in the call always wins, and an edit keeps the note's author unless the call
names one. The answer says who signed and where the name came from.
"""
from __future__ import annotations

import openpyxl
import pytest

import family_harness as fh
from xlsx_mcp import server as srv
from xlsx_mcp.ops import annotations
from xlsx_mcp.settings import Settings, SettingsError


@pytest.fixture()
def book(tmp_path):
    p = tmp_path / "notes.xlsx"
    srv.create_workbook(str(p))
    return p


@pytest.fixture()
def authorship():
    import importlib
    return importlib.import_module("xlsx_mcp.core.authorship")


@pytest.fixture()
def office(monkeypatch, authorship):
    """The Office user name and initials this test sets (None: not set)."""
    values = {"UserName": "Dana Reyes", "UserInitials": "DR"}
    monkeypatch.setattr(authorship, "_office_value", lambda name: values.get(name))
    return values


def _author(p, cell):
    return openpyxl.load_workbook(p).active[cell].comment.author


def test_the_assistant_label_reads_exactly(authorship):
    assert authorship.ASSISTANT_LABEL == "AI Assistant via KS4XL"


def test_a_note_is_signed_with_the_office_user_name(book, office):
    out = annotations.manage_comment(str(book), "add", location="A1", text="check this")
    assert _author(book, "A1") == "Dana Reyes"
    assert out["changed"]["comment"]["author"] == "Dana Reyes"
    assert out["changed"]["comment"]["author_source"] == "office_username"


def test_without_an_office_name_the_windows_sign_in_name_signs(book, office, monkeypatch):
    office["UserName"] = None
    monkeypatch.setenv("USERNAME", "dreyes")
    out = annotations.manage_comment(str(book), "add", location="A1", text="n")
    assert _author(book, "A1") == "dreyes"
    assert out["changed"]["comment"]["author_source"] == "windows_signin"


def test_with_no_name_anywhere_the_note_is_signed_author(book, office, monkeypatch):
    office["UserName"] = None
    for var in ("USERNAME", "USER", "LOGNAME"):
        monkeypatch.delenv(var, raising=False)
    out = annotations.manage_comment(str(book), "add", location="A1", text="n")
    assert _author(book, "A1") == "Author"
    assert out["changed"]["comment"]["author_source"] == "fallback"


def test_an_author_named_in_the_call_wins(book, office):
    out = annotations.manage_comment(str(book), "add", location="A1", text="n", author="Finance team")
    assert _author(book, "A1") == "Finance team"
    assert out["changed"]["comment"]["author_source"] == "explicit"


def test_the_assistant_setting_signs_the_label(book, office, monkeypatch, authorship):
    monkeypatch.setattr(authorship, "mode", lambda: authorship.ASSISTANT)
    out = annotations.manage_comment(str(book), "add", location="A1", text="n")
    assert _author(book, "A1") == "AI Assistant via KS4XL"
    assert out["changed"]["comment"]["author_source"] == "setting"


def test_an_edit_keeps_the_notes_author_unless_the_call_names_one(book, office):
    annotations.manage_comment(str(book), "add", location="A1", text="n", author="Lee")
    out = annotations.manage_comment(str(book), "edit", location="A1", text="changed")
    assert _author(book, "A1") == "Lee"
    assert out["changed"]["comment"]["author"] == "Lee"
    annotations.manage_comment(str(book), "edit", location="A1", author="Kim")
    assert _author(book, "A1") == "Kim"


@pytest.mark.parametrize("raw,mode", [(None, "user"), ("", "user"), ("user", "user"),
                                      ("assistant", "assistant"), (" Assistant ", "assistant")])
def test_the_launch_setting(raw, mode):
    env = {} if raw is None else {"KS4XL_AUTHOR": raw}
    assert Settings.from_env(env).author_mode == mode


def test_an_unknown_setting_stops_the_server_at_start():
    with pytest.raises(SettingsError, match="KS4XL_AUTHOR"):
        Settings.from_env({"KS4XL_AUTHOR": "claude"})


def test_through_the_tool_a_server_started_as_assistant_signs_the_label(book, office):
    h = fh.start(Settings(author_mode="assistant"))
    payload, err, _ = h.call("xl_format", {"op": "manage_comment",
                                           "args": {"path": str(book), "action": "add",
                                                    "location": "B2", "text": "n"}})
    assert not err, payload
    assert _author(book, "B2") == "AI Assistant via KS4XL"


def test_help_names_the_default_and_the_label():
    doc = " ".join((srv.manage_comment.__doc__ or "").split())
    assert "your own name" in doc and "AI Assistant via KS4XL" in doc and "KS4XL_AUTHOR" in doc
