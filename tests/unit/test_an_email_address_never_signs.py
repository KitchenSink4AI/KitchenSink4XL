"""An email address never signs a note or sets a workbook's author.

A model told to "sign it with my email" would write the user's address into
the workbook for everyone who opens it. An author that contains "@" is never
written, whether the call sends it or the Office or Windows name is one. The
next name in the usual order signs instead: the assistant label with
KS4XL_AUTHOR=assistant, else the user's Office user name, else the Windows
sign-in name, else "Author". The answer's author_source names the one that
signed, and its warnings carry one sentence saying the address was set aside.
"""
from __future__ import annotations

import importlib
import zipfile

import openpyxl
import pytest

import family_harness as fh
from xlsx_mcp import server as srv
from xlsx_mcp.ops import annotations, properties
from xlsx_mcp.settings import Settings

EMAIL = "jane.doe@example.com"


@pytest.fixture()
def book(tmp_path):
    p = tmp_path / "notes.xlsx"
    srv.create_workbook(str(p))
    return p


@pytest.fixture()
def authorship():
    return importlib.import_module("xlsx_mcp.core.authorship")


@pytest.fixture()
def office(monkeypatch, authorship):
    """The Office user name this test sets (None: not set)."""
    values = {"UserName": "Dana Reyes", "UserInitials": "DR"}
    monkeypatch.setattr(authorship, "_office_value", lambda name: values.get(name))
    return values


def _author(p, cell):
    return openpyxl.load_workbook(p).active[cell].comment.author


def _sentences(out):
    """The sentences in an answer's warnings that speak of an email address."""
    return [w for w in out.get("warnings") or [] if "email address" in w]


def _package_text(p) -> bytes:
    """Every part of the package, joined: all that anyone opening the file can find."""
    with zipfile.ZipFile(p) as zf:
        return b"\n".join(zf.read(name) for name in zf.namelist())


# ---- notes -------------------------------------------------------------------------------------


@pytest.mark.parametrize("sent", [EMAIL, "Jane Doe <jane.doe@example.com>", " jd@example.com "])
def test_an_email_sent_as_the_author_never_signs_a_note(book, office, sent):
    out = annotations.manage_comment(str(book), "add", location="A1", text="check", author=sent)
    assert _author(book, "A1") == "Dana Reyes"
    comment = out["changed"]["comment"]
    assert comment["author"] == "Dana Reyes" and comment["author_source"] == "office_username"
    sentences = _sentences(out)
    assert len(sentences) == 1, out.get("warnings")
    assert '"Dana Reyes"' in sentences[0] and "@" not in sentences[0]
    assert b"example.com" not in _package_text(book)


def test_the_sentence_says_what_was_set_aside_and_who_signed(book, office):
    out = annotations.manage_comment(str(book), "add", location="A1", text="n", author=EMAIL)
    assert _sentences(out) == [
        'You sent an email address as the author, which is never written into a workbook, '
        'so this note is signed "Dana Reyes", the user\'s Office user name, instead.']


def test_without_an_office_name_the_windows_sign_in_name_signs_instead(book, office, monkeypatch):
    office["UserName"] = None
    monkeypatch.setenv("USERNAME", "dreyes")
    out = annotations.manage_comment(str(book), "add", location="A1", text="n", author=EMAIL)
    assert _author(book, "A1") == "dreyes"
    assert out["changed"]["comment"]["author_source"] == "windows_signin"
    assert len(_sentences(out)) == 1 and "Windows sign-in name" in _sentences(out)[0]


def test_with_no_name_anywhere_the_note_is_signed_author(book, office, monkeypatch):
    office["UserName"] = None
    for var in ("USERNAME", "USER", "LOGNAME"):
        monkeypatch.delenv(var, raising=False)
    out = annotations.manage_comment(str(book), "add", location="A1", text="n", author=EMAIL)
    assert _author(book, "A1") == "Author"
    assert out["changed"]["comment"]["author_source"] == "fallback"
    assert len(_sentences(out)) == 1 and '"Author"' in _sentences(out)[0]


def test_the_assistant_setting_signs_its_label_instead(book, office, monkeypatch, authorship):
    monkeypatch.setattr(authorship, "mode", lambda: authorship.ASSISTANT)
    out = annotations.manage_comment(str(book), "add", location="A1", text="n", author=EMAIL)
    assert _author(book, "A1") == "AI Assistant via KS4XL"
    assert out["changed"]["comment"]["author_source"] == "setting"
    assert len(_sentences(out)) == 1 and "AI Assistant via KS4XL" in _sentences(out)[0]
    assert b"example.com" not in _package_text(book)


def test_an_edit_that_sends_an_email_is_signed_by_the_next_name(book, office):
    annotations.manage_comment(str(book), "add", location="A1", text="n", author="Lee")
    out = annotations.manage_comment(str(book), "edit", location="A1", text="changed", author=EMAIL)
    assert _author(book, "A1") == "Dana Reyes"
    assert out["changed"]["comment"]["author_source"] == "office_username"
    assert len(_sentences(out)) == 1 and "@" not in _sentences(out)[0]
    assert b"example.com" not in _package_text(book)


def test_an_edit_without_an_author_keeps_the_notes_author_and_says_nothing(book, office):
    annotations.manage_comment(str(book), "add", location="A1", text="n", author="Lee")
    out = annotations.manage_comment(str(book), "edit", location="A1", text="changed")
    assert _author(book, "A1") == "Lee" and out["changed"]["comment"]["author_source"] == "kept"
    assert _sentences(out) == []


def test_a_name_still_signs_and_the_answer_says_nothing_about_email(book, office):
    out = annotations.manage_comment(str(book), "add", location="A1", text="n", author="Sam Lee")
    assert _author(book, "A1") == "Sam Lee"
    assert out["changed"]["comment"]["author_source"] == "explicit"
    assert _sentences(out) == []


def test_an_office_or_windows_name_that_is_an_email_is_not_used(book, office, monkeypatch):
    office["UserName"] = "dana@example.com"
    monkeypatch.setenv("USERNAME", "dreyes")
    out = annotations.manage_comment(str(book), "add", location="A1", text="n")
    assert _author(book, "A1") == "dreyes"
    assert out["changed"]["comment"]["author_source"] == "windows_signin"
    assert _sentences(out) == []  # the call sent no address; there is nothing to report
    monkeypatch.setenv("USERNAME", "dreyes@example.com")
    for var in ("USER", "LOGNAME"):
        monkeypatch.delenv(var, raising=False)
    out = annotations.manage_comment(str(book), "add", location="B1", text="n")
    assert _author(book, "B1") == "Author"
    assert out["changed"]["comment"]["author_source"] == "fallback"
    assert b"example.com" not in _package_text(book)


# ---- the workbook's author property -------------------------------------------------------------


def test_an_email_never_sets_the_workbooks_author(book, office):
    out = properties.set_workbook_properties(str(book), author=EMAIL, title="Q3")
    changed = out["changed"]["properties"]
    assert changed["author"] == "Dana Reyes" and changed["author_source"] == "office_username"
    assert changed["title"] == "Q3"  # the other fields are set as usual
    sentences = _sentences(out)
    assert len(sentences) == 1 and "@" not in sentences[0]
    assert 'the author is set to "Dana Reyes"' in sentences[0]
    assert properties.set_workbook_properties(str(book))["properties"]["author"] == "Dana Reyes"
    assert b"example.com" not in _package_text(book)


def test_the_assistant_setting_sets_its_label_for_an_email_author(book, office, monkeypatch, authorship):
    monkeypatch.setattr(authorship, "mode", lambda: authorship.ASSISTANT)
    out = properties.set_workbook_properties(str(book), author=EMAIL)
    changed = out["changed"]["properties"]
    assert changed["author"] == "AI Assistant via KS4XL" and changed["author_source"] == "setting"
    assert len(_sentences(out)) == 1


def test_a_name_and_an_empty_author_set_the_workbook_author_as_before(book, office):
    out = properties.set_workbook_properties(str(book), author="Analyst")
    assert out["changed"]["properties"]["author"] == "Analyst"
    assert "author_source" not in out["changed"]["properties"] and _sentences(out) == []
    assert properties.set_workbook_properties(str(book))["properties"]["author"] == "Analyst"
    properties.set_workbook_properties(str(book), author="")
    assert not properties.set_workbook_properties(str(book))["properties"]["author"]


# ---- through the tool, and the help --------------------------------------------------------------


def test_through_the_tool_the_answers_carry_the_sentence(book, office):
    h = fh.start(Settings())
    payload, err, _ = h.call("xl_format", {"op": "manage_comment", "args": {
        "path": str(book), "action": "add", "location": "B2", "text": "n", "author": EMAIL}})
    assert not err, payload
    assert _author(book, "B2") == "Dana Reyes"
    assert payload["changed"]["comment"]["author_source"] == "office_username"
    assert len(_sentences(payload)) == 1 and "@" not in _sentences(payload)[0]
    payload, err, _ = h.call("xl_workbook", {"op": "set_workbook_properties", "args": {
        "path": str(book), "author": EMAIL}})
    assert not err, payload
    assert payload["changed"]["properties"]["author"] == "Dana Reyes"
    assert len(_sentences(payload)) == 1 and "@" not in _sentences(payload)[0]
    assert b"example.com" not in _package_text(book)


@pytest.mark.parametrize("op", ["manage_comment", "set_workbook_properties"])
def test_the_help_says_an_email_address_never_signs(op):
    doc = " ".join((getattr(srv, op).__doc__ or "").split())
    assert "email address is never written" in doc
