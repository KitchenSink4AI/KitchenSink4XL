"""A call can sign a note as the assistant: author="assistant".

The launch setting KS4XL_AUTHOR=assistant is the operator's choice, and a user in
conversation cannot reach it. Told to "sign the note as the AI assistant", a
model made up its own name. The author value "assistant" (any case, spaces
around it ignored) is therefore reserved: it signs exactly "AI Assistant via
KS4XL", the label the setting uses, whatever the setting is, and the answer's
author_source says the call asked for it (explicit_assistant). Every other
name is used as given, except an email address, which is never written.
"""
from __future__ import annotations

import importlib

import openpyxl
import pytest

import family_harness as fh
from xlsx_mcp import registry
from xlsx_mcp import server as srv
from xlsx_mcp.ops import annotations, properties
from xlsx_mcp.settings import Settings
from xlsx_mcp.surface.plan import op_text

LABEL = "AI Assistant via KS4XL"


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
    """The Office user name this test sets."""
    values = {"UserName": "Dana Reyes", "UserInitials": "DR"}
    monkeypatch.setattr(authorship, "_office_value", lambda name: values.get(name))
    return values


def _author(p, cell):
    return openpyxl.load_workbook(p).active[cell].comment.author


@pytest.mark.parametrize("asked", ["assistant", "Assistant", " ASSISTANT ", "aSsIsTaNt"])
def test_author_assistant_signs_the_label(book, office, asked):
    out = annotations.manage_comment(str(book), "add", location="A1", text="n", author=asked)
    assert _author(book, "A1") == LABEL
    comment = out["changed"]["comment"]
    assert comment["author"] == LABEL and comment["author_source"] == "explicit_assistant"
    assert out["warnings"] == []


def test_the_label_is_the_one_the_launch_setting_signs(authorship):
    assert authorship.ASSISTANT_LABEL == LABEL


def test_the_call_signs_the_label_whatever_the_launch_setting_is(book, office, monkeypatch, authorship):
    monkeypatch.setattr(authorship, "mode", lambda: authorship.ASSISTANT)
    out = annotations.manage_comment(str(book), "add", location="A1", text="n", author="assistant")
    assert _author(book, "A1") == LABEL
    assert out["changed"]["comment"]["author_source"] == "explicit_assistant"  # the call asked; not "setting"


@pytest.mark.parametrize("given", ["assistants", "AI Assistant", "the assistant", "Review Desk",
                                   "Assistant Manager", "assist"])
def test_every_other_name_is_used_as_given(book, office, given):
    out = annotations.manage_comment(str(book), "add", location="A1", text="n", author=given)
    assert _author(book, "A1") == given
    assert out["changed"]["comment"]["author_source"] == "explicit"


def test_an_edit_that_asks_for_the_assistant_re_signs_the_note(book, office):
    annotations.manage_comment(str(book), "add", location="A1", text="n", author="Lee")
    out = annotations.manage_comment(str(book), "edit", location="A1", text="changed", author="Assistant")
    assert _author(book, "A1") == LABEL
    assert out["changed"]["comment"]["author_source"] == "explicit_assistant"
    out = annotations.manage_comment(str(book), "edit", location="A1", text="changed again")
    assert _author(book, "A1") == LABEL and out["changed"]["comment"]["author_source"] == "kept"


def test_an_empty_author_still_signs_with_the_users_name(book, office):
    out = annotations.manage_comment(str(book), "add", location="A1", text="n", author="")
    assert _author(book, "A1") == "Dana Reyes"
    assert out["changed"]["comment"]["author_source"] == "office_username"


def test_an_email_address_still_gets_the_users_name_not_the_label(book, office):
    out = annotations.manage_comment(str(book), "add", location="A1", text="n",
                                     author="jane.doe@example.com")
    assert _author(book, "A1") == "Dana Reyes"
    assert out["changed"]["comment"]["author_source"] == "office_username"


def test_the_workbook_author_is_no_signature_so_assistant_stays_as_given(book, office):
    properties.set_workbook_properties(str(book), author="assistant")
    assert properties.set_workbook_properties(str(book))["properties"]["author"] == "assistant"


def test_through_the_tool_the_answer_names_the_source(book, office):
    h = fh.start(Settings())
    payload, err, _ = h.call("xl_format", {"op": "manage_comment", "args": {
        "path": str(book), "action": "add", "location": "B2", "text": "n", "author": "assistant"}})
    assert not err, payload
    assert _author(book, "B2") == LABEL
    assert payload["changed"]["comment"]["author"] == LABEL
    assert payload["changed"]["comment"]["author_source"] == "explicit_assistant"


def test_the_op_line_and_the_help_say_how_to_sign_as_the_assistant():
    line = op_text(registry.public_ops()["manage_comment"])
    assert len(line) <= 200
    assert ("Leave author out to sign with the user's own name, or pass author=\"assistant\" "
            "to sign as \"AI Assistant via KS4XL\".") in line
    doc = " ".join((srv.manage_comment.__doc__ or "").split())
    assert 'pass author="assistant" to sign as "AI Assistant via KS4XL"' in doc
    assert "any other author is used as given" in doc
