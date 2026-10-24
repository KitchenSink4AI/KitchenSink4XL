"""No Excel save can land anywhere but the target.

A writing COM route on a template (.xltx, .xltm) edited a COPY: Workbooks.Open
without Editable opens a template as a new, unsaved workbook ("Budget1"),
and session.save_workbook's wb.Save wrote that copy as Budget1.xlsx into
Excel's current folder (the user's Documents), silently replacing any file
of that name, while the template stayed untouched and the call answered ok,
saved, verified.

Now:
  - a read-write open passes Editable=True (the tenth Open argument), so a
    template opens as itself (measured in Excel 16: its own path, template
    format, and its save rewrites it in place);
  - open_workbook refuses, closed unsaved, any workbook that is not the file
    it was asked to open, before any edit is made;
  - save_workbook refuses to save a workbook that is not the target path,
    before Excel writes anything, so nothing else can take the file's place.

Excel writes a copy (SaveCopyAs) and never saves over the
file itself; the fake counts those copies as its saves.

These tests drive a fake Excel that models what Excel 16 was measured to do.
"""
from __future__ import annotations

import os
from pathlib import Path
from types import SimpleNamespace

import openpyxl
import pytest

from xlsx_mcp.com import session as S
from xlsx_mcp.core.errors import UnsupportedStructure


class FakeWB:
    def __init__(self, full_name, name=None, path=None, read_only=False):
        self.FullName = full_name
        self.Name = name if name is not None else os.path.basename(full_name)
        self.Path = path if path is not None else os.path.dirname(full_name)
        self.ReadOnly = read_only
        self.closed = False
        self.saves = 0

    def Close(self, *a):
        self.closed = True

    def Save(self):
        raise AssertionError("wb.Save lets Excel replace the file itself")

    def SaveCopyAs(self, path):
        Path(path).write_bytes(b"PK copy")
        self.saves += 1


def _save(wb, p):
    """save_workbook, with the copy it writes removed again."""
    S.discard_saved(S.save_workbook(wb, label="save", path=p))


class FakeWorkbooks:
    """Models the measured behaviour: a template opened without Editable is
    a new workbook named after it plus a number, with no location."""

    def __init__(self, honour_editable=True):
        self.calls = []
        self.books = []
        self.honour_editable = honour_editable

    def Open(self, *args):
        self.calls.append(args)
        path = args[0]
        editable = len(args) >= 10 and args[9] is True and self.honour_editable
        if path.lower().endswith((".xltx", ".xltm")) and not editable:
            stem = Path(path).stem
            wb = FakeWB(f"{stem}1", name=f"{stem}1", path="")
        else:
            wb = FakeWB(os.path.abspath(path), read_only=bool(args[2]))
        self.books.append(wb)
        return wb


def _app(**kw):
    return SimpleNamespace(Workbooks=FakeWorkbooks(**kw))


def _file(tmp_path, name):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(p)
    return str(p)


@pytest.fixture(autouse=True)
def _stand_in_measured(monkeypatch):
    monkeypatch.setattr(S, "_WRITE_STAND_IN_OK", True)


# ------------------------------------------------------------ the open call


def test_a_read_write_open_asks_for_the_file_itself(tmp_path):
    app = _app()
    S.open_workbook(app, _file(tmp_path, "plain.xlsx"))
    args = app.Workbooks.calls[0]
    assert len(args) == 10
    assert args[2] is False  # ReadOnly
    assert args[7] is None and args[8] is None  # Origin, Delimiter (not Missing: measured)
    assert args[9] is True  # Editable


def test_a_read_only_open_is_unchanged(tmp_path):
    app = _app()
    S.open_workbook(app, _file(tmp_path, "plain.xlsx"), read_only=True)
    args = app.Workbooks.calls[0]
    assert len(args) == 7 and args[2] is True


@pytest.mark.parametrize("name", ["Budget.xltx", "Budget.xltm"])
def test_a_template_opens_as_itself(tmp_path, name):
    app = _app()
    p = _file(tmp_path, name)
    wb = S.open_workbook(app, p)
    assert os.path.normcase(wb.FullName) == os.path.normcase(os.path.abspath(p))
    _save(wb, p)
    assert wb.saves == 1


@pytest.mark.parametrize("name", ["Budget.xltx", "Budget.xltm"])
def test_a_copy_excel_makes_is_refused_closed_unsaved(tmp_path, name):
    """An Excel that still opens the template as a new workbook: refused at
    the open, before any edit, and the copy is closed without saving."""
    app = _app(honour_editable=False)
    p = _file(tmp_path, name)
    with pytest.raises(UnsupportedStructure) as ei:
        S.open_workbook(app, p)
    msg = str(ei.value)
    assert "Nothing was changed" in msg and "Budget1" in msg
    copy = app.Workbooks.books[0]
    assert copy.closed and copy.saves == 0


# ------------------------------------------------------------ the one save


def test_save_refuses_a_workbook_that_is_not_the_target(tmp_path):
    p = _file(tmp_path, "Budget.xltx")
    wb = FakeWB("Budget1", name="Budget1", path="")
    with pytest.raises(UnsupportedStructure):
        _save(wb, p)
    assert wb.saves == 0


def test_save_refuses_another_saved_file(tmp_path):
    p = _file(tmp_path, "a.xlsx")
    other = _file(tmp_path, "b.xlsx")
    wb = FakeWB(other)
    with pytest.raises(UnsupportedStructure):
        _save(wb, p)
    assert wb.saves == 0


def test_save_accepts_the_target_in_any_spelling(tmp_path):
    p = _file(tmp_path, "Book.xlsx")
    spellings = [p, p.replace("\\", "/"), os.path.join(os.path.dirname(p), ".", "Book.xlsx")]
    if os.path.normcase("A") != "A":  # upper case names the same file only where paths ignore case
        spellings.insert(1, p.upper())
    for spelling in spellings:
        wb = FakeWB(spelling)
        _save(wb, p)
        assert wb.saves == 1, spelling


def test_a_hard_link_is_not_the_target(tmp_path):
    """Same file, another name: Excel's save would replace that other name
    and leave the target with its old content."""
    p = _file(tmp_path, "Book.xlsx")
    link = str(tmp_path / "Link.xlsx")
    try:
        os.link(p, link)
    except OSError:
        pytest.skip("hard links unavailable here")
    assert S.is_target(FakeWB(link), p) is False


def test_a_synced_folder_web_address_is_the_target_when_name_and_place_match(tmp_path):
    p = _file(tmp_path, "Budget.xlsx")
    url = "https://d.docs.live.net/0123abcd/Documents/Budget.xlsx"
    assert S.is_target(FakeWB(url, name="Budget.xlsx",
                              path="https://d.docs.live.net/0123abcd/Documents"), p)
    assert not S.is_target(FakeWB(url, name="Budget1", path=""), p)
    assert not S.is_target(FakeWB(url, name="Other.xlsx",
                                  path="https://d.docs.live.net/0123abcd/Documents"), p)


def test_a_workbook_that_cannot_say_where_it_lives_is_not_the_target(tmp_path):
    p = _file(tmp_path, "Book.xlsx")
    assert S.is_target(SimpleNamespace(), p) is False
    assert S.is_target(FakeWB(""), p) is False


# ------------------------------------------------------------ every caller


def test_every_save_workbook_call_names_its_target():
    """save_workbook takes the target as a required keyword, so a new route
    cannot save without saying where the save must land."""
    import inspect
    import re
    sig = inspect.signature(S.save_workbook)
    assert sig.parameters["path"].default is inspect.Parameter.empty
    src = Path(S.__file__).resolve().parents[1]
    calls = 0
    for f in src.rglob("*.py"):
        text = f.read_text(encoding="utf-8")
        for m in re.finditer(r"(def )?save_workbook\(", text):
            if m.group(1):
                continue
            depth, i = 0, m.end() - 1
            while i < len(text):
                depth += {"(": 1, ")": -1}.get(text[i], 0)
                if depth == 0:
                    break
                i += 1
            call = text[m.start():i + 1]
            calls += 1
            assert re.search(r"\bpath=", call), f"{f.name}: {call}"
    assert calls >= 2
