"""Excel's copy is checked before it replaces the workbook.

The structural verify, the integrity check and the encryption check ran on
the workbook AFTER Excel's copy had replaced it, and a failure put the
backup back. With backup=false there was no backup, so a bad save stayed in
place. They now run on the copy where Excel wrote it, so a bad save never
replaces the workbook, backup or not, and the backup slots stay as they
were. Once the copy is in place one cheap check remains: the file is, byte
for byte, the copy that passed (a SHA-256 comparison; the structural verify
re-reads every sheet and took 12.5 s on a 5.8 MB workbook, so it runs once).
"""
from __future__ import annotations

import hashlib
import io
import os
import tempfile
import zipfile
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.com import session as com_session
from xlsx_mcp.core import safesave
from xlsx_mcp.core import verify as _verify
from xlsx_mcp.core.errors import ValidationFailed, XlMcpError
from xlsx_mcp.ops import comtier

CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _sha(p) -> str:
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def _beside(book) -> list:
    return sorted(n for n in os.listdir(os.path.dirname(book))
                  if n not in (os.path.basename(book), safesave.BACKUP_DIR_NAME))


def _temp_left() -> list:
    return sorted(os.listdir(tempfile.gettempdir()))


def _slots(book) -> dict:
    d = safesave.slot_dir(book)
    return {n: _sha(d / n) for n in safesave.SLOT_POLICY if (d / n).exists()}


@pytest.fixture
def book(tmp_path, monkeypatch):
    folder = tmp_path / "work"
    folder.mkdir()
    p = str(folder / "book.xlsx")
    wb = openpyxl.Workbook()
    wb.active["A1"] = "before"
    wb.save(p)
    temp = tmp_path / "temp"
    temp.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(temp))
    return p


def _edited(src: bytes) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(src))
    wb.active["A1"] = "edited by Excel"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _with_broken_link(data: bytes) -> bytes:
    """A package that passes the structural verify and has a relationship
    pointing at a part that is not there (Excel repairs or refuses it)."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in zin.infolist():
            zout.writestr(i, zin.read(i.filename))
        zout.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            'relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.'
            'org/officeDocument/2006/relationships/drawing" '
            'Target="../drawings/drawing9.xml"/></Relationships>')
    return out.getvalue()


class _Worker:
    pid = 0
    app = object()


class _Manager:
    def acquire(self):
        return _Worker()

    def release_to_pool(self, w):
        pass


class _FakeWB:
    """Excel with the workbook open. `copy` decides what its SaveCopyAs
    writes: 'ok' (the edit), 'garbage' (not a package), 'broken_link',
    'plain' (an unencrypted package whatever the password), 'encrypted'
    (an OLE container, as Excel's encryption writes)."""

    def __init__(self, path, copy):
        self.FullName = os.path.abspath(path)
        self.ReadOnly = False
        self.Password = ""
        self.opened = Path(path).read_bytes()
        self.copy = copy

    def Save(self):
        raise AssertionError("wb.Save lets Excel replace the file itself")

    def SaveCopyAs(self, copy):
        kind = self.copy
        if kind == "garbage":
            data = b"PK\x03\x04 not a workbook at all"
        elif kind == "broken_link":
            data = _with_broken_link(_edited(self.opened))
        elif kind == "encrypted":
            data = CFB_MAGIC + b"\0" * 4088
        else:  # ok, plain
            data = _edited(self.opened)
        Path(copy).write_bytes(data)

    def Close(self, save):
        pass


def _fake_excel(monkeypatch, copy="ok"):
    monkeypatch.setattr(com_session, "run_com",
                        lambda label, job, timeout=None: job(_Manager()))
    monkeypatch.setattr(com_session, "_assert_hygiene", lambda app: ())
    monkeypatch.setattr(com_session, "_restore_hygiene", lambda app, *a: None)
    monkeypatch.setattr(com_session, "com_retry", lambda fn, label=None, **k: fn())
    monkeypatch.setattr(com_session, "open_workbook",
                        lambda app, p, password=None, **k: _FakeWB(p, copy))


def _autofit(book, **kw):
    return comtier._run_mutation("autofit(book.xlsx)", book,
                                 lambda app, wb: {"autofit": True}, **kw)


def _untouched(book, before, slots_before):
    assert _sha(book) == before, "a bad copy replaced the workbook"
    assert _slots(book) == slots_before, "the backup slots changed"
    assert _beside(book) == [] and _temp_left() == []


# ------------------------------------------------- the edit routes


@pytest.mark.parametrize("backup", [False, True])
def test_a_copy_that_is_not_a_workbook_never_replaces_it(book, monkeypatch, backup):
    """With backup=false the broken copy used to replace the workbook and stay."""
    before, slots = _sha(book), _slots(book)
    _fake_excel(monkeypatch, "garbage")
    with pytest.raises(ValidationFailed) as ei:
        _autofit(book, backup=backup)
    _untouched(book, before, slots)
    msg = str(ei.value)
    assert "structural verify" in msg and "book.xlsx was not changed" in msg, msg


@pytest.mark.parametrize("backup", [False, True])
def test_a_copy_with_a_broken_link_never_replaces_it(book, monkeypatch, backup):
    before, slots = _sha(book), _slots(book)
    _fake_excel(monkeypatch, "broken_link")
    with pytest.raises(ValidationFailed) as ei:
        _autofit(book, backup=backup)
    _untouched(book, before, slots)
    msg = str(ei.value)
    assert "broken internal links" in msg and "book.xlsx was not changed" in msg, msg
    assert "drawing9.xml" in msg


def test_the_checks_read_excels_copy_and_run_once(book, monkeypatch):
    """The structural verify (the slow one) reads the copy in the temp folder,
    once; the workbook itself is then compared byte for byte, not re-read."""
    seen = []
    real = _verify.structural_check

    def spy(path):
        seen.append(os.path.abspath(path))
        return real(path)

    monkeypatch.setattr(_verify, "structural_check", spy)
    _fake_excel(monkeypatch, "ok")
    res = _autofit(book)
    assert res["ok"] and res["verified"]
    assert len(seen) == 1, seen
    # Excel's copy is written beside the workbook, so its
    # relative links stay as they are.
    assert os.path.dirname(seen[0]) == os.path.dirname(os.path.abspath(book))
    assert openpyxl.load_workbook(book).active["A1"].value == "edited by Excel"
    assert _beside(book) == [] and _temp_left() == []


def test_the_workbook_in_place_must_be_the_copy_that_was_checked(book, monkeypatch):
    """Something other than the checked copy ends up in place (simulated by
    a promote that writes other bytes). It is reported, never called
    verified, and nothing is put over it: the state before the call stays
    in the prev slot."""
    before = _sha(book)
    other = openpyxl.Workbook()
    other.active["A1"] = "someone else's"
    real_promote = safesave.promote

    def promote(src, dst, **kw):  # promote takes the ticket, returns warnings
        other.save(src)
        return real_promote(src, dst, **kw)

    monkeypatch.setattr(safesave, "promote", promote)
    _fake_excel(monkeypatch, "ok")
    with pytest.raises(ValidationFailed) as ei:
        _autofit(book)
    msg = str(ei.value)
    assert "does not match the copy" in msg, msg
    assert openpyxl.load_workbook(book).active["A1"].value == "someone else's"
    slot = safesave.slot_dir(book) / safesave.PREV_SLOT
    assert _sha(slot) == before
    assert _beside(book) == [] and _temp_left() == []


# ---------------------------------------------- the password route


@pytest.mark.parametrize("backup", [False, True])
def test_an_encryption_that_did_not_take_never_replaces_the_workbook(book, monkeypatch, backup):
    """With backup=false the plain copy used to replace the workbook and the call
    said to check the file."""
    before, slots = _sha(book), _slots(book)
    _fake_excel(monkeypatch, "plain")
    with pytest.raises(ValidationFailed) as ei:
        comtier.com_save_with_password(book, "secret-1", backup=backup)
    _untouched(book, before, slots)
    msg = str(ei.value)
    assert "password did not take" in msg and "book.xlsx was not changed" in msg, msg


def test_an_encrypted_copy_goes_in_place(book, monkeypatch):
    _fake_excel(monkeypatch, "encrypted")
    res = comtier.com_save_with_password(book, "secret-1")
    assert res["ok"] and res["changed"] == {"encryption": "applied"}
    assert Path(book).read_bytes()[:8] == CFB_MAGIC
    assert _beside(book) == [] and _temp_left() == []


@pytest.mark.parametrize("backup", [False, True])
def test_a_broken_decrypted_copy_never_replaces_the_workbook(book, monkeypatch, backup):
    before, slots = _sha(book), _slots(book)
    _fake_excel(monkeypatch, "garbage")
    with pytest.raises(ValidationFailed) as ei:
        comtier.com_save_with_password(book, "", current_password="old", backup=backup)
    _untouched(book, before, slots)
    msg = str(ei.value)
    assert "structural verify" in msg and "book.xlsx was not changed" in msg, msg


def test_a_good_decrypted_copy_goes_in_place_verified(book, monkeypatch):
    _fake_excel(monkeypatch, "ok")
    res = comtier.com_save_with_password(book, "", current_password="old")
    assert res["ok"] and res["verified"] and res["changed"] == {"encryption": "removed"}
    assert openpyxl.load_workbook(book).active["A1"].value == "edited by Excel"
    assert _beside(book) == [] and _temp_left() == []


def test_no_check_failure_is_reported_as_a_restore(book, monkeypatch):
    """Nothing was replaced, so nothing is said to be restored."""
    _fake_excel(monkeypatch, "garbage")
    with pytest.raises(XlMcpError) as ei:
        _autofit(book)
    assert "restored" not in str(ei.value)
