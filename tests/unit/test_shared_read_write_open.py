"""The shared read-write open never
leaves Excel waiting on a dialog, and nothing saves a read-only open.

-    A workbook with a password to open AND a password to modify: the
     encryption hides the password to modify from the file check, so
     com_save_with_password(current_password=...) opened it read-write and
     Excel showed its Password dialog. Every read-write open now hands Excel
     a stand-in password to modify; a file that has one fails at once with
     Excel's wrong-password error instead of asking (measured on Excel 16:
     plain and encrypted-only files open normally with the stand-in). The
     failure is then told apart from a wrong current_password with a
     read-only open, which never asks about a password to modify.
-    Another program holds the workbook open for writing and shares it (no
     ~$ owner file). The old probe shared writing too and passed, Excel
     opened the file read-only, and wb.Save put a "Save As" dialog up. The
     guard now probes the way Excel opens (other writers refused), the open
     refuses a workbook Excel could open only read-only, and no route calls
     wb.Save on a read-only workbook.

These tests drive a fake Excel; none starts the real one.
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.com import session as S
from xlsx_mcp.core.errors import WorkbookLocked, WorkbookProtected

SRC = Path(__file__).resolve().parents[2] / "src" / "xlsx_mcp"


class FakeComError(Exception):
    """Excel's wrong-password refusal, in pywin32's shape."""

    def __init__(self, text="The password you supplied is not correct."):
        super().__init__(-2147352567, "Exception occurred.",
                         (0, "Microsoft Excel", text, None, 0, -2146827284), None)
        self.excepinfo = (0, "Microsoft Excel", text, None, 0, -2146827284)


class FakeWB:
    def __init__(self, app, path, read_only):
        self.app, self.path, self.ReadOnly = app, path, read_only
        # The shared open checks that Excel opened THE file
        self.FullName = os.path.abspath(path) if path else "Book1"
        self.closed = False
        self.saved = 0

    def Close(self, *a, **k):
        self.closed = True

    def Save(self):
        raise AssertionError("wb.Save lets Excel replace the file itself")

    def SaveCopyAs(self, path):
        if self.ReadOnly:
            raise AssertionError("a save of a read-only workbook")
        Path(path).write_bytes(b"PK copy")
        self.saved += 1

    def SaveAs(self, path, fmt=None, *a):
        Path(path).write_bytes(b"PK plain")
        self.app.kinds[os.path.abspath(path)] = "plain"


class FakeWorkbooks:
    def __init__(self, app):
        self.app = app

    def Open(self, path, upd, ro, fmt, pw, wrp, ignore, *rest):
        app = self.app
        # A read-write open also passes Origin, Delimiter and
        # Editable=True (rest); a read-only open passes none of them
        app.opens.append({"path": path, "ro": ro, "pw": pw, "wrp": wrp, "rest": rest})
        if wrp is not None and len(wrp) > 15:
            # Measured: Excel's limit for a password to modify is 15
            # characters, and a longer one fails every open outright.
            raise FakeComError("Open method of Workbooks class failed")
        kind = app.kinds.get(os.path.abspath(path), "plain")
        if kind.startswith("enc") and pw != "o":
            raise FakeComError()
        if not ro:
            if kind.endswith("reserved"):
                if wrp is None:
                    raise AssertionError("Excel shows its Password dialog")
                if wrp != "w":
                    raise FakeComError()
            elif wrp is not None and not app.build_accepts_stand_in:
                raise FakeComError()
        wb = FakeWB(app, path, ro or kind == "in_use")
        app.books.append(wb)
        return wb

    def Add(self):
        wb = FakeWB(self.app, None, False)
        self.app.books.append(wb)
        return wb


class FakeApp:
    def __init__(self, build_accepts_stand_in=True):
        self.kinds = {}
        self.opens = []
        self.books = []
        self.build_accepts_stand_in = build_accepts_stand_in
        self.Workbooks = FakeWorkbooks(self)


@pytest.fixture(autouse=True)
def _fresh_measurement(monkeypatch):
    monkeypatch.setattr(S, "_WRITE_STAND_IN_OK", None)
    # The network check that opens an encrypted workbook with calculation
    # off first has its own tests (test_every_com_open_passes_the_network_
    # check); these count the opens of the password sequence alone.
    monkeypatch.setattr(S, "network_gate", lambda *a, **k: None)


def _cfb_bytes(streams):
    """A minimal OLE compound file naming the given streams (Excel's
    encryption wraps the package in EncryptionInfo and EncryptedPackage)."""
    import struct
    header = bytearray(512)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    for off, fmt, val in ((26, "<H", 3), (28, "<H", 0xFFFE), (30, "<H", 9),
                          (32, "<H", 6), (44, "<i", 1), (48, "<i", 0)):
        struct.pack_into(fmt, header, off, val)
    for off in range(76, 512, 4):
        struct.pack_into("<i", header, off, -1)
    sector = bytearray(512)
    for i, (name, etype) in enumerate([("Root Entry", 5)] + [(s, 2) for s in streams]):
        e = bytearray(128)
        raw = name.encode("utf-16-le")
        e[0:len(raw)] = raw
        struct.pack_into("<H", e, 64, len(raw) + 2)
        e[66], e[67] = etype, 1
        for off in (68, 72, 76):
            struct.pack_into("<i", e, off, -1)
        sector[i * 128:(i + 1) * 128] = e
    return bytes(header) + bytes(sector)


def _file(tmp_path, name, app=None, kind="plain"):
    p = tmp_path / name
    if kind.startswith("enc"):
        p.write_bytes(_cfb_bytes(["EncryptionInfo", "EncryptedPackage"]))
    else:
        wb = openpyxl.Workbook()
        wb.active["A1"] = 1
        wb.save(p)
    if app is not None:
        app.kinds[os.path.abspath(p)] = kind
    return str(p)



def test_a_plain_read_write_open_carries_the_stand_in_and_opens(tmp_path):
    app = FakeApp()
    p = _file(tmp_path, "plain.xlsx", app)
    wb = S.open_workbook(app, p)
    assert wb.ReadOnly is False
    assert len(app.opens) == 1
    assert app.opens[0]["ro"] is False and app.opens[0]["wrp"] not in (None, "")


def test_encrypted_with_a_password_to_modify_is_refused_without_a_dialog(tmp_path):
    app = FakeApp()
    p = _file(tmp_path, "enc_wr.xlsx", app, "enc_reserved")
    with pytest.raises(WorkbookProtected) as ei:
        S.open_workbook(app, p, password="o")
    msg = str(ei.value)
    assert "password to modify" in msg and "Nothing was changed" in msg
    assert all(b.closed for b in app.books)  # the read-only probe was closed


def test_encrypted_without_a_password_to_modify_opens(tmp_path):
    app = FakeApp()
    p = _file(tmp_path, "enc.xlsx", app, "enc")
    wb = S.open_workbook(app, p, password="o")
    assert wb.ReadOnly is False and len(app.opens) == 1


def test_a_wrong_current_password_still_says_so(tmp_path):
    app = FakeApp()
    p = _file(tmp_path, "enc_wr.xlsx", app, "enc_reserved")
    with pytest.raises(FakeComError):
        S.open_workbook(app, p, password="bad")


def test_reserved_plain_file_the_file_check_missed_is_refused_by_excel(tmp_path, monkeypatch):
    """The stand-in is an independent backstop: even if the file check
    misses a form, Excel refuses at once instead of asking."""
    from xlsx_mcp.core import hazard
    monkeypatch.setattr(hazard, "write_reservation", lambda p: None)
    app = FakeApp()
    p = _file(tmp_path, "missed.xlsx", app, "reserved")
    with pytest.raises(WorkbookProtected) as ei:
        S.open_workbook(app, p)
    assert "Excel reports" in str(ei.value)


def test_an_excel_that_rejects_the_stand_in_falls_back_for_plain_files(tmp_path):
    app = FakeApp(build_accepts_stand_in=False)
    p = _file(tmp_path, "plain.xlsx", app)
    wb = S.open_workbook(app, p)
    assert wb.ReadOnly is False
    assert S._WRITE_STAND_IN_OK is False
    # the measurement is kept: the next open goes straight to the old form
    n = len(app.opens)
    S.open_workbook(app, _file(tmp_path, "plain2.xlsx", app))
    assert len(app.opens) == n + 1 and app.opens[-1]["wrp"] is None


def test_an_excel_that_rejects_the_stand_in_refuses_an_encrypted_file(tmp_path):
    """Encryption hides the setting and the stand-in cannot be used, so a
    password to modify cannot be excluded: refuse rather than risk the
    dialog."""
    app = FakeApp(build_accepts_stand_in=False)
    p = _file(tmp_path, "enc.xlsx", app, "enc")
    with pytest.raises(WorkbookProtected) as ei:
        S.open_workbook(app, p, password="o")
    assert "rule out" in str(ei.value)


def test_the_stand_in_fits_excels_limit_for_a_password_to_modify():
    """Measured on Excel 16: 15 characters work as a stand-in; 16 or more
    fail every open outright, plain files included."""
    assert 1 <= len(S._NO_WRITE_PASSWORD_SENTINEL) <= 15


def test_once_the_stand_in_is_known_refused_an_ole_container_never_opens(tmp_path, monkeypatch):
    monkeypatch.setattr(S, "_WRITE_STAND_IN_OK", False)
    app = FakeApp(build_accepts_stand_in=False)
    p = _file(tmp_path, "enc_wr.xlsb", app, "enc_reserved")
    with pytest.raises(WorkbookProtected) as ei:
        S.open_workbook(app, p, password="o")
    assert "rule out" in str(ei.value)
    assert app.opens == []


def test_read_only_opens_are_unchanged(tmp_path):
    app = FakeApp()
    p = _file(tmp_path, "r.xlsx", app, "reserved")
    wb = S.open_workbook(app, p, read_only=True)
    assert wb.ReadOnly is True and app.opens[0]["wrp"] is None



def test_a_workbook_excel_could_open_only_read_only_is_refused_and_closed(tmp_path):
    app = FakeApp()
    p = _file(tmp_path, "held.xlsx", app, "in_use")
    with pytest.raises(WorkbookLocked) as ei:
        S.open_workbook(app, p)
    assert "only read-only" in str(ei.value)
    assert "open in Excel" not in str(ei.value)
    assert app.books and all(b.closed and b.saved == 0 for b in app.books)


def test_save_workbook_never_saves_a_read_only_workbook(tmp_path):
    app = FakeApp()
    target = str(tmp_path / "x.xlsx")
    ro = FakeWB(app, target, True)
    with pytest.raises(WorkbookLocked):
        S.save_workbook(ro, label="save x.xlsx", path=target)
    assert ro.saved == 0
    rw = FakeWB(app, target, False)
    copy = S.save_workbook(rw, label="save x.xlsx", path=target)
    try:
        assert rw.saved == 1 and os.path.isfile(copy)
        # Beside the workbook, never the workbook itself
        assert os.path.dirname(copy) == str(tmp_path) and copy != target
    finally:
        S.discard_saved(copy)
    assert not os.path.exists(copy)


def _calls(pattern):
    found = []
    for f in SRC.rglob("*.py"):
        text = f.read_text(encoding="utf-8")
        for m in re.finditer(pattern, text):
            line = text.count("\n", 0, m.start()) + 1
            found.append(f"{f.relative_to(SRC).as_posix()}:{line}")
    return found


def test_excel_never_saves_over_the_workbook_itself():
    """No route calls wb.Save, which lets Excel replace the file
    itself (and show Save As when another program holds it). The one save
    is SaveCopyAs, inside save_workbook."""
    assert _calls(r"\.Save\s*[(,)]") == []
    s = (SRC / "com" / "session.py").read_text(encoding="utf-8")
    body = s[s.index("def save_workbook"):]
    body = body[:body.index("\ndef ", 1)]
    allowed = {f"com/session.py:{s.count(chr(10), 0, s.index(body)) + 1 + body.count(chr(10), 0, m.start())}"
               for m in re.finditer(r"\.SaveCopyAs\s*[(,)]", body)}
    assert allowed and set(_calls(r"\.SaveCopyAs\s*[(,)]")) == allowed


# ----------------------------------------------------- the guard's probes

win_only = pytest.mark.skipif(sys.platform != "win32", reason="Windows sharing modes")


def _hold(path, share_write: bool, write: bool = True, share_delete: bool = False):
    import ctypes
    from ctypes import wintypes
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    k32.CreateFileW.restype = wintypes.HANDLE
    access = 0x80000000 | (0x40000000 if write else 0)
    share = 0x1 | (0x2 if share_write else 0) | (0x4 if share_delete else 0)
    h = k32.CreateFileW(str(path), access, share, None, 3, 0x80, None)
    assert h and h != ctypes.c_void_p(-1).value
    return lambda: k32.CloseHandle(wintypes.HANDLE(h))


@win_only
def test_guard_refuses_a_file_another_program_holds_for_writing(tmp_path, monkeypatch):
    from xlsx_mcp.ops import comtier
    monkeypatch.setattr(S, "run_com", lambda *a, **k: (_ for _ in ()).throw(
        AssertionError("Excel was started for a file another program holds")))
    p = _file(tmp_path, "held.xlsx")
    before = Path(p).read_bytes()
    release = _hold(p, share_write=True)
    try:
        with pytest.raises(WorkbookLocked) as ei:
            comtier.com_autofit(p)
    finally:
        release()
    msg = str(ei.value)
    assert "in use by another program" in msg and "open in Excel" not in msg
    assert Path(p).read_bytes() == before


@win_only
def test_guard_says_open_in_excel_only_when_excels_owner_file_is_there(tmp_path):
    p = _file(tmp_path, "book.xlsx")
    (tmp_path / "~$book.xlsx").write_bytes(b"owner")
    release = _hold(p, share_write=False)
    try:
        with pytest.raises(WorkbookLocked) as ei:
            S.guard_target_closed(p)
    finally:
        release()
    assert "open in Excel" in str(ei.value)


@win_only
def test_guard_names_the_read_only_attribute(tmp_path):
    import stat
    p = _file(tmp_path, "ro.xlsx")
    os.chmod(p, stat.S_IREAD)
    try:
        with pytest.raises(WorkbookLocked) as ei:
            S.guard_target_closed(p)
    finally:
        os.chmod(p, stat.S_IREAD | stat.S_IWRITE)
    msg = str(ei.value)
    assert "read-only attribute" in msg and "open in Excel" not in msg


@win_only
def test_guard_lets_a_reader_that_shares_writing_and_delete_through(tmp_path):
    """A program that only READS the file and lets others write and delete
    it does not stop Excel opening it for editing or replacing it on save,
    so it must not stop the guard. A reader that does NOT
    share delete (the C runtime's default) is refused, because Excel's save
    cannot replace the file under it and shows a visible Save As dialog
    (tests/unit/test_write_access_probes.py)."""
    p = _file(tmp_path, "read.xlsx")
    release = _hold(p, share_write=True, write=False, share_delete=True)
    try:
        assert S.guard_target_closed(p) == []
    finally:
        release()
