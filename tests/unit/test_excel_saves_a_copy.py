"""Excel never replaces the workbook's own file; it writes a copy.

Excel's own save replaced the file itself. A program that opened the
workbook after the check that runs before Excel starts, and held it without
sharing delete (a script or viewer reading it with the C runtime's default
sharing), made that replace fail: Excel put a visible "Save As" dialog on the
screen until the call timed out, and left its edited copy beside the user's
file. Every route now has Excel write a copy (SaveCopyAs) into a private temp
folder; once Excel has closed the workbook, the copy replaces the file through
safesave.promote, the file tier's own route. It keeps the file's permissions,
Mark of the Web and creation time, retries a passing reader, and refuses in
plain words while a holder stays, with the workbook and its backups exactly as
they were and nothing left beside it.

The first tests drive a fake Excel; the last two start the real one (com).
"""
from __future__ import annotations

import ctypes
import hashlib
import os
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.com import session as com_session
from xlsx_mcp.core import safesave
from xlsx_mcp.core.errors import WorkbookLocked, XlMcpError
from xlsx_mcp.ops import comtier

IS_WIN = sys.platform == "win32"
needs_windows = pytest.mark.skipif(not IS_WIN, reason="Windows sharing modes and file identity")
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _sha(p) -> str:
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def _beside(book) -> list:
    return sorted(n for n in os.listdir(os.path.dirname(book))
                  if n not in (os.path.basename(book), safesave.BACKUP_DIR_NAME))


@pytest.fixture
def book(tmp_path, monkeypatch):
    folder = tmp_path / "work"
    folder.mkdir()
    p = str(folder / "book.xlsx")
    wb = openpyxl.Workbook()
    wb.active["A1"] = "before"
    wb.save(p)
    # Excel's copies go to a temp folder of this test, so leftovers can be seen.
    temp = tmp_path / "temp"
    temp.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(temp))
    yield p
    if IS_WIN:
        subprocess.run(["icacls", p, "/reset"], capture_output=True, creationflags=NO_WINDOW)


def _temp_left():
    return sorted(os.listdir(tempfile.gettempdir()))


class _Worker:
    pid = 0
    app = object()


class _Manager:
    def acquire(self):
        return _Worker()

    def release_to_pool(self, w):
        pass


class _FakeWB:
    """A workbook Excel opened read-write. Its copy holds the edit; a plain
    Save, which would let Excel replace the file, is a test failure."""

    def __init__(self, path, copy_fails=False):
        self.path = path
        self.FullName = os.path.abspath(path)
        self.ReadOnly = False
        self.Password = ""
        self.copy_fails = copy_fails
        self.copies = []

    def Save(self):
        raise AssertionError("wb.Save lets Excel replace the file itself")

    def SaveCopyAs(self, copy):
        self.copies.append(copy)
        if self.copy_fails:
            Path(copy).write_bytes(b"half a file")
            raise OSError("the disk went away during the save")
        if self.Password:
            # Excel's encryption: an OLE compound file, not a zip (the copy
            # is checked for it before it replaces the workbook).
            Path(copy).write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 4088)
            return
        wb = openpyxl.load_workbook(self.path)
        wb.active["A1"] = "edited by Excel"
        wb.save(copy)

    def Close(self, save):
        pass


def _fake_excel(monkeypatch, *, on_open=None, copy_fails=False):
    books = []
    monkeypatch.setattr(com_session, "run_com",
                        lambda label, job, timeout=None: job(_Manager()))
    monkeypatch.setattr(com_session, "_assert_hygiene", lambda app: ())
    monkeypatch.setattr(com_session, "_restore_hygiene", lambda app, *a: None)
    monkeypatch.setattr(com_session, "com_retry", lambda fn, label=None, **k: fn())

    def _open(app, p, password=None, **k):
        if on_open is not None:
            on_open(p)  # what another program does once the check has passed
        wb = _FakeWB(p, copy_fails=copy_fails)
        books.append(wb)
        return wb

    monkeypatch.setattr(com_session, "open_workbook", _open)
    return books


def _autofit(book):
    return comtier._run_mutation("autofit(book.xlsx)", book, lambda app, wb: {"autofit": True})


def test_the_copy_replaces_the_workbook_and_nothing_is_left(book, monkeypatch):
    books = _fake_excel(monkeypatch)
    res = _autofit(book)
    assert res["ok"] and res["verified"] and res["backup"] == "prev"
    assert openpyxl.load_workbook(book).active["A1"].value == "edited by Excel"
    slot = safesave.slot_dir(book) / safesave.PREV_SLOT
    assert openpyxl.load_workbook(slot).active["A1"].value == "before"
    copy = books[0].copies[0]
    # Beside the workbook (Excel writes relative links relative
    # to the folder it saves into), and gone once it is in place.
    assert os.path.dirname(copy) == os.path.dirname(os.path.abspath(book)), \
        "Excel's copy was not written beside the workbook"
    assert _beside(book) == []
    assert _temp_left() == []


@needs_windows
def test_a_program_that_opens_the_workbook_after_the_check_gets_a_plain_refusal(book, monkeypatch):
    """The same case without Excel: a reader takes the file once the check
    before Excel has passed and holds it through the save."""
    held = []
    _fake_excel(monkeypatch, on_open=lambda p: held.append(open(p, "rb")))
    before = _sha(book)
    slot = safesave.slot_dir(book) / safesave.PREV_SLOT
    slot_before = _sha(slot) if slot.exists() else None
    try:
        with pytest.raises(WorkbookLocked) as ei:
            _autofit(book)
    finally:
        for h in held:
            h.close()
    msg = str(ei.value)
    assert "in use by another program" in msg, msg
    assert "WinError" not in msg
    assert _sha(book) == before
    assert (_sha(slot) if slot.exists() else None) == slot_before, "the prev slot was burned"
    assert _beside(book) == []
    assert _temp_left() == []
    # Once the other program lets go, the same call goes through.
    _fake_excel(monkeypatch)
    assert _autofit(book)["ok"]
    assert openpyxl.load_workbook(book).active["A1"].value == "edited by Excel"


@needs_windows
def test_the_encrypted_save_goes_through_the_copy_too(book, monkeypatch):
    held = []
    _fake_excel(monkeypatch, on_open=lambda p: held.append(open(p, "rb")))
    before = _sha(book)
    try:
        with pytest.raises(WorkbookLocked):
            comtier.com_save_with_password(book, "secret-1")
    finally:
        for h in held:
            h.close()
    assert _sha(book) == before
    assert _beside(book) == [] and _temp_left() == []


def test_a_copy_that_fails_leaves_the_workbook_and_nothing_else(book, monkeypatch):
    _fake_excel(monkeypatch, copy_fails=True)
    before = _sha(book)
    with pytest.raises(XlMcpError) as ei:
        _autofit(book)
    assert "was not changed" in str(ei.value)
    assert _sha(book) == before
    assert _beside(book) == [] and _temp_left() == []


def _icacls(*a):
    return subprocess.run(["icacls", *a], capture_output=True, text=True,
                          creationflags=NO_WINDOW).stdout


def _explicit_aces(p):
    lines = [ln.strip() for ln in _icacls(p).splitlines() if ln.strip() and "Successfully" not in ln]
    lines[0] = lines[0].replace(p, "").strip()
    return sorted(ln for ln in lines if "(I)" not in ln)


@needs_windows
def test_the_workbook_keeps_its_identity_when_the_copy_replaces_it(book, monkeypatch):
    """Excel's own save kept the file's permissions, Mark of the Web and
    creation time; the copy put in place keeps them as well."""
    _icacls(book, "/grant", "*S-1-5-32-545:(R)")
    _icacls(book, "/deny", "*S-1-5-7:(R)")
    aces = _explicit_aces(book)
    # The two entries just added are there; an elevated account's new files
    # carry a few explicit entries of their own, which must survive too.
    assert {"BUILTIN\\Users:(R)", "NT AUTHORITY\\ANONYMOUS LOGON:(DENY)(R)"} <= set(aces), aces
    with open(book + ":Zone.Identifier", "w") as f:
        f.write("[ZoneTransfer]\r\nZoneId=3\r\n")
    born = os.stat(book).st_birthtime
    _fake_excel(monkeypatch)
    assert _autofit(book)["ok"]
    assert openpyxl.load_workbook(book).active["A1"].value == "edited by Excel"
    assert _explicit_aces(book) == aces
    with open(book + ":Zone.Identifier") as f:
        assert "ZoneId=3" in f.read()
    assert os.stat(book).st_birthtime == born


# ------------------------------------------------------------- real Excel


def _com_ok() -> bool:
    try:
        return bool(com_session.com_available()[0])
    except Exception:  # noqa: BLE001
        return False


def _shut_down_excel():
    try:
        com_session.get_executor().shutdown()
    except Exception:  # noqa: BLE001
        pass
    com_session._EXECUTOR = None


@pytest.mark.com
@needs_windows
@pytest.mark.skipif(not _com_ok(), reason="needs Excel (the product's private hidden pool)")
@pytest.mark.timeout(300)
def test_live_a_reader_after_the_check_gets_no_dialog_and_a_plain_refusal(book):
    """The live case: a reader waits for Excel's owner file (Excel has the
    workbook open read-write, so the check has passed), then holds the
    workbook through the save. Earlier versions put a visible Save As
    dialog on the screen, answered APP_BLOCKED after the timeout and left
    Excel's copy beside the file."""
    folder = os.path.dirname(book)
    before = _sha(book)
    done = threading.Event()
    held = {}

    def reader():
        owner = os.path.join(folder, "~$book.xlsx")
        t0 = time.time()
        while not done.is_set() and time.time() - t0 < 60:
            if os.path.exists(owner):
                try:
                    with open(book, "rb"):
                        held["at"] = round(time.time() - t0, 2)
                        done.wait(40)
                except OSError as e:
                    held["error"] = repr(e)
                return
            time.sleep(0.005)

    th = threading.Thread(target=reader, daemon=True)
    th.start()
    t0 = time.time()
    try:
        with pytest.raises(WorkbookLocked) as ei:
            comtier.com_autofit(book, timeout_seconds=45)
        elapsed = time.time() - t0
    finally:
        done.set()
        th.join(10)
    try:
        assert "at" in held, f"the reader never saw Excel open the workbook: {held}"
        assert elapsed < 30, f"the refusal took {elapsed:.0f}s (Excel waiting on a dialog?)"
        assert "in use by another program" in str(ei.value)
        assert _sha(book) == before
        assert [n for n in _beside(book) if not n.startswith("~$")] == []
        # With the reader gone, the same call goes through.
        res = comtier.com_autofit(book, timeout_seconds=45)
        assert res["ok"] and res["verified"]
        assert _beside(book) == []
    finally:
        _shut_down_excel()


@pytest.mark.com
@needs_windows
@pytest.mark.skipif(not _com_ok(), reason="needs Excel (the product's private hidden pool)")
@pytest.mark.timeout(300)
def test_live_an_excel_save_keeps_the_workbooks_identity(book):
    _icacls(book, "/grant", "*S-1-5-32-545:(R)")
    _icacls(book, "/deny", "*S-1-5-7:(R)")
    aces = _explicit_aces(book)
    with open(book + ":Zone.Identifier", "w") as f:
        f.write("[ZoneTransfer]\r\nZoneId=3\r\n")
    born = os.stat(book).st_birthtime
    try:
        res = comtier.com_autofit(book, timeout_seconds=60)
        assert res["ok"] and res["verified"]
        assert _explicit_aces(book) == aces
        with open(book + ":Zone.Identifier") as f:
            assert "ZoneId=3" in f.read()
        assert os.stat(book).st_birthtime == born
        assert _beside(book) == [] and _temp_left() == []
    finally:
        _shut_down_excel()
