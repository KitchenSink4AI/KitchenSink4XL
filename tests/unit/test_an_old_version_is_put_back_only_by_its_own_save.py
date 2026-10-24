"""An old version a save left aside is put back only when that save did not finish.

The old version still aside was moved back by name although its swap had
finished (so the user's rename of the workbook was undone by recreating it
from the old version), or although the path held newer content (so it stayed
forever and came back after the user deleted the workbook). And a create or
copy took the workbook's name while its old version could not be put back
(a program had it open), stranding the user's work. Now only a swap that did
not finish is put back; any other old version is removed when a backup slot
holds it, or kept beside the workbook under a visible name; and a new file
never takes the name of a workbook that is held aside. No Excel here.
"""

from __future__ import annotations

import ctypes
import io
import json
import os
import sys
import time
import uuid
from ctypes import wintypes
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp import envelope
from xlsx_mcp.core import safesave, swapwait
from xlsx_mcp.core.errors import XlMcpError
from xlsx_mcp.ops import cells, lifecycle

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Windows file semantics")


@pytest.fixture(autouse=True)
def journal(tmp_path, monkeypatch):
    folder = tmp_path / "journal"
    monkeypatch.setenv("KS4XL_SAVE_JOURNAL_DIR", str(folder))
    return folder


def _book(p, v):
    wb = openpyxl.Workbook()
    wb.active["A1"] = v
    wb.save(p)


def _a1(p):
    wb = openpyxl.load_workbook(io.BytesIO(Path(p).read_bytes()), read_only=True)
    try:
        return wb.active["A1"].value
    finally:
        wb.close()


def _gone_server(entry):
    """Make a written entry look like one a stopped server left."""
    record = json.loads(Path(entry).read_text("utf-8"))
    record.update(pid=2147483640, token="a-server-that-is-gone", pid_created=1.0,
                  time=time.time() - 7200)
    Path(entry).write_text(json.dumps(record), "utf-8")


def _stopped_save(tmp_path, *, finished, backup=True, v_old="v1", v_new="v2"):
    """Book.xlsx at v_old (with its backups when backup), then a save of v_new
    written down in the journal and stopped: mid-swap (the workbook aside,
    the new version beside it) or after the swap (only the clean-up cut)."""
    book = tmp_path / "Book.xlsx"
    if backup:
        _book(book, "v0")
        cells.set_cell(str(book), "A1", v_old)  # slots exist: prev and anchor hold v0
    else:
        _book(book, v_old)
    ticket = safesave.prepare_rotation(book) if backup else None
    new = tmp_path / f"{safesave.WRITE_TMP_PREFIX}{uuid.uuid4().hex}.xlsx"
    _book(new, v_new)
    aside = tmp_path / swapwait.aside_name("Book.xlsx", uuid.uuid4().hex[:12])
    entry = safesave._journal_open(str(book), str(aside), str(new), ticket)
    assert entry is not None
    safesave._journal_close(entry, keep=True)
    os.rename(book, aside)
    if finished:
        os.rename(new, book)
    _gone_server(entry)
    return book, aside, new, entry


# ------------------------------------------------- only an unfinished swap


def test_a_finished_swap_is_not_recreated_after_the_user_renames_the_workbook(tmp_path):
    book, aside, _new, _e = _stopped_save(tmp_path, finished=True)
    renamed = tmp_path / "Book renamed by the user.xlsx"
    os.rename(book, renamed)
    safesave.recover_interrupted_saves()
    assert not book.exists(), "a start recreated the workbook from an old version"
    assert _a1(renamed) == "v2" and not aside.exists()
    with safesave.write_lock(book):  # and no later write lock does either
        pass
    assert not book.exists()


def test_without_backups_the_old_version_is_kept_under_a_visible_name(tmp_path):
    book, aside, _new, _e = _stopped_save(tmp_path, finished=True, backup=False)
    os.remove(book)  # the user deletes it on purpose
    safesave.recover_interrupted_saves()
    assert not book.exists() and not aside.exists()
    kept = tmp_path / f"Book{safesave.KEPT_SUFFIX}.xlsx"
    assert _a1(kept) == "v1"
    attrs = ctypes.windll.kernel32.GetFileAttributesW(str(kept))
    assert not attrs & 0x2, "the kept copy is hidden"


def test_an_old_version_left_when_the_path_was_reused_is_never_resurrected(tmp_path):
    book, aside, new, _e = _stopped_save(tmp_path, finished=False, backup=False)
    _book(book, "v9 (the user's newer file)")
    with safesave.write_lock(book):
        pass  # settles: v9 stays, the unanswered edit goes
    assert _a1(book) == "v9 (the user's newer file)" and not new.exists()
    assert not aside.exists() and _a1(tmp_path / f"Book{safesave.KEPT_SUFFIX}.xlsx") == "v1"
    os.remove(book)
    lifecycle.create_workbook(str(book))
    assert _a1(book) is None, "an old version came back instead of a new workbook"


def test_an_unfinished_swap_is_still_put_back(tmp_path):
    book, aside, new, _e = _stopped_save(tmp_path, finished=False)
    safesave.recover_interrupted_saves()
    assert _a1(book) == "v1" and not aside.exists() and not new.exists()
    # the workbook is as it was, and so are its backups: prev still holds v0
    assert _a1(safesave.slot_dir(book) / safesave.PREV_SLOT) == "v0"
    assert list(safesave.slot_dir(book).glob(".slot-*.tmp")) == []


def test_two_unfinished_saves_put_back_the_latest(tmp_path, journal):
    book, aside1, new1, entry1 = _stopped_save(tmp_path, finished=False, backup=False,
                                                v_old="v1", v_new="v2")
    _book(book, "v9")
    os.utime(book, (time.time() - 30 * 86400, time.time() - 30 * 86400))  # an old date
    new2 = tmp_path / f"{safesave.WRITE_TMP_PREFIX}{uuid.uuid4().hex}.xlsx"
    _book(new2, "v10")
    aside2 = tmp_path / swapwait.aside_name("Book.xlsx", uuid.uuid4().hex[:12])
    entry2 = safesave._journal_open(str(book), str(aside2), str(new2))
    safesave._journal_close(entry2, keep=True)
    os.rename(book, aside2)
    _gone_server(entry2)
    rec = json.loads(Path(entry2).read_text("utf-8"))
    rec["time"] = time.time() - 60  # began after the first
    Path(entry2).write_text(json.dumps(rec), "utf-8")
    safesave.recover_interrupted_saves()
    assert _a1(book) == "v9"
    assert not aside1.exists() and not aside2.exists()
    assert _a1(tmp_path / f"Book{safesave.KEPT_SUFFIX}.xlsx") == "v1"


def _hold(path):
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p,
                                wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    h = k32.CreateFileW(str(path), 0x80000000, 0x1 | 0x2, None, 3, 0x80, None)
    assert h and h != wintypes.HANDLE(-1).value
    return k32, h


@pytest.mark.parametrize("route", ["create", "copy"])
def test_a_new_file_never_takes_the_name_while_the_workbook_is_held_aside(tmp_path, route):
    book, aside, _new, _e = _stopped_save(tmp_path, finished=False)
    src = tmp_path / "Src.xlsx"
    _book(src, "src")
    k32, h = _hold(aside)  # the user opened the old version in a program
    try:
        with pytest.raises(XlMcpError) as info:
            if route == "create":
                lifecycle.create_workbook(str(book))
            else:
                lifecycle.copy_workbook(str(src), str(book))
    finally:
        k32.CloseHandle(wintypes.HANDLE(h))
    assert str(aside) in str(info.value) and "Nothing was changed" in str(info.value)
    assert envelope.refusal(info.value)["error"]["code"] == "CONFLICT"
    assert not book.exists() and _a1(aside) == "v1"
    with safesave.write_lock(book):  # once let go, the next write lock puts it back
        pass
    assert _a1(book) == "v1"


def test_after_a_finished_swap_whose_backup_holds_it_nothing_is_kept_beside(tmp_path):
    book, aside, _new, _e = _stopped_save(tmp_path, finished=True)
    safesave.recover_interrupted_saves()
    assert _a1(book) == "v2" and not aside.exists()
    assert [p.name for p in tmp_path.iterdir() if safesave.KEPT_SUFFIX in p.name] == []


def test_an_old_version_a_dropped_stale_landing_leaves_is_kept(tmp_path, journal, monkeypatch):
    """The stale staged copy is dropped (another save rotated prev); the old
    version still aside is then held by no slot, so it is kept beside."""
    book, aside, _new, _e = _stopped_save(tmp_path, finished=True)
    monkeypatch.setenv("KS4XL_SAVE_JOURNAL_DIR", str(tmp_path / "another_journal"))
    cells.set_cell(str(book), "A1", "v3")
    monkeypatch.setenv("KS4XL_SAVE_JOURNAL_DIR", str(journal))
    safesave.recover_interrupted_saves()
    assert not aside.exists()
    kept = [p for p in tmp_path.iterdir() if safesave.KEPT_SUFFIX in p.name]
    assert [_a1(p) for p in kept] == ["v1"]
