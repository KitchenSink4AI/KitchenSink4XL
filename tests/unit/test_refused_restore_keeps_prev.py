"""A refused restore or overwrite keeps the undo it was working from.

restore_backup rotated prev from the current workbook BEFORE its swap, and
deleted its copy of the backup when the swap was refused: prev then held the
current content, the version being restored was gone, and a retry of
"restore prev" answered ok and changed nothing (MEASURED; in a folder
that refuses the swap, every restore did this). The
save path's RotationTicket now serves restore and overwrite too: the
rotation is staged, lands inside the promote after the swap, and is dropped
when the swap fails. A restore that would change nothing is refused.
"""
from __future__ import annotations

import ctypes
import io
import json
import os
import subprocess
import sys
import time
from ctypes import wintypes
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.core import safesave, swapwait
from xlsx_mcp.core.errors import WorkbookLocked, XlMcpError
from xlsx_mcp.ops import backups, cells, lifecycle

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Windows sharing modes")


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


def _hold_without_share_delete(path):
    """Open `path` the way a program that does not share delete does."""
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                                wintypes.HANDLE]
    h = k32.CreateFileW(str(path), 0x80000000, 0x1 | 0x2, None, 3, 0, None)
    assert h and h != ctypes.c_void_p(-1).value
    return h, k32


@pytest.fixture
def v1_v2_v3(tmp_path):
    p = tmp_path / "Book.xlsx"
    _book(p, "v1")
    cells.set_cell(str(p), "A1", "v2")
    cells.set_cell(str(p), "A1", "v3")
    prev = safesave.slot_dir(p) / "prev.xlsx"
    assert _a1(p) == "v3" and _a1(prev) == "v2"
    return p, prev


def test_a_restore_refused_at_its_swap_keeps_prev_and_says_why(v1_v2_v3, monkeypatch):
    p, prev = v1_v2_v3
    # The holder arrives after restore's own check, at the swap itself.
    monkeypatch.setattr(backups, "_refuse_if_excel_locked", lambda doc: [])
    h, k32 = _hold_without_share_delete(p)
    try:
        with pytest.raises(WorkbookLocked) as info:
            backups.manage_backups("restore", path=str(p), source="prev")
    finally:
        k32.CloseHandle(h)
    assert "PermissionError" not in str(info.value) and "WinError" not in str(info.value)
    assert _a1(p) == "v3"
    assert _a1(prev) == "v2", "the refused restore destroyed the version it was restoring"
    assert not [n for n in os.listdir(safesave.slot_dir(p)) if n.startswith(".slot-")]
    r = backups.manage_backups("restore", path=str(p), source="prev")
    assert r["prev_rotated"] is True and _a1(p) == "v2" and _a1(prev) == "v3"


def test_a_restore_that_would_change_nothing_is_refused(v1_v2_v3):
    p, prev = v1_v2_v3
    backups.manage_backups("restore", path=str(p), source="prev")  # now v2, prev v3
    backups.manage_backups("restore", path=str(p), source="prev")  # now v3, prev v2
    before = p.read_bytes()
    import shutil
    shutil.copyfile(p, prev)  # prev now holds exactly what the workbook holds
    with pytest.raises(XlMcpError) as info:
        backups.manage_backups("restore", path=str(p), source="prev")
    assert getattr(info.value, "code", None) == "CONFLICT"
    assert "Nothing to restore" in str(info.value)
    assert p.read_bytes() == before


def test_an_overwrite_refused_at_its_swap_keeps_prev(tmp_path):
    dst = tmp_path / "Dest.xlsx"
    _book(dst, "d1")
    cells.set_cell(str(dst), "A1", "d2")
    prev = safesave.slot_dir(dst) / "prev.xlsx"
    assert _a1(prev) == "d1"
    src = tmp_path / "src.xlsx"
    _book(src, "copied")
    h, k32 = _hold_without_share_delete(dst)
    try:
        with pytest.raises(Exception):
            lifecycle.copy_workbook(str(src), str(dst), overwrite=True)
    finally:
        k32.CloseHandle(h)
    assert _a1(dst) == "d2" and _a1(prev) == "d1", "a refused overwrite rotated prev"
    r = lifecycle.copy_workbook(str(src), str(dst), overwrite=True)
    assert r["backup"] == "prev" and _a1(dst) == "copied" and _a1(prev) == "d2"


_VICTIM_AFTER_SWAP = r'''
import os, sys, time
from xlsx_mcp.core import safesave
from xlsx_mcp.ops import cells
book, mark = sys.argv[1], sys.argv[2]
real = safesave._replace_file()

def whole_swap_then_stop(dst, src, bak, flags, a, b):
    os.rename(dst, bak)          # both moves done ...
    os.rename(src, dst)
    open(mark, "w").close()      # ... and stopped before the clean-up and the backup
    time.sleep(120)
    return 1

safesave._REPLACE_FILE = (whole_swap_then_stop, real[1], real[2], real[3])
cells.set_cell(book, "A1", "v3")
'''


def test_a_save_stopped_after_its_swap_still_ends_with_the_right_undo(tmp_path, monkeypatch):
    journal = tmp_path / "journal"
    monkeypatch.setenv("KS4XL_SAVE_JOURNAL_DIR", str(journal))
    p = tmp_path / "Book.xlsx"
    _book(p, "v1")
    cells.set_cell(str(p), "A1", "v2")
    mark = tmp_path / "after-the-swap"
    victim = tmp_path / "victim.py"
    victim.write_text(_VICTIM_AFTER_SWAP, encoding="utf-8")
    env = dict(os.environ, PYTHONPATH=os.pathsep.join(sys.path))
    proc = subprocess.Popen([sys.executable, "-X", "utf8", str(victim), str(p), str(mark)],
                            env=env, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    try:
        deadline = time.monotonic() + 60
        while not mark.exists():
            assert proc.poll() is None and time.monotonic() < deadline
            time.sleep(0.05)
    finally:
        proc.kill()
        proc.wait(30)
    prev = safesave.slot_dir(p) / "prev.xlsx"
    assert _a1(p) == "v3" and _a1(prev) == "v1", "the backup landed before the stop"
    entry = json.loads(next(journal.glob("*.json")).read_text("utf-8"))
    assert entry["staged"], "the journal did not record the staged backup"
    assert swapwait.asides(p)
    # The next write lock lands the staged backup and tidies the old copy.
    with safesave.write_lock(p):
        pass
    assert _a1(prev) == "v2", "the undo of the stopped save was lost"
    assert not swapwait.asides(p)
    assert not list(journal.glob("*.json"))
