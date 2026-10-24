r"""The write probes.

-     A program that holds the workbook open for READING and shares reading
      and writing but not delete (the C runtime's default, so Python's own
      open(path, "rb"), many viewers and scripts) passed the deny-writers
      probe, because it does not write. Excel then edited the workbook, could
      not replace the original on save and showed a visible "Save As"
      dialog; the file tier failed at the replace with raw system text. Both
      probes now ask for DELETE access, which is what replacing the file
      needs, so exactly that holder is refused before anything starts.
-     A network path over 240 characters was given the drive form of the
      long-path prefix (\\?\\\server\share\...), which names no file, so the
      probe failed and the edit was refused as "in use".

No test here starts Excel: the COM routes run with run_com as a tripwire.
"""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.com import session as S
from xlsx_mcp.core import access
from xlsx_mcp.core.errors import WorkbookLocked

win_only = pytest.mark.skipif(sys.platform != "win32", reason="Windows sharing modes")

GENERIC_READ, GENERIC_WRITE = 0x80000000, 0x40000000
SHARE_READ, SHARE_WRITE, SHARE_DELETE = 0x1, 0x2, 0x4


def _book(folder, name="book.xlsx") -> str:
    p = Path(folder) / name
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(p)
    return str(p)


def _hold(path, access_mask, share):
    import ctypes
    from ctypes import wintypes
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    k32.CreateFileW.restype = wintypes.HANDLE
    h = k32.CreateFileW(str(path), access_mask, share, None, 3, 0x80, None)
    assert h and h != ctypes.c_void_p(-1).value, ctypes.get_last_error()
    return lambda: k32.CloseHandle(wintypes.HANDLE(h))


@pytest.fixture
def no_excel(monkeypatch):
    def _trip(*a, **k):
        raise AssertionError("Excel was started on a workbook it could not save")
    monkeypatch.setattr(S, "run_com", _trip)


def _leftovers(folder):
    return [n for n in os.listdir(folder) if n.startswith(".ks4xl-write-")]


# ------------------------------------------------------------- COM


@win_only
def test_a_reader_that_does_not_share_delete_is_refused_before_excel(tmp_path, no_excel):
    from xlsx_mcp.ops import comtier
    p = _book(tmp_path)
    before = Path(p).read_bytes()
    release = _hold(p, GENERIC_READ, SHARE_READ | SHARE_WRITE)
    try:
        with pytest.raises(WorkbookLocked) as ei:
            comtier.com_autofit(p)
    finally:
        release()
    msg = str(ei.value)
    assert "in use by another program" in msg and "open in Excel" not in msg
    assert Path(p).read_bytes() == before


@win_only
def test_pythons_own_open_for_reading_is_that_reader(tmp_path, no_excel):
    from xlsx_mcp.ops import comtier
    p = _book(tmp_path)
    with open(p, "rb"):
        with pytest.raises(WorkbookLocked):
            comtier.com_autofit(p)
        with pytest.raises(WorkbookLocked):
            comtier.com_save_with_password(p, "secret")


@win_only
def test_a_reader_that_shares_delete_passes_the_guard(tmp_path):
    p = _book(tmp_path)
    release = _hold(p, GENERIC_READ, SHARE_READ | SHARE_WRITE | SHARE_DELETE)
    try:
        assert S.guard_target_closed(p) == []
    finally:
        release()


@win_only
def test_no_holder_passes_the_guard(tmp_path):
    assert S.guard_target_closed(_book(tmp_path)) == []


@win_only
def test_a_writer_that_shares_writing_is_still_refused(tmp_path):
    p = _book(tmp_path)
    release = _hold(p, GENERIC_READ | GENERIC_WRITE, SHARE_READ | SHARE_WRITE | SHARE_DELETE)
    try:
        with pytest.raises(WorkbookLocked):
            S.guard_target_closed(p)
    finally:
        release()


# -------------------------------------------------------- file tier


@win_only
def test_the_file_tier_refuses_that_reader_in_plain_words(tmp_path):
    from xlsx_mcp.ops import cells
    p = _book(tmp_path)
    before = Path(p).read_bytes()
    release = _hold(p, GENERIC_READ, SHARE_READ | SHARE_WRITE)
    try:
        with pytest.raises(WorkbookLocked) as ei:
            cells.set_cell(p, "A1", 2)
    finally:
        release()
    msg = str(ei.value)
    assert "in use by another program" in msg
    assert "WinError" not in msg and "may be open in Excel" not in msg
    assert Path(p).read_bytes() == before
    assert _leftovers(tmp_path) == []




@win_only
def test_a_long_network_path_takes_the_unc_long_form():
    long = "\\\\server\\share\\" + "d" * 240 + "\\book.xlsx"
    assert access.long_path(long) == "\\\\?\\UNC\\server\\share\\" + "d" * 240 + "\\book.xlsx"


@win_only
def test_a_long_drive_path_takes_the_drive_long_form():
    long = "C:\\" + "d" * 240 + "\\book.xlsx"
    assert access.long_path(long) == "\\\\?\\C:\\" + "d" * 240 + "\\book.xlsx"


@win_only
def test_short_and_already_prefixed_paths_are_left_alone():
    short = os.path.abspath("book.xlsx")
    assert access.long_path(short) == short
    pre = "\\\\?\\UNC\\server\\share\\" + "d" * 240 + "\\book.xlsx"
    assert access.long_path(pre) == pre


@win_only
def test_the_probe_opens_a_real_long_network_path(tmp_path):
    """Through the administrative share of this machine, when it is
    reachable: a 250-character UNC path to a free workbook passes both
    probes (it was refused as "in use")."""
    drive = str(tmp_path)[0]
    unc_root = f"\\\\localhost\\{drive}$" + str(tmp_path)[2:]
    if not os.path.isdir(unc_root):
        pytest.skip("the administrative share is not reachable here")
    deep = unc_root
    while len(deep) < 236:
        deep = os.path.join(deep, "d" * min(60, 236 - len(deep)))
    target = os.path.join(deep, "book.xlsx")
    assert len(target) > 240
    os.makedirs(access.long_path(deep), exist_ok=True)
    shutil.copy(_book(tmp_path), access.long_path(target))
    try:
        assert access.write_block(target) is None
        assert access.write_block(target, deny_writers=True) is None
        assert S.guard_target_closed(target) == []
    finally:
        shutil.rmtree(access.long_path(os.path.join(unc_root, "d" * 60)), ignore_errors=True)
