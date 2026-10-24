"""A read-only attribute, a file whose permissions deny
writing, or another program holding the file is never described as "open in
Excel". The COM routes said "open in Excel or locked by another process" for
all three; the file tier said "open in Excel (or another
process holds it)" for the last two. "Open in Excel" now appears only when
Excel's owner file (~$name) is beside the workbook.
"""
from __future__ import annotations

import os
import stat
import sys

import openpyxl
import pytest

from xlsx_mcp.core import access
from xlsx_mcp.core.errors import WorkbookLocked

win_only = pytest.mark.skipif(sys.platform != "win32", reason="Windows sharing modes")


def _book(path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(path)
    return str(path)


def _hold_denying_writers(path):
    import ctypes
    from ctypes import wintypes
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    k32.CreateFileW.restype = wintypes.HANDLE
    h = k32.CreateFileW(str(path), 0x80000000, 0x1, None, 3, 0x80, None)
    assert h and h != ctypes.c_void_p(-1).value
    return lambda: k32.CloseHandle(wintypes.HANDLE(h))


def test_the_texts_name_their_cause(tmp_path):
    p = tmp_path / "b.xlsx"
    assert "open in Excel" not in access.refusal_text(p, access.IN_USE)
    assert "in use by another program" in access.refusal_text(p, access.IN_USE)
    assert "not allowed to change the file" in access.refusal_text(p, access.ACCESS_DENIED)
    assert "open in Excel" not in access.refusal_text(p, access.ACCESS_DENIED)
    assert "read-only attribute" in access.refusal_text(p, access.READ_ONLY_ATTRIBUTE)
    (tmp_path / "~$b.xlsx").write_bytes(b"owner")
    assert "open in Excel" in access.refusal_text(p, access.IN_USE)


def test_file_tier_names_a_permissions_refusal(tmp_path, monkeypatch):
    from xlsx_mcp.ops import cells
    p = _book(tmp_path / "acl.xlsx")
    monkeypatch.setattr(access, "write_block", lambda path, deny_writers=False: access.ACCESS_DENIED)
    with pytest.raises(WorkbookLocked) as ei:
        cells.set_cell(p, "A1", 2)
    assert "not allowed to change the file" in str(ei.value)
    assert "open in Excel" not in str(ei.value)


@win_only
def test_file_tier_names_another_program_holding_the_file(tmp_path):
    from xlsx_mcp.ops import cells
    p = _book(tmp_path / "held.xlsx")
    release = _hold_denying_writers(p)
    try:
        with pytest.raises(WorkbookLocked) as ei:
            cells.set_cell(p, "A1", 2)
    finally:
        release()
    assert "in use by another program" in str(ei.value)
    assert "open in Excel" not in str(ei.value)


@win_only
def test_restore_names_the_cause(tmp_path):
    from xlsx_mcp.ops import cells, backups
    p = _book(tmp_path / "r.xlsx")
    cells.set_cell(p, "A1", 2)  # makes a prev slot
    os.chmod(p, stat.S_IREAD)
    try:
        with pytest.raises(WorkbookLocked) as ei:
            backups.restore_backup(p, "prev")
    finally:
        os.chmod(p, stat.S_IREAD | stat.S_IWRITE)
    assert "read-only attribute" in str(ei.value)
    assert "open in Excel" not in str(ei.value)


@win_only
def test_com_guard_names_a_deny_write_handle_without_excel(tmp_path):
    from xlsx_mcp.com import session
    p = _book(tmp_path / "dw.xlsx")
    release = _hold_denying_writers(p)
    try:
        with pytest.raises(WorkbookLocked) as ei:
            session.guard_target_closed(p)
    finally:
        release()
    assert "in use by another program" in str(ei.value)
    assert "open in Excel" not in str(ei.value)
