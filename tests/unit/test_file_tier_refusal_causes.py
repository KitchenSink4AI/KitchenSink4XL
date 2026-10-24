r"""file-tier refusals name their cause.

Three file-tier refusals still read "it may be
open in Excel" with raw system text, or blamed a lock that was not there:

  - a folder that does not let this account create files: "cannot write (it
    may be open in Excel). [Errno 13] Permission denied: '...\.ks4xl-write-...'";
  - another program holding the file at the moment of the swap: "cannot
    replace the file (it may be open in Excel). [WinError 5] ...";
  - a file whose permissions deny reading: "another process holds a lock on
    part of this file (an antivirus scanner, indexer ...)".

Each now names the cause in plain words (core/access.py), as the COM guard
does.
"""
from __future__ import annotations

import getpass
import os
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.core import access
from xlsx_mcp.core.errors import WorkbookLocked, XlMcpError

win_only = pytest.mark.skipif(sys.platform != "win32", reason="Windows permissions")


def _icacls(*a):
    return subprocess.run(["icacls", *a], capture_output=True, text=True,
                          creationflags=0x08000000)


def _book(path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(path)
    return str(path)


def _me():
    return "\\".join(x for x in (os.environ.get("USERDOMAIN", ""), getpass.getuser()) if x)


def _plain(msg):
    assert "may be open in Excel" not in msg and "WinError" not in msg and "Errno" not in msg, msg


@win_only
def test_a_folder_that_refuses_new_files_is_named(tmp_path):
    from xlsx_mcp.ops import cells
    folder = tmp_path / "locked"
    folder.mkdir()
    p = _book(folder / "b.xlsx")
    before = Path(p).read_bytes()
    r = _icacls(str(folder), "/deny", f"{_me()}:(WD)")
    if r.returncode != 0:
        pytest.skip(f"cannot set a deny entry here: {r.stdout} {r.stderr}")
    try:
        with pytest.raises(XlMcpError) as ei:
            cells.set_cell(p, "A1", 2, backup=False)
    finally:
        _icacls(str(folder), "/reset")
    msg = str(ei.value)
    _plain(msg)
    assert "cannot create files in its folder" in msg
    assert Path(p).read_bytes() == before


@win_only
def test_a_file_whose_permissions_deny_reading_is_not_called_locked(tmp_path):
    from xlsx_mcp.ops import cells
    p = _book(tmp_path / "noread.xlsx")
    r = _icacls(p, "/deny", f"{_me()}:(R)")
    if r.returncode != 0:
        pytest.skip(f"cannot set a deny entry here: {r.stdout} {r.stderr}")
    try:
        assert access.read_denied(p) is True
        with pytest.raises(XlMcpError) as ei:
            cells.set_cell(p, "A1", 2)
    finally:
        _icacls(p, "/reset")
    msg = str(ei.value)
    assert "not allowed to read" in msg and "holds a lock" not in msg, msg


@win_only
def test_a_free_file_is_not_read_denied(tmp_path):
    assert access.read_denied(_book(tmp_path / "ok.xlsx")) is False


def test_the_replace_refusal_says_retry_when_the_holder_is_gone(tmp_path):
    msg = access.replace_refusal_text(_book(tmp_path / "gone.xlsx"))
    _plain(msg)
    assert "was not changed" in msg and "Retry" in msg


def test_the_folder_text_names_the_folder(tmp_path):
    msg = access.folder_refuses_text(tmp_path / "x.xlsx")
    assert str(tmp_path) in msg and "copy_workbook" in msg
    _plain(msg)
