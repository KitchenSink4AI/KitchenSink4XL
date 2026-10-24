"""A create or copy overwrite refused at the swap answers in words.

copy_workbook and create_workbook with overwrite=true, refused because
another program held the workbook at the swap, surfaced the raw OS text
"[WinError 32] The process cannot access the file because it is being used
by another process.: '<path>'". An edit and a restore refused the same way
answer WORKBOOK_LOCKED in words; so do these now. The workbook and its
backups are as they were. No Excel here: a real holder that does not share
delete, opened exactly at the product's first swap attempt.
"""
from __future__ import annotations

import ctypes
import sys
from ctypes import wintypes

import openpyxl
import pytest

from xlsx_mcp import envelope
from xlsx_mcp.core import safesave
from xlsx_mcp.core.errors import WorkbookLocked
from xlsx_mcp.ops import lifecycle

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Windows sharing modes")


@pytest.fixture(autouse=True)
def journal(tmp_path, monkeypatch):
    monkeypatch.setenv("KS4XL_SAVE_JOURNAL_DIR", str(tmp_path / "journal"))


def _book(p, v):
    wb = openpyxl.Workbook()
    wb.active["A1"] = v
    wb.save(p)


class _HoldAtSwap:
    """Open the workbook without sharing delete at the first ReplaceFileW."""

    def __init__(self):
        self.real = safesave._replace_file()
        self.h = None

    def __enter__(self):
        fn, move, last_error, text = self.real
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.CreateFileW.restype = wintypes.HANDLE
        k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                    ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                                    wintypes.HANDLE]
        self.k32 = k32

        def holding(dst, src, bak, flags, a, b):
            if self.h is None:
                self.h = k32.CreateFileW(dst, 0x80000000, 0x1 | 0x2, None, 3, 0x80, None)
            return fn(dst, src, bak, flags, a, b)
        safesave._REPLACE_FILE = (holding, move, last_error, text)
        return self

    def __exit__(self, *a):
        safesave._REPLACE_FILE = self.real
        if self.h:
            self.k32.CloseHandle(wintypes.HANDLE(self.h))


@pytest.mark.parametrize("route", ["copy", "create"])
def test_a_refused_overwrite_answers_in_words(tmp_path, route):
    book = tmp_path / "Book.xlsx"
    _book(book, "v1")
    src = tmp_path / "Src.xlsx"
    _book(src, "copied")
    before = book.read_bytes()
    with _HoldAtSwap():
        with pytest.raises(WorkbookLocked) as info:
            if route == "copy":
                lifecycle.copy_workbook(str(src), str(book), overwrite=True)
            else:
                lifecycle.create_workbook(str(book), overwrite=True)
    text = str(info.value)
    assert "[WinError" not in text and "Book.xlsx" in text and "another program" in text, text
    assert envelope.classify(info.value) == "WORKBOOK_LOCKED"
    assert book.read_bytes() == before
    assert not (safesave.slot_dir(book) / safesave.PREV_SLOT).exists(), "a refused overwrite cost the undo"
