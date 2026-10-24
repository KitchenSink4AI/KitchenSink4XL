"""diagnose_workbook waits out another save's swap, as every read does.

Every route that reads a workbook waits out the few milliseconds another
save's swap leaves its path empty; diagnose_workbook did not, and answered
ok with readable false, "file not found", during a save. It now waits, and
when a save that did not finish left the workbook aside, it says where the
workbook is (and moves nothing). No Excel here.
"""
from __future__ import annotations

import os
import threading
import time

import openpyxl
import pytest

from xlsx_mcp.core import swapwait
from xlsx_mcp.ops import lifecycle


@pytest.fixture(autouse=True)
def journal(tmp_path, monkeypatch):
    monkeypatch.setenv("KS4XL_SAVE_JOURNAL_DIR", str(tmp_path / "journal"))


def _book(p):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "=1+1"
    wb.save(p)


def test_a_swap_in_progress_is_waited_out(tmp_path):
    book = tmp_path / "Book.xlsx"
    _book(book)
    aside = tmp_path / swapwait.aside_name("Book.xlsx", "0123456789ab")
    os.replace(book, aside)

    def back():
        time.sleep(0.3)
        os.replace(aside, book)
    th = threading.Thread(target=back)
    th.start()
    try:
        r = lifecycle.diagnose_workbook(str(book))
    finally:
        th.join()
    assert r["readable"] is True and r["health"]["formula_cells"] == 1


def test_a_workbook_a_stopped_save_left_aside_is_named_not_moved(tmp_path):
    book = tmp_path / "Book.xlsx"
    _book(book)
    aside = tmp_path / swapwait.aside_name("Book.xlsx", "0123456789ab")
    os.replace(book, aside)
    t0 = time.monotonic()
    r = lifecycle.diagnose_workbook(str(book))
    assert r["readable"] is False
    assert str(aside) in r["error"] and "content is safe there" in r["error"]
    assert aside.exists() and not book.exists(), "a read moved a file"
    assert time.monotonic() - t0 < 3.0


def test_a_workbook_that_is_not_there_is_answered_at_once(tmp_path):
    t0 = time.monotonic()
    r = lifecycle.diagnose_workbook(str(tmp_path / "Nope.xlsx"))
    assert r == {"file": str(tmp_path / "Nope.xlsx"), "readable": False, "error": "file not found"}
    assert time.monotonic() - t0 < 0.5
