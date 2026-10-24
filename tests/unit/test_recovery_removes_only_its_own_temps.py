"""Recovery removes only the temporary files this server's saves write.

Recovery of a save that did not finish removed whatever path a journal
entry's "new" named: an entry naming a user's file deleted that file, and one
naming the workbook deleted the workbook. The product only ever writes its
own temporary file there, so this took a damaged or foreign entry; it now
takes nothing: only a write temp in the workbook's folder, or a restore temp
in its backup folder, is ever removed, and never the workbook under any
name. No Excel here.
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


# ------------------------------------------------- only its own temporary files


def _entry(journal, **record):
    journal.mkdir(exist_ok=True)
    base = {"staged": [], "pid": 2147483640, "token": "a-server-that-is-gone",
            "pid_created": 1.0, "host": safesave._local_host(), "time": time.time() - 7200}
    base.update(record)
    path = journal / f"{uuid.uuid4().hex}.json"
    path.write_text(json.dumps(base), "utf-8")
    return path


@pytest.mark.parametrize("named", ["another-file", "the-workbook", "a-link-of-the-workbook"])
def test_recovery_never_removes_a_file_that_is_not_its_own_temp(tmp_path, journal, named):
    book = tmp_path / "Book.xlsx"
    _book(book, "v1")
    if named == "another-file":
        victim = tmp_path / "Precious.xlsx"
        _book(victim, "the user's")
    elif named == "the-workbook":
        victim = book
    else:
        victim = tmp_path / f"{safesave.WRITE_TMP_PREFIX}looks-like-a-temp.xlsx"
        os.link(book, victim)
    _entry(journal, workbook=str(book), aside=str(tmp_path / "nothing.tmp"), new=str(victim))
    safesave.recover_interrupted_saves()
    assert victim.exists() and book.exists()
    assert _a1(book) == "v1"


def test_recovery_still_removes_its_own_unanswered_edit(tmp_path, journal):
    book = tmp_path / "Book.xlsx"
    _book(book, "v1")
    new = tmp_path / f"{safesave.WRITE_TMP_PREFIX}{uuid.uuid4().hex}.xlsx"
    _book(new, "v2")
    _entry(journal, workbook=str(book), aside=str(tmp_path / "nothing.tmp"), new=str(new))
    safesave.recover_interrupted_saves()
    assert not new.exists() and _a1(book) == "v1"
    assert list(journal.glob("*.json")) == []


def test_a_put_back_never_removes_a_foreign_file_named_as_new(tmp_path, journal):
    book = tmp_path / "Book.xlsx"
    _book(book, "v1")
    aside = tmp_path / swapwait.aside_name("Book.xlsx", "0123456789ab")
    os.rename(book, aside)
    victim = tmp_path / "Precious.xlsx"
    _book(victim, "the user's")
    _entry(journal, workbook=str(book), aside=str(aside), new=str(victim))
    with safesave.write_lock(book):
        pass
    assert victim.exists()


