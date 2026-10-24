"""A stopped save's staged backup rotation lands only on the slot it was staged for.

A save stopped after its swap leaves its backup rotation staged. When a save
the journal did not see (an older build, another journal folder, another
Windows user on a shared folder) rotated the slot meanwhile, the next start
landed the stale staged copy over it, and the version that save had put
there existed nowhere. A staged move now records which file its slot was and
lands only while the slot is still that file (an entry from the previous
version, which did not record it, lands only over an older slot). No Excel
here.
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


def _entry(journal, **record):
    journal.mkdir(exist_ok=True)
    base = {"staged": [], "pid": 2147483640, "token": "a-server-that-is-gone",
            "pid_created": 1.0, "host": safesave._local_host(), "time": time.time() - 7200}
    base.update(record)
    path = journal / f"{uuid.uuid4().hex}.json"
    path.write_text(json.dumps(base), "utf-8")
    return path



# ------------------------------------------------------------- staged rotation


def test_a_stop_after_the_swap_still_ends_with_the_right_undo(tmp_path):
    book, aside, _new, _entry_ = _stopped_save(tmp_path, finished=True)
    safesave.recover_interrupted_saves()
    assert _a1(book) == "v2"
    assert _a1(safesave.slot_dir(book) / safesave.PREV_SLOT) == "v1", "the undo of v2 is v1"
    assert not aside.exists(), "prev holds that very file: nothing to keep"


def test_a_stale_staged_rotation_never_lands_over_a_newer_slot(tmp_path, journal, monkeypatch):
    book, aside, _new, _entry_ = _stopped_save(tmp_path, finished=True)
    # A save the journal does not see (another journal folder; an older build).
    monkeypatch.setenv("KS4XL_SAVE_JOURNAL_DIR", str(tmp_path / "another_journal"))
    cells.set_cell(str(book), "A1", "v3")
    monkeypatch.setenv("KS4XL_SAVE_JOURNAL_DIR", str(journal))
    assert _a1(safesave.slot_dir(book) / safesave.PREV_SLOT) == "v2"
    safesave.recover_interrupted_saves()
    assert _a1(safesave.slot_dir(book) / safesave.PREV_SLOT) == "v2", "v2 would exist nowhere"
    assert _a1(book) == "v3"
    assert list(safesave.slot_dir(book).glob(".slot-*.tmp")) == [], "the stale copy stayed staged"


@pytest.mark.parametrize("slot_newer", [True, False])
def test_an_entry_without_slot_identity_lands_only_over_an_older_slot(tmp_path, journal, slot_newer):
    """Entries written by the previous version record (temp, slot) pairs."""
    book = tmp_path / "Book.xlsx"
    _book(book, "v1")
    cells.set_cell(str(book), "B1", "x")
    slots = safesave.slot_dir(book)
    staged = slots / ".slot-0123456789abcdef0123456789abcdef.tmp"
    _book(staged, "staged")
    time.sleep(0.05)
    if slot_newer:
        _book(slots / safesave.PREV_SLOT, "newer")
    else:
        os.utime(slots / safesave.PREV_SLOT, (time.time() - 3600, time.time() - 3600))
    _entry(journal, workbook=str(book), aside=str(tmp_path / "nothing.tmp"),
           new=str(tmp_path / f"{safesave.WRITE_TMP_PREFIX}gone.xlsx"),
           staged=[[str(staged), str(slots / safesave.PREV_SLOT)]])
    safesave.recover_interrupted_saves()
    assert _a1(slots / safesave.PREV_SLOT) == ("newer" if slot_newer else "staged")
    assert not staged.exists()


