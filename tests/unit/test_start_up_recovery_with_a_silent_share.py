"""Start-up recovery never waits out a share that does not answer.

One journal entry for a workbook on a share that does not answer held every
server start for the network's timeout (about 21 s an entry) before the
server answered its client, and the entry was never dropped. Each folder now
gets one second, a silent share is skipped for the rest of the run and for a
while after, and an entry whose folder stays unreachable for weeks is
dropped (the workbook is still put back from the files beside it at its next
write lock). No Excel here.
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



# ------------------------------------------------------- a share that does not answer


def test_a_silent_share_costs_a_start_one_second(tmp_path, journal, monkeypatch):
    unc = "\\\\silent-share-host\\share"
    real_isdir = os.path.isdir

    def isdir(p):
        if str(p).startswith(unc):
            time.sleep(20)  # the network's timeout
            return False
        return real_isdir(p)
    monkeypatch.setattr(safesave.os.path, "isdir", isdir)
    for i in range(3):
        _entry(journal, workbook=f"{unc}\\Book{i}.xlsx", aside="x", new="y")
    book, aside, _new, _e = _stopped_save(tmp_path, finished=False, backup=False)
    t0 = time.monotonic()
    restored = safesave.recover_interrupted_saves()
    took = time.monotonic() - t0
    assert took < 3.0, f"{took:.1f} s"
    assert restored == [str(book)], "a local workbook after the silent share is still put back"
    assert len(list(journal.glob("*.json"))) == 3, "kept: the share may come back"


def test_an_entry_for_a_folder_gone_for_weeks_is_dropped(tmp_path, journal):
    _entry(journal, workbook=str(tmp_path / "no-such-folder" / "Book.xlsx"), aside="x", new="y",
           time=time.time() - 40 * 86400)
    _entry(journal, workbook=str(tmp_path / "no-such-folder" / "Two.xlsx"), aside="x", new="y")
    safesave.recover_interrupted_saves()
    left = [json.loads(p.read_text("utf-8"))["workbook"] for p in journal.glob("*.json")]
    assert left == [str(tmp_path / "no-such-folder" / "Two.xlsx")]
