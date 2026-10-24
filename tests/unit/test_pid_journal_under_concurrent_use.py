"""The journal of the Excel processes a server starts, under concurrent use.

Every kitchensink4xl server of one user shares one journal of the Excels it
started (ks4xl_pid_journal.json in the temp folder). A server writes it when
it starts an Excel (a record, then a note on it) and when one ends, and
every server reads it when it starts. A write replaced the file while
another process had it open for reading; Windows refused the replace, and
the call that was starting Excel answered WORKBOOK_LOCKED about the journal,
telling the user to close a file in Excel. Two servers updating it at the
same moment could also lose each other's records, and an Excel that no
record names is never reclaimed.

Now a write waits out a reader for a moment, every update holds a short lock
that all servers share, and a journal that still cannot be written is
logged, never reported as the user's workbook being locked. No Excel here.
"""
from __future__ import annotations

import json
import logging
import subprocess
import sys
import textwrap
import threading
import time
from pathlib import Path

import pytest

from xlsx_mcp.com import instances

SRC = Path(instances.__file__).resolve().parents[2]
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _hammer(path: Path, stop: threading.Event, hold: float = 0.004,
            gap: float = 0.002) -> threading.Thread:
    """A reader that keeps the journal open most of the time, as another
    server's start-up sweep or a status read does: open for `hold`, closed
    for `gap`, again and again."""
    def run():
        while not stop.is_set():
            try:
                with open(path, "rb") as f:
                    f.read()
                    time.sleep(hold)
            except OSError:
                pass
            time.sleep(gap)
    t = threading.Thread(target=run, daemon=True)
    t.start()
    return t


def test_a_reader_holding_the_journal_never_fails_a_write(tmp_path):
    path = tmp_path / "ks4xl_pid_journal.json"
    journal = instances.PidJournal(path)
    journal.record(1)
    stop = threading.Event()
    _hammer(path, stop)
    try:
        for pid in range(100, 140):
            journal.record(pid)
            journal.annotate(pid, remote_prior=False)
    finally:
        stop.set()
    records = journal.records()
    assert set(map(int, records)) == {1} | set(range(100, 140))
    assert all(records[str(p)].get("remote_prior") is False for p in range(100, 140))


def test_starting_excel_while_the_journal_is_read_is_never_a_locked_workbook(tmp_path, monkeypatch):
    path = tmp_path / "ks4xl_pid_journal.json"
    manager = instances.ExcelInstanceManager(path)
    pids = iter(range(5000, 5100))
    monkeypatch.setattr(instances, "list_excel_pids", lambda: set())
    monkeypatch.setattr(manager, "_new_app", lambda: object())
    monkeypatch.setattr(manager, "_ignore_dde", lambda app: False)
    monkeypatch.setattr(manager, "_resolve_pid", lambda app, before: next(pids))
    instances.PidJournal(path).record(1)
    stop = threading.Event()
    _hammer(path, stop)
    try:
        started = [manager.spawn().pid for _ in range(30)]
    finally:
        stop.set()
    assert set(started) <= set(map(int, manager.journal.records()))


_RECORDER = textwrap.dedent(
    r"""
    import sys, time
    sys.path.insert(0, sys.argv[1])
    from pathlib import Path
    from xlsx_mcp.com.instances import PidJournal
    journal = PidJournal(sys.argv[2])
    first, count, go = int(sys.argv[3]), int(sys.argv[4]), Path(sys.argv[5])
    while not go.exists():
        time.sleep(0.001)
    for pid in range(first, first + count):
        journal.record(pid)
        journal.annotate(pid, remote_prior=False)
    """
)


@pytest.mark.timeout(240)
def test_servers_recording_at_once_lose_no_record(tmp_path):
    path = tmp_path / "ks4xl_pid_journal.json"
    script = tmp_path / "recorder.py"
    script.write_text(_RECORDER, encoding="utf-8")
    go = tmp_path / "go"
    procs = [subprocess.Popen([sys.executable, str(script), str(SRC), str(path), str(10_000 * (i + 1)),
                               "40", str(go)], stdin=subprocess.DEVNULL, creationflags=NO_WINDOW)
             for i in range(4)]
    time.sleep(1.0)
    go.write_text("go", encoding="utf-8")
    assert [p.wait(180) for p in procs] == [0, 0, 0, 0]
    records = json.loads(path.read_text("utf-8"))
    want = {10_000 * (i + 1) + k for i in range(4) for k in range(40)}
    assert set(map(int, records)) == want, sorted(want - set(map(int, records)))[:10]
    assert all(r.get("remote_prior") is False for r in records.values())


def test_a_journal_that_cannot_be_written_is_logged_not_raised(tmp_path, monkeypatch, caplog):
    path = tmp_path / "ks4xl_pid_journal.json"
    journal = instances.PidJournal(path)

    def refused(src, dst):
        raise PermissionError(13, "The process cannot access the file", str(dst))
    monkeypatch.setattr(instances.os, "replace", refused)
    monkeypatch.setattr(instances, "JOURNAL_REPLACE_WAIT_S", 0.1)
    with caplog.at_level(logging.WARNING):
        journal.record(7)  # no exception reaches the call that started Excel
    assert "journal" in caplog.text and "could not be written" in caplog.text
    assert journal.records() == {}
    assert not list(tmp_path.glob("*.tmp")), "the write's temporary file was left behind"
