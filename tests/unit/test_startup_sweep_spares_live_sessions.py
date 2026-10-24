"""Another server's COM start-up sweep never reclaims a live session's Excel.

The sweep reclaims abandoned records (their server is gone) and, for records
written before owner tracking, those older than STALE_JOURNAL_SECONDS. The
age clause also caught records WITH a live owner: a second session that kept
its pooled Excel for 15 minutes lost it, mid-call, to a server starting
(MEASURED with real Excels: the 20-minute-old record of a live server, the
next server's first COM call, that server's Excel killed). No Excel here:
stand-in processes.
"""
from __future__ import annotations

import subprocess
import sys
import time

import pytest

from xlsx_mcp.com import instances
from xlsx_mcp.com import session as com_session


@pytest.fixture
def live_owner():
    proc = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"],
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    yield proc
    proc.kill()
    proc.wait()


def _sweep(tmp_path, monkeypatch, records):
    killed = []
    monkeypatch.setattr(instances, "pid_alive", lambda pid: True)
    monkeypatch.setattr(instances, "identity_verdict", lambda pid, rec: "match")
    monkeypatch.setattr(instances, "excel_window_shown", lambda pid: False, raising=False)
    monkeypatch.setattr(instances, "taskkill", lambda pid: killed.append(pid) or True)
    mgr = instances.ExcelInstanceManager(tmp_path / "journal.json")
    mgr.journal._save(records)
    com_session.ComExecutor._sweep_stale_journal(mgr)
    return killed, mgr.journal.records()


def test_an_aged_record_of_a_live_session_is_left_alone(tmp_path, monkeypatch, live_owner):
    started = instances.process_create_time(live_owner.pid)
    killed, left = _sweep(tmp_path, monkeypatch, {"5001": {
        "spawned_at": time.time() - 1200, "status": "owned", "create_time": 1.0,
        "owner_pid": live_owner.pid, "owner_started": started}})
    assert killed == [] and "5001" in left


def test_an_aged_record_without_an_owner_is_still_reclaimed(tmp_path, monkeypatch):
    killed, left = _sweep(tmp_path, monkeypatch, {"5002": {
        "spawned_at": time.time() - 1200, "status": "owned", "create_time": 1.0}})
    assert killed == [5002] and left == {}


def test_a_record_whose_owner_is_gone_is_reclaimed_at_any_age(tmp_path, monkeypatch):
    dead = subprocess.Popen([sys.executable, "-c", "pass"],
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    dead.wait()
    killed, _left = _sweep(tmp_path, monkeypatch, {"5003": {
        "spawned_at": time.time() - 5, "status": "owned", "create_time": 1.0,
        "owner_pid": dead.pid, "owner_started": time.time() - 60}})
    assert killed == [5003]


def test_a_fresh_record_without_an_owner_is_left_alone(tmp_path, monkeypatch):
    killed, left = _sweep(tmp_path, monkeypatch, {"5004": {
        "spawned_at": time.time() - 5, "status": "owned", "create_time": 1.0}})
    assert killed == [] and "5004" in left
