"""One server's exit or COM timeout never kills another server's Excel.

The Excel instance journal (%TEMP%\\ks4xl_pid_journal.json) is one file per
user, shared by every kitchensink4xl server. force_reclaim (every shutdown,
every COM timeout) reclaimed EVERY record in it, so one server's exit killed
the pooled Excel of every other live server in the middle of its call
(MEASURED). A server now reclaims only its own records
and those whose owning server is gone, checked by creation time, so a PID
Windows recycled does not keep a dead owner's Excel alive. No Excel starts
here: the process calls are stubbed.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time

import pytest

from xlsx_mcp.com import instances
from xlsx_mcp.com import session as com_session


@pytest.fixture
def other_server():
    proc = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"],
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    yield proc
    proc.kill()
    proc.wait()


def _journal(tmp_path, records):
    path = tmp_path / "ks4xl_pid_journal.json"
    path.write_text(json.dumps(records), "utf-8")
    return path


def _stub(monkeypatch, alive_excel):
    killed = []
    monkeypatch.setattr(instances, "pid_alive", lambda pid: pid in alive_excel)
    monkeypatch.setattr(instances, "identity_verdict", lambda pid, rec: "match")
    monkeypatch.setattr(instances, "taskkill", lambda pid: killed.append(pid) or True)
    return killed


def _rec(owner_pid, owner_started):
    return {"spawned_at": time.time(), "status": "owned", "create_time": 1.0,
            "owner_pid": owner_pid, "owner_started": owner_started}


def test_another_live_servers_excel_is_never_reclaimed(tmp_path, monkeypatch, other_server):
    started = instances.process_create_time(other_server.pid)
    path = _journal(tmp_path, {"424242": _rec(other_server.pid, started)})
    killed = _stub(monkeypatch, {424242})
    instances.ExcelInstanceManager(path).force_reclaim()
    assert killed == []
    assert "424242" in json.loads(path.read_text("utf-8"))


def test_this_servers_and_abandoned_excels_are_reclaimed(tmp_path, monkeypatch):
    dead_owner = subprocess.Popen([sys.executable, "-c", "pass"],
                                  creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    dead_owner.wait()
    path = _journal(tmp_path, {
        "111": _rec(os.getpid(), instances._own_start_time()),  # mine
        "222": _rec(dead_owner.pid, time.time() - 60),           # owner gone
    })
    killed = _stub(monkeypatch, {111, 222})
    instances.ExcelInstanceManager(path).force_reclaim()
    assert sorted(killed) == [111, 222]


def test_an_owner_whose_pid_was_recycled_counts_as_gone(tmp_path, monkeypatch, other_server):
    # The record names other_server's PID but a creation time an hour earlier:
    # the server that wrote it is gone and Windows gave its PID to another process.
    path = _journal(tmp_path, {"333": _rec(other_server.pid, time.time() - 3600)})
    killed = _stub(monkeypatch, {333})
    manager = instances.ExcelInstanceManager(path)
    assert manager.journal.abandoned_pids() == {333}
    manager.force_reclaim()
    assert killed == [333]


def test_another_servers_excel_is_foreign_to_the_timeout_reap(tmp_path, monkeypatch, other_server):
    started = instances.process_create_time(other_server.pid)
    path = _journal(tmp_path, {"424242": _rec(other_server.pid, started)})
    manager = instances.ExcelInstanceManager(path)
    assert 424242 in manager.owned_pids()  # still not the user's own Excel
    assert 424242 not in manager.reclaimable_pids()


def test_a_lookup_on_an_excel_that_is_gone_is_not_a_bad_parameter():
    pywintypes = pytest.importorskip("pywintypes")
    try:
        try:
            raise pywintypes.com_error(-2147417848, "The object invoked has disconnected",
                                       None, None)
        except pywintypes.com_error:
            raise AttributeError("Open.SaveCopyAs")
    except AttributeError as exc:
        gone = exc
    assert com_session.excel_process_died(gone)
    assert not com_session.excel_process_died(AttributeError("Open.NoSuchMember"))
    from xlsx_mcp.ops import comtier
    from xlsx_mcp.core.errors import ExcelDisconnected
    assert isinstance(comtier._wrap_com_error(gone, "saving the workbook"), ExcelDisconnected)
