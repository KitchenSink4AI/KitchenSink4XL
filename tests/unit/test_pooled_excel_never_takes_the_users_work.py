"""The pooled hidden Excel never takes, and never discards, the user's work.

The pool left IgnoreRemoteRequests False, so a DDE open (a file association
using DDE, a program sending [OPEN(...)] to Excel|System) landed in the
product's hidden Excel, and Excel made that instance visible and
user-controlled (MEASURED): the user worked in the pool without knowing.
The server's exit (Quit with alerts off, then a reclaim by PID), a COM
timeout and another server's start-up sweep then discarded their unsaved
work without a prompt.

Now every pooled instance ignores DDE (the user's own Excel option is put
back before any Quit; a clean Quit writes Excel's options, a killed Excel
writes nothing, MEASURED). An instance holding a workbook the server did not
open, or one that shows its window, is never closed, quit or killed: it is
made visible with its alerts back, dropped from the pool and the journal,
and logged.

The tests marked LIVE start the product's hidden Excel: run them on a
private desktop only.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time

import openpyxl
import pytest

from xlsx_mcp.com import instances
from xlsx_mcp.com import session as com_session

LIVE = pytest.mark.skipif(not com_session.com_available()[0], reason="needs Excel (COM)")


# ------------------------------------------------------------ stand-ins


class _Book:
    def __init__(self, full):
        self.FullName = full


class _Books:
    def __init__(self, names):
        self._books = [_Book(n) for n in names]

    @property
    def Count(self):
        return len(self._books)

    def Item(self, i):
        return self._books[i - 1]


class _App:
    def __init__(self, books=(), visible=False, user_control=False, remote=False):
        self.Workbooks = _Books(books)
        self.Visible = visible
        self.UserControl = user_control
        self.IgnoreRemoteRequests = remote
        self.DisplayAlerts = False
        self.EnableEvents = False
        self.ScreenUpdating = False
        self.quits = []

    def Quit(self):
        self.quits.append({"IgnoreRemoteRequests": self.IgnoreRemoteRequests})


def _manager(tmp_path, app, pid=4321, prior=False, opened=()):
    mgr = instances.ExcelInstanceManager(tmp_path / "journal.json")
    mgr.journal.record(pid)
    w = instances.Worker(app=app, pid=pid, remote_prior=prior)
    mgr._workers.append(w)
    mgr._pool = w
    instances._OPENED[id(app)] = set()
    for name in opened:
        instances.note_opened(app, name)
    return mgr, w


def test_a_new_pooled_excel_ignores_dde_and_keeps_the_users_option(tmp_path, monkeypatch):
    app = _App(remote=False)
    mgr = instances.ExcelInstanceManager(tmp_path / "journal.json")
    monkeypatch.setattr(mgr, "_new_app", lambda: app)
    monkeypatch.setattr(mgr, "_resolve_pid", lambda a, before: os.getpid())
    w = mgr.spawn()
    assert app.IgnoreRemoteRequests is True
    assert w.remote_prior is False
    assert mgr.journal.records()[str(os.getpid())]["remote_prior"] is False


def test_quit_puts_the_users_dde_option_back_first(tmp_path):
    app = _App(books=[r"C:\work\Mine.xlsx"], remote=True)
    mgr, w = _manager(tmp_path, app, prior=False, opened=[r"C:\work\Mine.xlsx"])
    mgr.quit(w)
    assert app.quits == [{"IgnoreRemoteRequests": False}]
    assert mgr._workers == [] and mgr._pool is None


@pytest.mark.parametrize("books,visible,user_control", [
    ([r"C:\Users\someone\UserBook.xlsx"], False, False),
    ([r"C:\work\Mine.xlsx", "Book1"], False, False),
    ([], True, False),
    ([], False, True),
], ids=["their-workbook", "an-unsaved-workbook", "made-visible", "user-control"])
def test_exit_hands_the_instance_over_instead_of_quitting(tmp_path, books, visible, user_control):
    app = _App(books=books, visible=visible, user_control=user_control, remote=True)
    mgr, w = _manager(tmp_path, app, prior=False, opened=[r"C:\work\Mine.xlsx"])
    mgr.quit(w)
    assert app.quits == [], "quit an Excel that holds the user's work"
    assert (app.DisplayAlerts, app.EnableEvents, app.ScreenUpdating, app.Visible,
            app.UserControl, app.IgnoreRemoteRequests) == (True, True, True, True, True, False)
    assert mgr._workers == [] and mgr._pool is None
    assert str(4321) not in mgr.journal.records(), "a later sweep could still kill it"
    assert instances.HANDED_OVER[-1]["pid"] == 4321


def test_the_pool_is_never_reused_once_the_user_has_it(tmp_path, monkeypatch):
    theirs = _App(books=[r"C:\Users\someone\UserBook.xlsx"], remote=True)
    mgr, old = _manager(tmp_path, theirs, pid=os.getpid())
    monkeypatch.setattr(instances, "pid_alive", lambda pid: True)
    fresh = instances.Worker(app=_App(), pid=99)
    monkeypatch.setattr(mgr, "spawn", lambda: mgr._workers.append(fresh) or fresh)
    got = mgr.acquire()
    assert got is fresh and mgr._pool is fresh
    assert theirs.quits == [] and theirs.Visible is True and theirs.DisplayAlerts is True
    assert old not in mgr._workers


def test_the_pool_is_reused_while_it_holds_only_this_servers_workbooks(tmp_path, monkeypatch):
    app = _App(books=[r"C:\work\Mine.xlsx"])
    mgr, w = _manager(tmp_path, app, opened=[r"c:\WORK\mine.xlsx"])
    monkeypatch.setattr(instances, "pid_alive", lambda pid: True)
    assert mgr.acquire() is w


def test_no_kill_path_ends_an_excel_that_shows_its_window(tmp_path, monkeypatch):
    """A COM timeout, the exit's reclaim and every sweep reclaim through
    reclaim_journaled_pid: an Excel showing its window is left running and
    dropped from the journal."""
    killed = []
    monkeypatch.setattr(instances, "pid_alive", lambda pid: True)
    monkeypatch.setattr(instances, "identity_verdict", lambda pid, rec: "match")
    monkeypatch.setattr(instances, "taskkill", lambda pid: killed.append(pid) or True)
    monkeypatch.setattr(instances, "excel_window_shown", lambda pid: pid == 4321)
    monkeypatch.setattr(instances, "_return_to_user", lambda pid, prior=None, timeout=5.0: True)
    mgr, w = _manager(tmp_path, _App(), pid=4321)
    mgr.journal.record(5555)
    res = mgr.force_reclaim()
    assert res.handed_to_user == [4321] and killed == [5555]
    assert mgr.journal.records() == {}
    assert mgr._workers == [] and mgr._pool is None


def test_another_servers_start_up_sweep_spares_an_excel_that_shows_its_window(tmp_path, monkeypatch):
    killed = []
    monkeypatch.setattr(instances, "pid_alive", lambda pid: True)
    monkeypatch.setattr(instances, "identity_verdict", lambda pid, rec: "match")
    monkeypatch.setattr(instances, "taskkill", lambda pid: killed.append(pid) or True)
    monkeypatch.setattr(instances, "excel_window_shown", lambda pid: True)
    monkeypatch.setattr(instances, "_return_to_user", lambda pid, prior=None, timeout=5.0: True)
    mgr = instances.ExcelInstanceManager(tmp_path / "journal.json")
    mgr.journal._save({"777": {"spawned_at": time.time() - 3600, "status": "owned",
                               "create_time": 1.0, "owner_pid": 2147483640,
                               "owner_started": 1.0}})
    outcomes = com_session.ComExecutor._sweep_stale_journal(mgr)
    assert outcomes == {"handed_to_user": [777]} and killed == []
    assert mgr.journal.records() == {}


def test_the_window_check_reads_this_desktops_windows_without_waiting():
    assert instances.excel_window_shown(os.getpid()) is False  # no Excel window
    assert instances.excel_window_shown(0) is False


# ------------------------------------------------------------------ LIVE


def _book(p, v):
    wb = openpyxl.Workbook()
    wb.active["A1"] = v
    wb.active["B1"] = "text long enough for autofit to change a width"
    wb.save(p)


@pytest.fixture
def pool(tmp_path, monkeypatch):
    pytest.importorskip("win32com.client")
    monkeypatch.setenv("KS4XL_SAVE_JOURNAL_DIR", str(tmp_path / "saves"))
    from xlsx_mcp.ops import comtier
    tool = tmp_path / "Tool.xlsx"
    _book(tool, "tool")
    assert comtier.com_autofit(str(tool), timeout_seconds=120)["ok"]
    ex = com_session.get_executor()
    left = []
    yield tmp_path, ex, left
    ex.shutdown()
    import psutil
    for pid, created in left:  # the test's own Excel, exact PID and creation time
        try:
            p = psutil.Process(pid)
            if abs(p.create_time() - created) < 0.01 and p.name().upper() == "EXCEL.EXE":
                p.kill()
        except Exception:  # noqa: BLE001
            pass


def _job(fn):
    return com_session.run_com("users-work-probe", lambda mgr: fn(mgr), timeout=60)


def _land(user, typed, show=True):
    """What a DDE open and the user's typing leave in the pooled instance
    (MEASURED: Excel shows the instance, user-controlled)."""
    def body(mgr):
        app = mgr._pool.app
        wb = app.Workbooks.Open(str(user))
        wb.Worksheets(1).Range("A1").Value = typed
        if show:
            app.Visible = True
            app.UserControl = True
        return mgr._pool.pid
    return _job(body)


def _theirs(user, pid):
    """The user's workbook in the Excel ``pid``, reached through its window
    (as an accessibility client would; nothing is started), and that Excel."""
    import pythoncom
    import win32com.client
    import win32gui
    import win32process
    pythoncom.CoInitialize()
    tops, kids = [], []

    def on_top(h, _):
        if win32process.GetWindowThreadProcessId(h)[1] == pid and win32gui.GetClassName(h) == "XLMAIN":
            tops.append(h)
        return True

    def on_child(h, _):
        if win32gui.GetClassName(h) == "EXCEL7":
            kids.append(h)
        return True
    win32gui.EnumWindows(on_top, None)
    for top in tops:
        win32gui.EnumChildWindows(top, on_child, None)
    for h in kids:
        _r, lres = win32gui.SendMessageTimeout(h, 0x003D, 0, 0xFFFFFFF0, 0x0002, 2000)
        if lres:
            app = win32com.client.Dispatch(
                pythoncom.ObjectFromLresult(lres, pythoncom.IID_IDispatch, 0)).Application
            for i in range(1, app.Workbooks.Count + 1):
                wb = app.Workbooks(i)
                if os.path.normcase(str(wb.FullName)) == os.path.normcase(str(user)):
                    return wb, app
    raise AssertionError(f"no window of Excel {pid} holds {user}")


@LIVE
def test_live_the_pooled_excel_ignores_dde(pool):
    assert _job(lambda mgr: bool(mgr._pool.app.IgnoreRemoteRequests)) is True


@LIVE
def test_live_exit_leaves_the_users_workbook_open_with_their_typing(pool):
    import psutil
    tmp, ex, left = pool
    user = tmp / "UserBook.xlsx"
    _book(user, "saved")
    pid = _land(user, "unsaved typing")
    left.append((pid, psutil.Process(pid).create_time()))
    ex.shutdown()
    time.sleep(2)
    assert psutil.Process(pid).is_running(), "the server's exit ended the Excel holding the user's work"
    wb, app = _theirs(user, pid)
    assert wb.Worksheets(1).Range("A1").Value == "unsaved typing"
    assert bool(app.DisplayAlerts) and bool(app.Visible) and bool(app.UserControl)
    assert not bool(app.IgnoreRemoteRequests)
    journal = instances.ExcelInstanceManager().journal.records()
    assert str(pid) not in journal, "a later sweep could still kill it"


@LIVE
def test_live_the_next_call_runs_in_a_fresh_excel_and_leaves_theirs(pool):
    import psutil
    from xlsx_mcp.ops import comtier
    tmp, ex, left = pool
    user = tmp / "UserBook.xlsx"
    _book(user, "saved")
    pid = _land(user, "unsaved typing", show=False)  # not even shown: their workbook is enough
    left.append((pid, psutil.Process(pid).create_time()))
    other = tmp / "Other.xlsx"
    _book(other, "other")
    assert comtier.com_autofit(str(other), timeout_seconds=120)["ok"]
    assert _job(lambda mgr: mgr._pool.pid) != pid
    assert psutil.Process(pid).is_running()
    wb, app = _theirs(user, pid)
    assert wb.Worksheets(1).Range("A1").Value == "unsaved typing"
    assert bool(app.DisplayAlerts) and bool(app.Visible)


@LIVE
def test_live_a_com_timeouts_reclaim_spares_an_excel_that_shows_its_window(pool):
    """The reclaim a COM timeout runs (force_reclaim, no COM on the stuck
    instance): the window decides, and the alerts are put back through the
    workbook's window."""
    import psutil
    tmp, ex, left = pool
    user = tmp / "UserBook.xlsx"
    _book(user, "saved")
    pid = _land(user, "unsaved typing")
    left.append((pid, psutil.Process(pid).create_time()))
    manager = ex._manager
    res = manager.force_reclaim()
    assert pid in res.handed_to_user and pid not in res.killed
    time.sleep(1)
    assert psutil.Process(pid).is_running()
    wb, app = _theirs(user, pid)
    assert wb.Worksheets(1).Range("A1").Value == "unsaved typing"
    assert bool(app.DisplayAlerts), "closing their changed workbook would not ask to save"
    assert str(pid) not in manager.journal.records()


@LIVE
def test_live_another_servers_start_up_sweep_spares_it(pool, tmp_path):
    """A server that dies leaves its record; the next server's start-up sweep
    reclaims abandoned records, and spares one whose Excel shows its window."""
    import psutil
    tmp, ex, left = pool
    user = tmp / "UserBook.xlsx"
    _book(user, "saved")
    pid = _land(user, "unsaved typing")
    created = psutil.Process(pid).create_time()
    left.append((pid, created))
    dead = subprocess.Popen([sys.executable, "-c", "pass"],
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    dead.wait()
    journal = tmp_path / "another_servers_journal.json"
    journal.write_text(json.dumps({str(pid): {
        "spawned_at": time.time() - 3600, "status": "owned", "create_time": created,
        "owner_pid": dead.pid, "owner_started": time.time() - 3600}}), "utf-8")
    outcomes = com_session.ComExecutor._sweep_stale_journal(instances.ExcelInstanceManager(journal))
    assert outcomes == {"handed_to_user": [pid]}
    assert psutil.Process(pid).is_running()
    wb, _app = _theirs(user, pid)
    assert wb.Worksheets(1).Range("A1").Value == "unsaved typing"
