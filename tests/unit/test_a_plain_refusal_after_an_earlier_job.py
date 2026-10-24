"""A plain Excel refusal stays a plain refusal after an earlier COM job.

The answer tells "Excel ended under the call" (CONFLICT, retry with a fresh
Excel) from Excel's own refusal (BAD_PARAMS, the call is wrong) by asking
whether the Excel of the job in flight is still running. That Excel was
remembered per thread when a job acquired it and never forgotten, so once an
earlier job's Excel had ended, every later refusal on the same thread read
as "the Excel worker process stopped answering", and a wrong call was
answered as one to retry. The thread now forgets its Excel when the job lets
it go (release, hand-over, quit) and when the job ends, however it ends.

No Excel here except the test marked LIVE, which starts the product's hidden
Excel (run it on a private desktop).
"""
from __future__ import annotations

import time

import pytest

from test_com_process_safety import _FakeComError
from xlsx_mcp import envelope
from xlsx_mcp.com import instances
from xlsx_mcp.com import session as com_session
from xlsx_mcp.core.errors import ExcelDisconnected
from xlsx_mcp.ops import comtier

LIVE = pytest.mark.skipif(not com_session.com_available()[0], reason="needs Excel (COM)")

#: No Excel has this PID: an Excel that has ended.
ENDED = 4_000_001


def _plain_refusal():
    return _FakeComError(-2147352567, "boom", (0, "Microsoft Excel", "Range of cells is invalid"))


def _assert_plain(wrapped):
    assert not isinstance(wrapped, ExcelDisconnected), str(wrapped)
    assert envelope.classify(wrapped) == "BAD_PARAMS"
    assert "Range of cells is invalid" in str(wrapped)


class _App:
    """An Excel that holds nothing of the user's."""
    Visible = False
    UserControl = False

    class Workbooks:
        Count = 0


@pytest.fixture()
def manager(tmp_path, monkeypatch):
    m = instances.ExcelInstanceManager(tmp_path / "journal.json")
    monkeypatch.setattr(m, "spawn", lambda: instances.Worker(app=_App(), pid=ENDED))
    monkeypatch.setattr(m, "users_claim", lambda w: None)
    yield m
    instances._THREAD.pid = None


def test_a_refusal_after_a_released_job_is_plain(manager):
    w = manager.acquire()
    assert instances.current_pid() == ENDED
    manager.release_to_pool(w)
    assert instances.current_pid() is None
    _assert_plain(comtier._wrap_com_error(_plain_refusal(), "autofitting"))


def test_a_refusal_after_a_handed_over_excel_is_plain(manager):
    w = manager.acquire()
    manager.hand_over(w, "it was made visible")
    assert instances.current_pid() is None
    _assert_plain(comtier._wrap_com_error(_plain_refusal(), "autofitting"))


def test_a_refusal_after_a_quit_excel_is_plain(manager):
    w = manager.acquire()
    manager.quit(w)
    assert instances.current_pid() is None
    _assert_plain(comtier._wrap_com_error(_plain_refusal(), "autofitting"))


def test_the_excel_in_flight_still_decides(manager):
    """While a job holds an Excel that has ended, the answer is still that
    Excel ended (the check this must not weaken)."""
    manager.acquire()
    assert isinstance(comtier._wrap_com_error(_plain_refusal(), "autofitting"), ExcelDisconnected)


@pytest.mark.skipif(not com_session.com_available()[0],
                    reason="COM plumbing absent (run_com needs Excel's COM registration)")
def test_a_job_that_ends_in_an_exception_forgets_its_excel():
    """A job that acquired an Excel and failed before letting it go: the next
    job on the COM thread does not inherit it."""
    def failing(_manager):
        instances._THREAD.pid = ENDED   # what acquire() does
        raise RuntimeError("the job failed before it released its Excel")

    def ask(_manager):
        return instances.current_pid(), envelope.classify(
            comtier._wrap_com_error(_plain_refusal(), "autofitting"))

    try:
        with pytest.raises(RuntimeError):
            com_session.run_com("fails", failing, timeout=60)
        assert com_session.run_com("asks", ask, timeout=60) == (None, "BAD_PARAMS")
    finally:
        com_session.get_executor().shutdown()


@LIVE
@pytest.mark.timeout(300)
def test_live_an_earlier_jobs_excel_ended_and_a_refusal_is_plain():
    """A real job acquires and releases the product's hidden Excel; that
    Excel ends; a later job's plain refusal is still plain."""
    def first(manager):
        w = manager.acquire()
        pid = w.pid
        manager.release_to_pool(w)
        return pid

    def later(_manager):
        return envelope.classify(comtier._wrap_com_error(_plain_refusal(), "autofitting"))

    try:
        pid = com_session.run_com("first", first, timeout=120)
        com_session.run_com("end it", lambda m: m.force_reclaim(), timeout=120)
        deadline = time.monotonic() + 30
        while instances.pid_alive(pid) and time.monotonic() < deadline:
            time.sleep(0.2)
        assert not instances.pid_alive(pid)
        assert com_session.run_com("later", later, timeout=60) == "BAD_PARAMS"
    finally:
        com_session.get_executor().shutdown()
