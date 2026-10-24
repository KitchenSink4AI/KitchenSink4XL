"""Excel ending under a COM call is answered as Excel gone, never as a bad
parameter.

Excel killed (or crashed) after writing its copy and before closing the
workbook: the copy was discarded and the workbook left unchanged, correctly,
but the answer was "Excel refused while closing the workbook: AttributeError:
Open.Close", a plain refusal the envelope codes BAD_PARAMS, which tells an
assistant its arguments were wrong. pywin32 raises that AttributeError with
no COM error behind it when the member lookup meets a dead Excel (MEASURED),
so the exception alone cannot tell. The process can: when the Excel the call
acquired has ended, the answer is ExcelDisconnected (CONFLICT), whatever the
exception. The LIVE test starts the product's hidden Excel: private desktop
only.
"""
from __future__ import annotations

import hashlib
import os

import openpyxl
import pytest

from xlsx_mcp import envelope
from xlsx_mcp.com import instances
from xlsx_mcp.com import session as com_session
from xlsx_mcp.core.errors import ExcelDisconnected, XlMcpError
from xlsx_mcp.ops import comtier

LIVE = pytest.mark.skipif(not com_session.com_available()[0], reason="needs Excel (COM)")


@pytest.fixture
def job_pid():
    old = getattr(instances._THREAD, "pid", None)
    yield instances._THREAD
    instances._THREAD.pid = old


def test_a_failed_lookup_on_an_excel_that_ended_is_excel_gone(job_pid, monkeypatch):
    monkeypatch.setattr(instances, "pid_alive", lambda pid: False)
    job_pid.pid = 4242
    out = comtier._wrap_com_error(AttributeError("Open.Close"), "closing the workbook")
    assert isinstance(out, ExcelDisconnected)
    assert envelope.classify(out) == "CONFLICT"
    assert "did not complete" in str(out)


def test_a_failed_lookup_on_a_live_excel_stays_a_refusal(job_pid, monkeypatch):
    monkeypatch.setattr(instances, "pid_alive", lambda pid: True)
    job_pid.pid = 4242
    out = comtier._wrap_com_error(AttributeError("Open.NoSuchMember"), "closing the workbook")
    assert type(out) is XlMcpError


def test_outside_a_com_job_the_exception_alone_decides(job_pid, monkeypatch):
    monkeypatch.setattr(instances, "pid_alive", lambda pid: False)
    job_pid.pid = None
    assert type(comtier._wrap_com_error(AttributeError("Open.Close"), "closing")) is XlMcpError


def test_an_excel_that_is_going_gets_a_moment(monkeypatch):
    calls = []

    def alive(pid):
        calls.append(pid)
        return len(calls) < 3  # still listed twice, then gone
    monkeypatch.setattr(instances, "pid_alive", alive)
    assert com_session.excel_gone(4242, wait=1.0) is True
    assert com_session.excel_gone(4242, wait=0.0) is True


def _sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


@LIVE
def test_live_excel_dying_after_its_copy_is_answered_as_excel_gone(tmp_path, monkeypatch):
    pytest.importorskip("win32com.client")
    import psutil
    import win32process

    monkeypatch.setenv("KS4XL_SAVE_JOURNAL_DIR", str(tmp_path / "journal"))
    p = tmp_path / "Book.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "v0"
    wb.active["B1"] = "text long enough for autofit to change a width"
    wb.save(p)
    before = _sha(p)
    real = com_session.save_workbook
    killed = {}

    def save_then_excel_dies(wb_, **kw):
        out = real(wb_, **kw)  # Excel wrote its whole copy
        _t, pid = win32process.GetWindowThreadProcessId(int(wb_.Application.Hwnd))
        proc = psutil.Process(pid)
        assert proc.name().upper() == "EXCEL.EXE"
        assert pid in com_session.get_executor()._manager.reclaimable_pids()  # this server's own
        proc.kill()
        proc.wait(10)
        killed["pid"] = pid
        return out

    monkeypatch.setattr(com_session, "save_workbook", save_then_excel_dies)
    try:
        with pytest.raises(ExcelDisconnected):
            comtier.com_autofit(str(p), timeout_seconds=120)
    finally:
        monkeypatch.setattr(com_session, "save_workbook", real)
        com_session.get_executor().shutdown()
    assert killed
    assert _sha(p) == before, "the workbook changed"
    assert [n for n in os.listdir(tmp_path) if n.startswith((".ks4xl-write-", "~$"))] == []
    # and the next call works in a fresh Excel
    assert comtier.com_autofit(str(p), timeout_seconds=120)["ok"]
    com_session.get_executor().shutdown()
