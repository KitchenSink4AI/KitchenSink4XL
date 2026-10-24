"""The network refusal tells the AI to ask the user first, and a success says it was allowed.

When a workbook would contact the internet, the refusal names how to go ahead
(allow_network=true). A weaker model that reads only that sentence passes the
flag on its own, Excel makes the request, and the user is told "Done". So the
refusal ends by saying the user has to agree first and has to be told, every
tool text that carries allow_network says the same, and every operation that
ran with allow_network and opened a workbook that can reach the network adds
network_allowed and network_note to its result.

NO test in the first half starts Excel; the workbook is made up by code.
The tests marked LIVE start the product's hidden Excel (on a private desktop)
and send their one request to a listener on 127.0.0.1.
"""
from __future__ import annotations

import http.server
import inspect
import re
import threading
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp import registry, server  # noqa: F401  (the import registers the ops)
from xlsx_mcp.com import session as com_session
from xlsx_mcp.core import netreach
from xlsx_mcp.ops import comtier

ASK_FIRST = ("Ask the user whether this workbook may contact the internet. Only if they agree, run the same "
             "call again with allow_network=true, and tell them it was allowed.")
PARAM_TEXT = ("Let Excel contact the internet for this workbook's network functions. Set it only after the "
              "user has agreed, for this workbook.")
NOTE = "Excel was allowed to contact the internet for this workbook because allow_network was set."

_CONSENT = re.compile(r"\b(ask|asks|asked|confirm\w*|agree\w*|consent\w*|permission)\b[^.]{0,80}\buser\b"
                      r"|\buser\b[^.]{0,80}\b(ask|asks|asked|confirm\w*|agree\w*|consent\w*|permission)\b",
                      re.IGNORECASE)


def _fetching(tmp_path: Path, name: str = "rates.xlsx") -> str:
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rates"
    ws["A1"] = "made-up"
    ws["B1"] = '=_xlfn.WEBSERVICE("http://127.0.0.1:9/made-up")'
    wb.save(p)
    return str(p)


@pytest.mark.parametrize("op", ["export_pdf", "recalculate", "validate that it opens clean"])
def test_the_network_refusal_says_the_user_must_agree_before_the_override(tmp_path, op):
    with pytest.raises(Exception) as ei:
        netreach.refuse_network_reach(_fetching(tmp_path), False, op)
    msg = str(ei.value)
    assert "allow_network=true" in msg  # the override is named
    assert _CONSENT.search(msg), msg  # and the user's agreement comes before it
    assert msg.endswith(ASK_FIRST), msg


def test_the_refusal_keeps_the_part_that_names_the_functions_and_says_nothing_was_opened(tmp_path):
    with pytest.raises(Exception) as ei:
        netreach.refuse_network_reach(_fetching(tmp_path), False, "export_pdf")
    msg = str(ei.value)
    assert "Rates!B1 (WEBSERVICE)" in msg
    assert "makes no network request you did not ask for" in msg and "so Excel was not started" in msg
    assert ei.value.detail["override"] == "allow_network=true"


def test_every_tool_text_that_carries_allow_network_says_the_user_agrees_first():
    carrying = []
    for name, spec in registry.public_ops().items():
        doc = " ".join((inspect.getdoc(spec.handler) or "").split())
        if "allow_network" in doc:
            carrying.append(name)
            assert PARAM_TEXT in doc, f"{name} describes allow_network without the user's agreement"
            assert "unless allow_network=true" not in doc, name
    # every op that takes the parameter carries the sentence
    takers = [n for n, s in registry.public_ops().items()
              if "allow_network" in inspect.signature(s.handler).parameters]
    assert sorted(carrying) == sorted(takers) and len(takers) == 10, (carrying, takers)


# ------------------------------------------------ the mechanism, without Excel


class _Book:
    Worksheets: list = []
    Names: list = []
    Connections = type("K", (), {"Count": 0})()

    def Close(self, *_a):
        pass


class _Books:
    Count = 0

    def Open(self, *_a):
        return _Book()


class _App:
    DisplayAlerts = False
    EnableEvents = False
    Calculation = -4105

    def __init__(self):
        self.Workbooks = _Books()


class _Executor:
    """Runs the job on this thread, as the COM worker would."""

    def submit(self, label, fn, timeout=None):
        class W:
            pid = 1
            app = _App()

        class M:
            def acquire(self):
                return W()

            def release_to_pool(self, _w):
                pass
        return fn(M())


@pytest.fixture
def standin(monkeypatch):
    monkeypatch.setattr(com_session, "get_executor", lambda: _Executor())
    monkeypatch.setattr(com_session, "com_available", lambda **_: (True, "a stand-in for the test"))
    monkeypatch.setattr(com_session, "_assert_hygiene", lambda app: ())
    monkeypatch.setattr(com_session, "_restore_hygiene", lambda app, *a: None)
    from xlsx_mcp import runtime
    monkeypatch.setattr(runtime, "excel_forbidden", lambda: None)


def _job_opening(path, **kw):
    def job(manager):
        w = manager.acquire()
        com_session.open_workbook(w.app, path, read_only=True, **kw)
        return {"ok": True}
    return job


def test_a_shared_open_with_allow_network_marks_the_result(tmp_path, standin):
    out = com_session.run_com("t", _job_opening(_fetching(tmp_path), allow_network=True))
    assert out == {"ok": True, "network_allowed": True, "network_note": NOTE}


def test_no_mark_without_allow_network_or_without_a_finding(tmp_path, standin):
    plain = tmp_path / "plain.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(plain)
    assert com_session.run_com("t", _job_opening(str(plain), allow_network=True)) == {"ok": True}
    assert com_session.run_com("t", _job_opening(str(plain))) == {"ok": True}


def test_the_mark_does_not_carry_over_to_the_next_job(tmp_path, standin):
    com_session.run_com("t", _job_opening(_fetching(tmp_path), allow_network=True))
    assert com_session.run_com("t", lambda manager: {"ok": True}) == {"ok": True}


def test_a_failed_job_leaves_no_mark_behind(tmp_path, standin):
    def boom(manager):
        com_session.open_workbook(manager.acquire().app, _fetching(tmp_path), read_only=True, allow_network=True)
        raise RuntimeError("the job failed")
    with pytest.raises(RuntimeError):
        com_session.run_com("t", boom)
    assert com_session.run_com("t", lambda manager: {"ok": True}) == {"ok": True}


def test_every_job_reaches_excel_through_the_one_wrapper():
    """The mark is made where a job is submitted and where the workbook is opened, so no op can skip it:
    nothing outside com.session submits a job, and no op opens a workbook except through open_workbook."""
    src = Path(inspect.getsourcefile(comtier)).parent.parent
    for py in src.rglob("*.py"):
        if py.name == "session.py" and py.parent.name == "com":
            continue
        text = py.read_text(encoding="utf-8")
        assert ".submit(" not in text, f"{py} submits a COM job itself"
        assert "app.Workbooks.Open(" not in text, f"{py} opens a workbook itself"
    session_text = Path(inspect.getsourcefile(com_session)).read_text(encoding="utf-8")
    assert session_text.count("get_executor().submit(") == 1
    assert "_with_network_note(fn)" in session_text


def test_a_rendered_sheet_keeps_the_mark_the_shared_open_made(tmp_path, monkeypatch):
    out = tmp_path / "a.png"

    def fake_run(label, path, body, **_kw):
        out.write_bytes(b"\x89PNG made-up")
        return {"rendered": str(out), "network_allowed": True, "network_note": NOTE}

    monkeypatch.setattr(comtier, "_run_readonly", fake_run)
    monkeypatch.setattr(comtier, "_clipboard_available", lambda: True)
    monkeypatch.setattr(comtier, "_render_is_blank", lambda _p: False)
    r = comtier.com_render_sheet(_fetching(tmp_path), str(out), allow_network=True)
    assert r["ok"] is True and r["network_allowed"] is True and r["network_note"] == NOTE


# ------------------------------------------------------------- live, per op


LIVE = pytest.mark.skipif(not com_session.com_available()[0], reason="needs Excel (COM)")


@pytest.fixture(scope="module")
def listener():
    hits: list[str] = []

    class H(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            hits.append(self.path)
            self.send_response(200)
            self.send_header("Content-Length", "2")
            self.end_headers()
            self.wfile.write(b"42")

        def log_message(self, *a):
            pass

    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield srv.server_address[1], hits
    srv.shutdown()


def _book(p, port, tag):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rates"
    ws["A1"] = "made-up-" + tag
    ws["A2"] = 3
    ws["B1"] = f'=_xlfn.WEBSERVICE("http://127.0.0.1:{port}/{tag}?v="&A1)'
    ws["B2"] = "=A2*2"
    ws["D1"], ws["E1"] = "Item", "Amount"
    ws["D2"], ws["E2"] = "made-up-x", 4
    ws["D3"], ws["E3"] = "made-up-y", 6
    wb.save(p)
    return str(p)


def _allowed_calls(tmp_path):
    out = tmp_path / "out"
    out.mkdir(exist_ok=True)
    k = {"allow_network": True}
    return {
        "com_validate_opens_clean": lambda p: comtier.com_validate_opens_clean(p, **k),
        "com_export_pdf": lambda p: comtier.com_export_pdf(p, str(out / "a.pdf"), **k),
        "com_render_sheet": lambda p: comtier.com_render_sheet(p, str(out / "a.png"), **k),
        "com_convert_format_xlsb": lambda p: comtier.com_convert_format(p, str(out / "b.xlsb"), **k),
        "com_convert_format_ods": lambda p: comtier.com_convert_format(p, str(out / "b.ods"), **k),
        "com_autofit": lambda p: comtier.com_autofit(p, **k),
        "com_save_with_password": lambda p: comtier.com_save_with_password(p, "made-up-pw", **k),
        "com_manage_pivot_list": lambda p: comtier.com_manage_pivot(p, "list", **k),
        "com_manage_pivot_create": lambda p: comtier.com_manage_pivot(
            p, "create", source_sheet="Rates", source_range="D1:E3", rows=["Item"],
            values=[{"field": "Amount", "func": "sum"}], **k),
        "com_set_sparkline_list": lambda p: comtier.com_set_sparkline(p, "list", **k),
        "com_set_sparkline_create": lambda p: comtier.com_set_sparkline(
            p, "create", location="G2", source="A2:B2", sheet="Rates", **k),
        "recalculate": lambda p: comtier.recalculate(p, engine="com", **k),
        "com_goal_seek": lambda p: comtier.com_goal_seek(p, "B2", 10, "A2", sheet="Rates", **k),
        "opens_clean": lambda p: com_session.opens_clean(p, allow_network=True),
    }


@LIVE
@pytest.mark.parametrize("op", ["com_validate_opens_clean", "com_export_pdf", "com_render_sheet",
                                "com_convert_format_xlsb", "com_convert_format_ods", "com_autofit",
                                "com_save_with_password", "com_manage_pivot_list", "com_manage_pivot_create",
                                "com_set_sparkline_list", "com_set_sparkline_create", "recalculate",
                                "com_goal_seek", "opens_clean"])
def test_every_xl_app_op_says_it_was_allowed_online(tmp_path, listener, op):
    port, hits = listener
    p = _book(tmp_path / "rates.xlsx", port, op)
    if op == "com_render_sheet" and not comtier._clipboard_available():
        # A render goes through the Windows clipboard, which a private desktop does not have: the product
        # refuses for that reason (named), and the carry-through is covered below with a stand-in.
        with pytest.raises(Exception) as ei:
            _allowed_calls(tmp_path)[op](p)
        assert "clipboard" in str(ei.value)
        return
    r = _allowed_calls(tmp_path)[op](p)
    assert r.get("network_allowed") is True, r
    assert r.get("network_note") == NOTE, r


@LIVE
def test_a_workbook_that_cannot_reach_the_network_gets_no_mark_even_when_allowed(tmp_path):
    p = tmp_path / "plain.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.active["B1"] = "=A1*2"
    wb.save(p)
    r = comtier.com_validate_opens_clean(str(p), allow_network=True)
    assert "network_allowed" not in r and "network_note" not in r
