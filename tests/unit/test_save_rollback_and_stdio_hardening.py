"""Excel saves, CI bounds, served hints, the pack-free surface and stdio
behaviour on the Excel family surface:

  * com_save_with_password, and every Excel save through _run_mutation,
    put the original back when Excel's save fails after the commit
    point or the encryption did not take
  * every CI job is bounded (timeout-minutes), not only tests.yml
  * no served hint offers live='auto'; Excel has no live route
  * the 24-hour pack expiry cannot happen: there is no pack and no
    visibility state on the family surface
  * unknown methods before initialize are answered with a JSON-RPC
    error and the server stays alive, even when the client never
    reads stderr

The four silent data-loss defects of released 1.2.5 have their own file,
test_silent_data_loss_defects.py, written against the engine so the same file
runs red on 1.2.5 and green here.
"""

from __future__ import annotations

import ast
import hashlib
import json
import os
import queue
import subprocess
import sys
import threading
import time
import types
from pathlib import Path

import openpyxl
import pytest

import family_harness as fh
from xlsx_mcp import envelope
from xlsx_mcp import server as srv
from xlsx_mcp.com import session as com_session
from xlsx_mcp.core import safesave
from xlsx_mcp.core.errors import ValidationFailed, XlMcpError
from xlsx_mcp.ops import comtier

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src" / "xlsx_mcp"
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _sha(p) -> str:
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


@pytest.fixture
def book(tmp_path):
    p = tmp_path / "b.xlsx"
    srv.create_workbook(str(p))
    srv.write_range(str(p), "A1", [["Name", "Qty"], ["Alpha", 1]])
    return str(p)


# ======================================================================
# A failed Excel save after the commit point never leaves the file
# changed without the backup put back
# ======================================================================


class _FakeWorker:
    pid = 0
    app = object()


class _FakeManager:
    def acquire(self):
        return _FakeWorker()

    def release_to_pool(self, w):
        pass


class _FakeWB:
    """A workbook Excel opened. mode decides what its save writes: 'plain'
    a different plain package (the password did not take); 'raise' half a
    file, and then it fails, as a save cut off midway. Excel now
    writes a copy (SaveCopyAs) and the file is replaced only after
    Excel has closed it, so a save cut off midway never reaches the file."""

    def __init__(self, path: str, mode: str):
        import os
        self.path, self.mode = path, mode
        self.Password = ""
        # A save lands only on the workbook Excel says it is
        self.FullName = os.path.abspath(path)

    def Save(self):
        raise AssertionError("wb.Save lets Excel replace the file itself")

    def SaveCopyAs(self, copy):
        if self.mode == "plain":
            wb = openpyxl.load_workbook(self.path)
            wb.active["Z9"] = "Excel touched this"
            wb.save(copy)
        elif self.mode == "raise":
            Path(copy).write_bytes(b"half a file")
            raise OSError("the disk went away during the save")

    def Close(self, save):
        pass


def _fake_excel(monkeypatch, mode: str):
    monkeypatch.setattr(com_session, "run_com",
                        lambda label, job, timeout=None: job(_FakeManager()))
    monkeypatch.setattr(com_session, "_assert_hygiene", lambda app: ())
    monkeypatch.setattr(com_session, "_restore_hygiene", lambda app, *a: None)
    monkeypatch.setattr(com_session, "open_workbook",
                        lambda app, p, password=None, **k: _FakeWB(p, mode))
    monkeypatch.setattr(com_session, "com_retry", lambda fn, label=None, **k: fn())


def test_941_encryption_that_did_not_take_never_replaces_the_original(book, monkeypatch):
    """Excel's copy is checked before it replaces the workbook, so the
    unencrypted copy never does; nothing needs putting back and the slots
    stay as they were."""
    before = _sha(book)
    slot = safesave.slot_dir(book) / safesave.PREV_SLOT
    slot_before = _sha(slot) if slot.exists() else None
    _fake_excel(monkeypatch, "plain")
    with pytest.raises(ValidationFailed) as ei:
        comtier.com_save_with_password(book, "secret-1")
    assert _sha(book) == before, "the file Excel rewrote unencrypted was left in place"
    assert "password did not take" in str(ei.value)
    assert "was not changed" in str(ei.value)
    assert (_sha(slot) if slot.exists() else None) == slot_before, "the slots changed"


def _left_beside(book) -> list:
    return sorted(n for n in os.listdir(os.path.dirname(book))
                  if n not in (os.path.basename(book), safesave.BACKUP_DIR_NAME))


def test_941_a_save_that_fails_after_the_commit_leaves_the_file_as_it_was(book, monkeypatch):
    before = _sha(book)
    _fake_excel(monkeypatch, "raise")
    with pytest.raises(XlMcpError) as ei:
        comtier.com_save_with_password(book, "secret-1")
    assert _sha(book) == before, "a save cut off midway left a damaged file"
    assert "was not changed" in str(ei.value)
    assert _left_beside(book) == []


def test_941_the_same_holds_for_every_save_through_run_mutation(book, monkeypatch):
    """The class: _run_mutation (recalculate, com_autofit, com_goal_seek,
    com_set_sparkline, com_manage_pivot) saves through Excel the same way."""
    before = _sha(book)
    slot = safesave.slot_dir(book) / safesave.PREV_SLOT
    slot_before = _sha(slot) if slot.exists() else None
    _fake_excel(monkeypatch, "raise")
    with pytest.raises(XlMcpError) as ei:
        comtier._run_mutation("autofit(b.xlsx)", book, lambda app, wb: {"x": 1})
    assert _sha(book) == before
    assert "was not changed" in str(ei.value)
    assert (_sha(slot) if slot.exists() else None) == slot_before, "prev was burned"
    assert _left_beside(book) == []


def test_941_without_a_backup_a_failed_save_still_leaves_the_file_as_it_was(book, monkeypatch):
    before = _sha(book)
    _fake_excel(monkeypatch, "raise")
    with pytest.raises(XlMcpError) as ei:
        comtier._run_mutation("autofit(b.xlsx)", book, lambda app, wb: {"x": 1},
                              backup=False)
    assert _sha(book) == before
    assert "was not changed" in str(ei.value)


def test_941_a_failure_before_the_commit_point_leaves_the_slots_alone(book, monkeypatch):
    """Nothing reached the disk, so nothing is restored and prev is not
    burned (the two-phase rotation's covenant)."""
    _fake_excel(monkeypatch, "plain")
    before = _sha(book)
    slot = safesave.slot_dir(book) / safesave.PREV_SLOT
    slot_before = _sha(slot) if slot.exists() else None

    def body(app, wb):
        raise XlMcpError("the body refused")

    with pytest.raises(XlMcpError, match="the body refused"):
        comtier._run_mutation("autofit(b.xlsx)", book, body)
    assert _sha(book) == before
    assert (_sha(slot) if slot.exists() else None) == slot_before


# ======================================================================
# Every CI job carries timeout-minutes
# ======================================================================


def test_934_every_job_in_every_workflow_is_bounded():
    import yaml

    missing = []
    for wf in sorted((ROOT / ".github" / "workflows").glob("*.yml")):
        jobs = (yaml.safe_load(wf.read_text("utf-8")) or {}).get("jobs") or {}
        for name, job in jobs.items():
            if "uses" in job:  # a reusable workflow carries its own bound
                continue
            minutes = job.get("timeout-minutes")
            if not isinstance(minutes, int) or not 0 < minutes <= 60:
                missing.append(f"{wf.name}:{name}={minutes!r}")
    assert not missing, f"jobs without a bound: {missing}"


# ======================================================================
# No served hint names live='auto'
# ======================================================================


def test_935_no_hint_offers_a_live_route():
    for code, hint in envelope.HINTS.items():
        assert "live" not in hint.lower(), (code, hint)
    assert "live" not in envelope.RECOVERY_HINT.lower()


def test_935_a_locked_workbook_refusal_names_no_live_route(book):
    with open(book, "r+b"):
        pass
    from xlsx_mcp.core import package
    msg = package.held_refusal(book)
    assert "live" not in msg.lower()
    assert "WORKBOOK_LOCKED" in envelope.HINTS


# ======================================================================
# No pack, no visibility state, nothing that can expire
# ======================================================================


def test_933_the_family_surface_has_no_visibility_state_that_could_expire():
    """1.2.5 switched packs on with fastmcp's per-session
    ctx.enable_components; fastmcp forgot that state after 24 hours. The
    family surface lists every tool from the start and keeps no per-session
    visibility record at all, so there is nothing to expire."""
    assert not (SRC / "packs.py").exists()
    banned = ("enable_components", "disable_components", "set_visibility_hook",
              "_visibility_hook", "enable_tools", "disable_tools")
    hits = []
    for py in SRC.rglob("*.py"):
        tree = ast.parse(py.read_text("utf-8"))
        for node in ast.walk(tree):
            name = (node.attr if isinstance(node, ast.Attribute)
                    else node.id if isinstance(node, ast.Name)
                    else node.name if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
                    else None)
            if name in banned:
                hits.append(f"{py.relative_to(SRC)}:{node.lineno}:{name}")
    assert not hits, hits


def test_933_every_tool_is_listed_and_callable_from_the_first_message():
    h = fh.start()
    names = [t.name for t in h.list_tools()]
    assert len(names) == 7
    payload, err, _ = h.call("xl_read", {"op": "get_server_info", "args": {}})
    assert not err
    assert payload["surface"]["ops"] == 69


# ======================================================================
# Unknown methods before initialize, stderr never read
# ======================================================================


def _rpc_session(messages, drain_stderr: bool, wait_for: int, timeout: float = 60.0):
    proc = subprocess.Popen(fh.launch_cmd(), stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, env=fh._env(None),
                            creationflags=NO_WINDOW)
    got: dict = {}
    q: queue.Queue = queue.Queue()

    def reader():
        for line in proc.stdout:
            try:
                q.put(json.loads(line))
            except ValueError:
                pass

    threading.Thread(target=reader, daemon=True).start()
    if drain_stderr:
        threading.Thread(target=proc.stderr.read, daemon=True).start()

    def writer():
        # A server that stopped reading stdin must fail the test, not hang it.
        try:
            for m in messages:
                proc.stdin.write((json.dumps(m) + "\n").encode())
                proc.stdin.flush()
        except (OSError, ValueError):
            pass

    threading.Thread(target=writer, daemon=True).start()
    try:
        end = time.monotonic() + timeout
        while wait_for not in got and time.monotonic() < end:
            try:
                msg = q.get(timeout=0.5)
            except queue.Empty:
                continue
            if "id" in msg:
                got[msg["id"]] = msg
    finally:
        if wait_for not in got:
            proc.kill()  # silent: nothing more will come
        try:
            proc.stdin.close()
        except OSError:
            pass
        try:
            proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait(timeout=30)
    return got


def _probe(n_unknown: int) -> list[dict]:
    msgs = [{"jsonrpc": "2.0", "id": 1, "method": "server/discover", "params": {}}]
    msgs += [{"jsonrpc": "2.0", "id": i, "method": f"nope/unknown{i}", "params": {}}
             for i in range(2, n_unknown + 1)]
    msgs.append({"jsonrpc": "2.0", "id": 100000, "method": "initialize", "params": {
        "protocolVersion": "2025-06-18", "capabilities": {},
        "clientInfo": {"name": "p948", "version": "0"}}})
    msgs.append({"jsonrpc": "2.0", "method": "notifications/initialized"})
    msgs.append({"jsonrpc": "2.0", "id": 100001, "method": "tools/call", "params": {
        "name": "xl_read", "arguments": {"op": "get_workflows", "args": {}}}})
    return msgs


@pytest.mark.timeout(180)
@pytest.mark.parametrize("drain", [False, True], ids=["stderr-never-read", "stderr-read"])
def test_948_unknown_methods_before_initialize_are_answered_and_the_server_stays_up(drain):
    n = 120  # An earlier build went silent after about 36 with stderr unread
    got = _rpc_session(_probe(n), drain_stderr=drain, wait_for=100001)
    assert 100000 in got, f"initialize never answered ({len(got)} answers came back)"
    assert "result" in got[100000]
    assert 100001 in got and "result" in got[100001], "the server died after initialize"
    errors = [got[i] for i in range(1, n + 1) if i in got]
    assert len(errors) == n, f"only {len(errors)} of {n} unknown requests were answered"
    assert all("error" in e and isinstance(e["error"].get("code"), int) for e in errors)


def test_948_the_log_writer_drops_lines_instead_of_waiting():
    """The unit form of the guarantee: a stderr that never accepts a byte
    stalls only the writer thread; emitting stays instant."""
    import logging

    class _Stuck:
        def write(self, s):
            threading.Event().wait()  # never returns

        def flush(self):
            pass

    w = srv._NonBlockingStderr(stream=_Stuck(), maxsize=8)
    log = logging.getLogger("ks4xl-test-948")
    log.propagate = False
    log.addHandler(w.handler)
    try:
        t0 = time.monotonic()
        for i in range(200):
            log.warning("line %d", i)
        assert time.monotonic() - t0 < 2.0
        assert w.dropped >= 200 - 8 - 1
    finally:
        log.removeHandler(w.handler)


@pytest.mark.timeout(120)
def test_948_the_server_exits_when_stdin_closes_while_stderr_is_full():
    """With stderr never read and the
    pipe full, the writer thread used to sit in sys.stderr.write holding the
    stream's lock; at shutdown Python waited on that lock, wrote a fatal
    error to the same full pipe and never exited. The writer now uses its
    own descriptor, so closing stdin ends the server promptly and cleanly."""
    proc = subprocess.Popen(fh.launch_cmd(), stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, env=fh._env(None),
                            creationflags=NO_WINDOW)
    threading.Thread(target=proc.stdout.read, daemon=True).start()  # stdout drained
    try:
        for m in _probe(150)[:150]:  # unknown methods: one log line each, >4 KB
            proc.stdin.write((json.dumps(m) + "\n").encode())
        proc.stdin.flush()
        time.sleep(3)  # let the writer fill the pipe and block
        proc.stdin.close()
        code = proc.wait(timeout=20)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.wait(timeout=30)
        pytest.fail("the server did not exit within 20 s of stdin closing")
    err = proc.stderr.read().decode("utf-8", "replace")
    assert code == 0, (code, err[-500:])
    assert "Fatal Python error" not in err


def _reading_client(cmd: list[str], env: dict) -> tuple[int, str]:
    """A client that reads stderr to the end, as Claude Desktop does."""
    proc = subprocess.run(cmd, env=env, stdin=subprocess.DEVNULL, capture_output=True,
                          timeout=120, creationflags=NO_WINDOW)
    return proc.returncode, proc.stderr.decode("utf-8", "replace")


@pytest.mark.timeout(600)
def test_948_f8_the_retired_setting_refusal_reaches_a_reading_client_every_time():
    """Output written just
    before an exit was lost to a client that reads stderr. Excel's refusal
    to start on a retired 1.x lock setting is that output: ten starts, ten
    whole refusal lines."""
    for _ in range(10):
        code, err = _reading_client(fh.launch_cmd(), fh._env({"KS4XL_PACK_POLICY": "locked"}))
        assert code == 2 and "KS4XL_PACK_POLICY is set to a value that limited" in err, (code, err)
        assert err.rstrip().endswith("The server did not start."), err


_CRASH = (
    "from ks4xl_community_launch import build_policy\n"
    "from xlsx_mcp import license_policy\n"
    "license_policy.install(build_policy())\n"
    "import logging\n"
    "from xlsx_mcp import server\n"
    "server._bound_stderr_logging()\n"
    "for i in range(40):\n"
    "    logging.getLogger('probe').warning('queued line %d', i)\n"
    "raise RuntimeError('the last line')\n")


@pytest.mark.timeout(600)
def test_948_f8_the_last_lines_before_a_crash_reach_a_reading_client():
    """F8, the crash half: log lines still in the writer's queue when the
    process dies, and the exception's own last line, all arrive."""
    for _ in range(10):
        code, err = _reading_client([sys.executable, "-X", "utf8", "-c", _CRASH],
                                    fh._env(None))
        assert code == 1, (code, err[-400:])
        lines = [line.rstrip("\r") for line in err.splitlines()]
        # Every queued line, whole, and the traceback's last line, whole (the
        # writer's lines may land before or after the traceback).
        assert [f"WARNING:probe:queued line {i}" for i in range(40)] == \
            [line for line in lines if line.startswith("WARNING:probe:")], err[-400:]
        assert "RuntimeError: the last line" in lines, err[-400:]


def test_948_last_resort_writers_go_through_logging(monkeypatch, caplog):
    """An exception that ends a thread, an
    unraisable exception and a warning reach logging, not sys.stderr."""
    import logging
    import warnings

    monkeypatch.setattr(sys, "unraisablehook", sys.unraisablehook)
    monkeypatch.setattr(threading, "excepthook", threading.excepthook)
    # Warning capture is process-wide state that a library imported earlier
    # in the run may already have switched on; start from off, and let
    # monkeypatch put both pieces of state back afterwards.
    monkeypatch.setattr(warnings, "showwarning", warnings.showwarning)
    monkeypatch.setattr(logging, "_warnings_showwarning", logging._warnings_showwarning)
    logging.captureWarnings(False)
    srv._hooks_to_logging()
    with caplog.at_level(logging.WARNING):
        t = threading.Thread(target=lambda: 1 / 0, name="probe-thread")
        t.start()
        t.join()
        sys.unraisablehook(types.SimpleNamespace(
            exc_type=ValueError, exc_value=ValueError("late"), exc_traceback=None,
            err_msg="Exception ignored in", object=None))
        # pytest records warnings.warn itself, so the hook is called the way
        # the warnings machinery calls it.
        warnings.showwarning("probe warning", UserWarning, "probe.py", 1)
    text = caplog.text
    assert "probe-thread ended on ZeroDivisionError" in text
    assert "unraisable ValueError: late" in text
    assert "probe warning" in text


# ======================================================================
# KS4XL_FORBID=xl_app closes every route to Excel
# ======================================================================


def _writers_with_verify_com() -> list[str]:
    from xlsx_mcp import registry
    return sorted(n for n, s in registry.public_ops().items()
                  if "verify_com" in s.args.model_fields)


@pytest.fixture
def no_excel(monkeypatch):
    """Any attempt to start Excel fails the test."""
    started = []

    def submit(self, label, fn, timeout=None):
        started.append(label)
        raise AssertionError(f"Excel was started for {label}")

    monkeypatch.setattr(com_session.ComExecutor, "submit", submit)
    return started


def test_forbid_xl_app_refuses_verify_com_on_every_writer(book, no_excel):
    from xlsx_mcp import registry
    from xlsx_mcp.settings import Settings
    from test_op_sweep import _args_for
    from pathlib import Path as _P

    names = _writers_with_verify_com()
    assert len(names) >= 30, names  # 34 at 0268d8b
    h = fh.start(Settings(forbid=("xl_app",)))
    work = _P(book).parent
    for name in names:
        d = work / name
        d.mkdir()
        args = _args_for(name, d, _P(book))
        args["verify_com"] = True
        payload, err, _ = h.call(registry.REGISTRY[name].group, {"op": name, "args": args})
        assert err and payload["error"]["code"] == "NOT_AVAILABLE", (name, payload)
        assert "KS4XL_FORBID" in payload["error"]["message"], payload
    assert no_excel == []


def test_forbid_xl_app_wins_over_ks4xl_verify_com(book, no_excel, monkeypatch):
    from xlsx_mcp.settings import Settings
    monkeypatch.setenv("KS4XL_VERIFY_COM", "1")
    h = fh.start(Settings(forbid=("xl_app",)))
    payload, err, _ = h.call("xl_edit", {"op": "set_cell", "args": {
        "path": book, "location": "C1", "value": 1}})
    assert not err, payload
    assert "verified_com" not in payload
    info, _, _ = h.call("xl_read", {"op": "get_server_info", "args": {}})
    assert info["config"]["verify_com_default"] is False and "KS4XL_FORBID" in info["excel"]
    payload, err, _ = h.call("xl_read", {"op": "validate", "args": {"path": book}})
    assert not err, payload
    assert no_excel == []


def test_without_the_forbid_verify_com_still_reaches_excel(book, no_excel, monkeypatch):
    """The guard is the setting, not a blanket off switch (GAP guard)."""
    monkeypatch.setattr(com_session, "com_available", lambda refresh=False: (True, "test"))
    h = fh.start()
    payload, err, _ = h.call("xl_edit", {"op": "set_cell", "args": {
        "path": book, "location": "C1", "value": 1, "verify_com": True}})
    assert no_excel, "verify_com never tried to start Excel"


@pytest.mark.timeout(120)
def test_forbid_xl_app_over_stdio(book):
    res = fh.stdio_session([
        ("xl_edit", {"op": "set_cell", "args": {"path": book, "location": "C1", "value": 1,
                                                 "verify_com": True}}),
        ("xl_read", {"op": "get_server_info", "args": {}}),
    ], env={"KS4XL_FORBID": "xl_app", "KS4XL_VERIFY_COM": "1"})
    first = json.loads(res["results"][0]["text"])
    assert first["error"]["code"] == "NOT_AVAILABLE" and "verify_com" in first["error"]["message"]
    info = json.loads(res["results"][1]["text"])
    assert info["config"]["verify_com_default"] is False
