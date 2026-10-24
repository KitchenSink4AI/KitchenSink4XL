"""The update notice and the update the user chooses (core/update_flow.py).

Nothing here reaches PyPI or GitHub and nothing opens on the desktop: every
test runs on its own state folder with a landmine on the PyPI fetch, the
bundle download and os.startfile, and an install child is a fake unless the
test builds a local package index of its own (the uvx probe at the end,
which talks to 127.0.0.1 only).
"""

from __future__ import annotations

import asyncio
import hashlib
import http.server
import io
import json
import os
import shutil
import subprocess
import sys
import threading
import time
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace

import pytest

from xlsx_mcp import server
from xlsx_mcp.core import update_check as uc
from xlsx_mcp.core import update_flow as uf

ROOT = Path(__file__).resolve().parents[2]
NEW = "999.0.1"
_REAL_FETCH = uc._fetch
_REAL_ROUTE = uf.install_route
_REAL_SPAWN = uf._spawn


def _landmine(*a, **k):
    raise AssertionError("this path must not reach the network or the desktop")


@pytest.fixture(autouse=True)
def _isolated(tmp_path, monkeypatch):
    for name in (uc.OFF_ENV, uc.LEGACY_OFF_ENV, uc.NOTICES_ENV, uc.INSTALL_ROUTE_ENV, "UV"):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv(uc.CACHE_DIR_ENV, str(tmp_path / "state"))
    monkeypatch.setattr(uc, "_fetch", _landmine)
    monkeypatch.setattr(uf, "_download", _landmine)
    monkeypatch.setattr(uf, "_open_bundle", _landmine)
    if hasattr(os, "startfile"):
        monkeypatch.setattr(os, "startfile", _landmine)
    monkeypatch.setattr(uf, "_CLAIMED_AT", None)
    monkeypatch.setattr(uf, "_JOB", None)
    monkeypatch.setattr(uf, "_DIALOG_UNUSABLE", None)
    monkeypatch.setattr(uf, "_SAID", set())
    monkeypatch.setattr(uc, "_REFRESH_THREAD", None)
    monkeypatch.setattr(uc, "_REFRESH_STARTED_AT", None)
    # The dialog is offered only on a route that acts; no child may start
    # unless a test installs its own fake.
    monkeypatch.setattr(uf, "install_route", lambda: "uvx")
    monkeypatch.setattr(uf, "_spawn", _landmine)
    yield
    job = uf._JOB
    if job is not None and job.proc is not None:
        job.proc.kill()


def _cache(latest=NEW, age_days=0.0, ok=True):
    stamp = datetime.now(timezone.utc) - timedelta(days=age_days)
    uc.write_cache({"last_check": stamp.isoformat(), "latest_version": latest, "ok": ok,
                    "last_success": stamp.isoformat(), "disclosed": True})


def _payload(*versions):
    return {"info": {"version": versions[-1]},
            "releases": {v: [{"filename": f"x-{v}.whl", "yanked": False}] for v in versions}}


class Dialogs:
    """An elicitation handler that records every dialog and answers."""

    def __init__(self, answer=None, delay=0.0):
        self.seen: list = []
        self.answer = answer if answer is not None else {"choice": uf.NOT_NOW}
        self.delay = delay

    async def __call__(self, message, response_type, params, context):
        self.seen.append(params)
        if self.delay:
            await asyncio.sleep(self.delay)
        if isinstance(self.answer, Exception):
            raise self.answer
        return self.answer


def _session(calls, handler=None, settings=None, parallel=False):
    """Calls in one in-process MCP session; [(payload, is_error)]."""
    from fastmcp import Client

    mcp = server.create_server(settings)

    async def go():
        kwargs = {"elicitation_handler": handler} if handler is not None else {}
        async with Client(mcp, **kwargs) as client:
            async def one(tool, args):
                res = await client.call_tool_mcp(tool, args)
                return json.loads(res.content[0].text), bool(res.isError)
            if parallel:
                return list(await asyncio.gather(*(one(t, a) for t, a in calls)))
            return [await one(t, a) for t, a in calls]

    return asyncio.run(go())


HELP = ("xl_help", {})


def _read(doc):
    return ("xl_read", {"op": "get_workbook_metadata", "args": {"path": doc}})


@pytest.fixture
def doc(tmp_path):
    path = str(tmp_path / "d.xlsx")
    server.create_workbook(path)
    return path


# ------------------------------------------------------- 1. background check


def test_background_check_never_blocks_the_call_and_starts_once(monkeypatch):
    calls, release = [], threading.Event()

    def slow_fetch(url=uc.PYPI_URL, timeout=uc.TIMEOUT_SECONDS):
        calls.append((timeout, threading.current_thread().name))
        release.wait(60)
        return _payload("1.0.0", NEW)

    monkeypatch.setattr(uc, "_fetch", slow_fetch)
    (out, err), = _session([HELP])
    # No wall-clock bound here: refresh_running() below proves the call did not
    # wait for the fetch, and a loaded machine can make the session slow.
    assert not err
    assert uc.refresh_running()
    assert uc.maybe_refresh_in_background() is False        # already running
    release.set()
    uc._REFRESH_THREAD.join(10)
    assert calls == [(uc.BACKGROUND_TIMEOUT_SECONDS, f"{uc.LOGGER_NAME}-update-check")]
    assert uc.read_cache()["latest_version"] == NEW
    _cache(age_days=8)                                       # due again, same week
    assert uc.maybe_refresh_in_background() is False         # started < 7 days ago here


def test_on_demand_timeout_is_unchanged(monkeypatch):
    seen = []

    def urlopen(req, timeout=None):
        seen.append(timeout)
        raise OSError("down")

    monkeypatch.setattr(uc, "_fetch", _REAL_FETCH)
    monkeypatch.setattr(uc.urllib.request, "urlopen", urlopen)
    uc.status()
    assert seen == [uc.TIMEOUT_SECONDS] == [2.0]


@pytest.mark.parametrize("off", ["check", "legacy", "policy"])
def test_off_switches_stop_the_thread_and_the_notice(off, monkeypatch, doc):
    if off == "check":
        monkeypatch.setenv(uc.OFF_ENV, "off")
    elif off == "legacy":
        monkeypatch.setenv(uc.LEGACY_OFF_ENV, "1")
    else:
        monkeypatch.setattr(uc.license_policy, "update_check_suppression",
                            lambda: {"disabled_by": "policy"})
    _cache(age_days=30)
    handler = Dialogs()
    results = _session([_read(doc), HELP], handler)
    assert uc._REFRESH_THREAD is None and handler.seen == []
    assert all("update" not in out for out, _ in results)


def test_fetch_errors_are_survived(monkeypatch):
    monkeypatch.setattr(uc, "_fetch", lambda *a, **k: (_ for _ in ()).throw(OSError("down")))
    (out, err), = _session([HELP])
    uc._REFRESH_THREAD.join(5)
    assert not err and uc.read_cache()["ok"] is False


def test_status_does_not_fetch_again_while_the_background_check_runs(monkeypatch):
    gate = threading.Event()
    count = []

    def fetch(*a, **k):
        count.append(1)
        gate.wait(10)
        return _payload(NEW)

    monkeypatch.setattr(uc, "_fetch", fetch)
    assert uc.maybe_refresh_in_background()
    uc.status()
    gate.set()
    uc._REFRESH_THREAD.join(5)
    assert len(count) == 1


def test_the_process_exits_promptly_while_a_fetch_hangs(tmp_path):
    code = ("import sys, time; sys.path[:0] = {paths!r}; import conftest; "
            "from xlsx_mcp.core import update_check as uc; "
            "uc._fetch = lambda *a, **k: time.sleep(120); "
            "assert uc.maybe_refresh_in_background(); print('started')")
    env = {**os.environ, uc.CACHE_DIR_ENV: str(tmp_path / "s"), uc.OFF_ENV: "on",
           "PYTHONPATH": str(ROOT / "src")}
    started = time.monotonic()
    res = subprocess.run([sys.executable, "-X", "utf8", "-c", code.format(
        paths=[str(ROOT / "src"), str(ROOT / "tests"), str(ROOT / "tests" / "unit")])],
        env=env, capture_output=True, text=True, timeout=60, stdin=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    assert res.returncode == 0 and "started" in res.stdout, res.stderr[-800:]
    assert time.monotonic() - started < 30


# ---------------------------------------------------------------- 2. dialog


def test_a_read_call_shows_the_dialog_with_the_exact_wire_shape(doc):
    _cache()
    handler = Dialogs()
    (out, err), = _session([_read(doc)], handler)
    assert not err and out["ok"] is True and "update" not in out
    (params,) = handler.seen
    assert params.message == (f"{uc.PRODUCT} {NEW} is available. You have "
                              f"{uc.current_version()}. Update now?")
    schema = params.requestedSchema
    assert schema["required"] == ["choice"] and list(schema["properties"]) == ["choice"]
    choice = schema["properties"]["choice"]
    assert choice["enum"] == ["Update now", "Not now"]
    assert choice["description"] == ("You can turn update notices off at any time: tell your "
                                     f"AI to turn off {uc.PRODUCT} update notices.")
    assert "http" not in params.message


def _backed_off():
    last = datetime.fromisoformat(uf.read_state()["last_attempt"])
    return datetime.now(timezone.utc) - last < timedelta(minutes=1)


def test_not_now_backs_off_a_week_and_never_fails_the_call(doc):
    _cache()
    handler = Dialogs({"choice": uf.NOT_NOW})
    results = _session([_read(doc), _read(doc), HELP], handler)
    assert len(handler.seen) == 1
    assert all(not err and "update" not in out for out, err in results)
    assert _backed_off() and uf._JOB is None
    assert uf.read_state()["last_dialog"] == "answered"


@pytest.mark.parametrize("action", ["decline", "cancel"])
def test_a_slow_decline_is_the_users_answer(action, monkeypatch, doc):
    """A decline or cancel 2 s after the request (fake clock)."""
    from fastmcp.client.elicitation import ElicitResult

    clock = iter([100.0, 102.0])
    monkeypatch.setattr(uf, "_monotonic", lambda: next(clock))
    _cache()
    (out, err), = _session([_read(doc)], Dialogs(ElicitResult(action=action)))
    assert not err and "update" not in out and _backed_off()
    assert uf._DIALOG_UNUSABLE is None


@pytest.mark.parametrize("answer,reason", [
    ("decline", "auto-decline"), ("cancel", "auto-decline"),
    (RuntimeError("client broke"), "error"), ({"choice": "Something else"}, "error")])
def test_an_unusable_dialog_falls_back_on_the_same_result(answer, reason, doc):
    """An instant decline or cancel, an error, an invalid answer."""
    from fastmcp.client.elicitation import ElicitResult

    _cache()
    if answer in ("decline", "cancel"):
        answer = ElicitResult(action=answer)
    handler = Dialogs(answer)
    results = _session([_read(doc), _read(doc)], handler)
    assert len(handler.seen) == 1
    assert results[0][0]["update"]["relay"] == uf.RELAY_OFFER.format(tool=uc.UPDATE_TOOL)
    assert "update" not in results[1][0]
    assert _backed_off() and uf._DIALOG_UNUSABLE == reason
    assert uf.read_state()["last_dialog"] == f"unavailable:{reason}"


def test_after_an_unusable_dialog_the_process_goes_straight_to_the_key(monkeypatch, doc):
    """The next notice in the same process does not elicit."""
    from fastmcp.client.elicitation import ElicitResult

    _cache()
    handler = Dialogs(ElicitResult(action="decline"))
    _session([_read(doc)], handler)
    monkeypatch.setattr(uf, "_CLAIMED_AT", None)               # a week later
    uf._record(last_attempt=(datetime.now(timezone.utc) - timedelta(days=8)).isoformat())
    (out, _), = _session([_read(doc)], handler)
    assert len(handler.seen) == 1 and "update" in out


def test_no_fallback_room_records_nothing_and_shows_no_dialog(monkeypatch, doc):
    _cache()
    monkeypatch.setattr(uf, "_fits", lambda *a: False)
    handler = Dialogs()
    _session([_read(doc)], handler)
    assert handler.seen == [] and "last_attempt" not in uf.read_state()


def test_a_write_never_gets_the_dialog_a_read_does(tmp_path):
    _cache()
    handler = Dialogs()
    path = str(tmp_path / "w.xlsx")
    results = _session([("xl_workbook", {"op": "create_workbook",
                                         "args": {"path": path}}),
                        ("xl_edit", {"op": "set_cell", "args": {
                            "path": path, "location": "A1", "value": "x"}})], handler)
    assert handler.seen == [] and all("update" not in out for out, _ in results)
    _session([_read(path)], handler)
    assert len(handler.seen) == 1


def test_two_parallel_calls_give_one_dialog(doc):
    _cache()
    handler = Dialogs(delay=0.3)
    _session([_read(doc), _read(doc), HELP], handler, parallel=True)
    assert len(handler.seen) == 1


def test_a_late_answer_does_not_break_the_session(monkeypatch, doc):
    _cache()
    monkeypatch.setattr(uf, "DIALOG_TIMEOUT_S", 0.3)
    handler = Dialogs({"choice": uf.UPDATE_NOW}, delay=1.0)
    started = time.monotonic()
    results = _session([_read(doc), HELP, _read(doc)], handler)
    assert all(not err for _, err in results)
    assert time.monotonic() - started < 15
    assert uf._JOB is None                       # the late "Update now" was dropped
    assert "update" in results[0][0]             # a timeout puts the key on the same result
    assert uf._DIALOG_UNUSABLE == "timeout"


def test_a_client_cancel_during_the_dialog_propagates():
    class Ctx:
        session = SimpleNamespace(_request_id=7)

        async def elicit(self, *a, **k):
            await asyncio.sleep(30)

    async def go():
        task = asyncio.ensure_future(uf.ask(Ctx(), "1.0", "2.0"))
        await asyncio.sleep(0.05)
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task

    asyncio.run(go())


def _ctx(elicitation=..., protocol=max(uf.DIALOG_PROTOCOLS), background=False,
         name="claude-code"):
    from mcp import types

    caps = types.ClientCapabilities(
        elicitation=None if elicitation is None else (
            types.ElicitationCapability() if elicitation is ... else elicitation))
    params = types.InitializeRequestParams(protocolVersion=protocol, capabilities=caps,
                                           clientInfo=types.Implementation(name=name, version="1"))
    return SimpleNamespace(session=SimpleNamespace(client_params=params),
                           is_background_task=background)


def test_the_dialog_gate():
    from mcp import types

    assert uf._session_facts(_ctx()) == (True, "claude-code")                  # {} = form
    assert uf._session_facts(_ctx(types.ElicitationCapability(
        form=types.FormElicitationCapability())))[0]
    assert not uf._session_facts(_ctx(types.ElicitationCapability(
        url=types.UrlElicitationCapability())))[0]                            # URL only
    assert not uf._session_facts(_ctx(None))[0]                               # not declared
    from mcp.shared.version import SUPPORTED_PROTOCOL_VERSIONS as known

    older = [v for v in known if v < min(uf.DIALOG_PROTOCOLS)]
    assert older and not any(uf._session_facts(_ctx(protocol=v))[0] for v in older)
    assert uf._session_facts(_ctx(protocol=min(uf.DIALOG_PROTOCOLS)))[0]
    assert not uf._session_facts(_ctx(protocol="a-later-era"))[0]           # unknown era
    assert not uf._session_facts(_ctx(background=True))[0]
    assert uf._session_facts(object()) == (False, "")


# ------------------------------------------------------------- 3. fallback


def test_no_dialog_capability_gets_the_update_key_once_a_week(doc):
    _cache()
    results = _session([_read(doc), _read(doc), HELP])
    first = results[0][0]
    assert first["update"] == {
        "notice": f"{uc.PRODUCT} {NEW} is available (you have {uc.current_version()}).",
        "relay": "Tell the user once, in one sentence, and offer to update: if they agree, "
                 "call xl_app op update_now."}
    assert all("update" not in out for out, _ in results[1:])


def test_a_route_that_cannot_act_gets_no_dialog_only_the_key(monkeypatch, doc):
    monkeypatch.setattr(uf, "install_route", lambda: "pip")
    _cache()
    handler = Dialogs()
    (out, _), = _session([_read(doc)], handler)
    assert handler.seen == [] and out["update"]["relay"] == uf.RELAY_OFFER.format(
        tool=uc.UPDATE_TOOL)


def test_fallback_without_update_now_gives_the_manual_command(doc):
    from xlsx_mcp.settings import Settings

    _cache()
    (out, _), = _session([_read(doc)], settings=Settings(forbid=("update_now",)))
    assert out["update"]["relay"] == uf.RELAY_ONCE
    assert f"run uvx --refresh {uc.PACKAGE} --version once" in out["update"]["notice"]


def test_refusals_existing_keys_and_the_cap_get_no_notice(tmp_path):
    from xlsx_mcp.settings import Settings

    (plain, _), = _session([HELP])                  # no cache yet: no notice
    room = len(json.dumps(plain, ensure_ascii=False, separators=(",", ":")))
    _cache()
    results = _session([("xl_read", {"op": "get_workbook_metadata",
                                       "args": {"path": str(tmp_path / "none.xlsx")}}),
                        ("xl_read", {"op": "get_server_info"})])
    assert results[0][1] is True and "update" not in results[0][0]
    assert results[1][0]["update"] == uc.update_notice()      # its own key, not replaced
    assert "last_attempt" not in uf.read_state()
    (out, _), = _session([HELP], settings=Settings(max_output_chars=room + 40))
    assert "update" not in out and "last_attempt" not in uf.read_state()   # deferred


# ------------------------------------------------------ 4. switches, state


def test_notices_env_off_and_on(monkeypatch, doc):
    _cache()
    monkeypatch.setenv(uc.NOTICES_ENV, "off")
    handler = Dialogs()
    assert all("update" not in o for o, _ in _session([_read(doc)]))
    _session([_read(doc)], handler)
    assert handler.seen == []
    uf._record(notices_off=True)
    monkeypatch.setenv(uc.NOTICES_ENV, "on")
    _session([_read(doc)], handler)
    assert len(handler.seen) == 1 and uf.read_state()["notices_off"] is False


def test_the_update_notices_op_from_any_client(doc):
    _cache()
    (off, err), = _session([("xl_app", {"op": "update_notices", "args": {"action": "off"}})])
    assert not err and off["notices"] == "off" and off["decided_by"] == "setting"
    handler = Dialogs()
    _session([_read(doc), HELP], handler)
    assert handler.seen == []
    (st, _), = _session([("xl_app", {"op": "update_notices", "args": {"action": "status"}})])
    assert st["notices"] == "off" and st["newer_version"] == NEW and st["check"] == "on"
    (on, _), = _session([("xl_app", {"op": "update_notices", "args": {"action": "on"}})])
    assert on["notices"] == "on"
    _session([_read(doc)], handler)
    assert len(handler.seen) == 1
    assert uf.notices_path() == uc.cache_path().parent / "update-notices.json"


def test_corrupt_state_and_a_future_attempt_mean_nothing_recorded(doc):
    _cache()
    uf.notices_path().parent.mkdir(parents=True, exist_ok=True)
    uf.notices_path().write_text("{not json", encoding="utf-8")
    assert uf.due() is not None
    uf.notices_path().write_text(json.dumps({"last_attempt": (
        datetime.now(timezone.utc) + timedelta(days=30)).isoformat()}), encoding="utf-8")
    assert uf.due() is not None


def test_two_processes_claim_one_notice(tmp_path):
    code = ("import sys; sys.path[:0] = {paths!r}; import conftest; "
            "from xlsx_mcp.core import update_flow as uf; print(uf.claim())")
    paths = [str(ROOT / "src"), str(ROOT / "tests"), str(ROOT / "tests" / "unit")]
    env = {**os.environ, uc.CACHE_DIR_ENV: str(tmp_path / "shared")}
    procs = [subprocess.Popen([sys.executable, "-X", "utf8", "-c", code.format(paths=paths)],
                              env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                              stdin=subprocess.DEVNULL,
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
             for _ in range(2)]
    outs = [p.communicate(timeout=60)[0].strip() for p in procs]
    assert sorted(outs) == ["False", "True"], outs
    json.loads((tmp_path / "shared" / "update-notices.json").read_text(encoding="utf-8"))


def test_an_unwritable_state_folder_gives_one_attempt_per_process(monkeypatch):
    import contextlib

    @contextlib.contextmanager
    def no_lock(wait_s=1.0):
        yield None

    monkeypatch.setattr(uf, "_state_lock", no_lock)
    monkeypatch.setattr(uc, "atomic_write_json", lambda *a, **k: False)
    assert uf.claim() is True
    assert uf.claim() is False


def test_versions_rc_builds_get_the_final_and_stable_builds_never_an_rc(monkeypatch):
    monkeypatch.setattr(uc, "current_version", lambda: "2.0.0rc1")
    _cache("2.0.0")
    assert uc.versions() == ("2.0.0rc1", "2.0.0")
    monkeypatch.setattr(uc, "current_version", lambda: "2.0.0")
    _cache("2.0.1rc1")
    assert uc.versions() is None
    _cache("2.0.1; rm -rf /")
    assert uc.versions() is None


# --------------------------------------------------- 5. Business and the wire


def _business(monkeypatch):
    monkeypatch.setattr(uc, "resolve_backend", lambda: SimpleNamespace(name="business"))


def test_business_is_byte_identical(monkeypatch, doc):
    _cache()
    _business(monkeypatch)
    payload = {"ok": True, "x": 1}
    assert asyncio.run(uf.after_call(payload, False, read_only=True, cap=80000,
                                     can_update=True)) is payload
    handler = Dialogs()
    results = _session([_read(doc), HELP], handler)
    assert handler.seen == [] and uc._REFRESH_THREAD is None
    assert all("update" not in out for out, _ in results)
    (out, err), = _session([("xl_app", {"op": "update_now"})])
    assert err and out["error"]["code"] == "NOT_FOUND"


def test_business_tool_list_keeps_the_released_figures(monkeypatch):
    from xlsx_mcp.settings import Settings
    from xlsx_mcp.surface import plan as plan_mod

    _business(monkeypatch)
    plan = plan_mod.build(Settings())
    app = plan.tools["xl_app"]
    assert not set(uf.OPS) & set(app.ops) and app.annotations["openWorldHint"] is False
    assert plan.op_count() == 67
    # Byte-identical to the tool list before the update ops (222e2c4):
    # 5,377 tokens in process, 5,433 over stdio (scripts/measure_surface.py).
    assert server._approx_active_tokens(plan) == 5377
    tools = {n: (t.description, t.input_schema, t.annotations) for n, t in plan.tools.items()}
    assert hashlib.sha256(json.dumps(tools, sort_keys=True).encode()).hexdigest() == (
        "6759172fdec8df27beae02df31ae9164b4eb02e4979e67f2b151cd50631c4204")


def test_instructions_stay_under_the_cap():
    mcp = server.create_server()
    assert len(mcp.instructions or "") <= 2048


def test_the_ops_live_in_xl_app_and_xl_read_stays_read_only():
    from xlsx_mcp.settings import Settings
    from xlsx_mcp.surface import plan as plan_mod

    plan = plan_mod.build(Settings())
    assert plan.op_tool["update_now"] == plan.op_tool["update_notices"] == uc.UPDATE_TOOL
    assert plan.tools["xl_read"].read_only and not plan.tools["xl_app"].read_only
    op_text = plan.tools["xl_app"].input_schema["properties"]["op"]["description"]
    assert "update_now(): Install the newer version the update notice named; call only " \
           "when the user asks or agrees." in op_text


# ------------------------------------------------------------ 6. update_now


class FakeChild:
    calls: list = []
    printed = b"kitchensink4xl 999.0.1\n"
    returncode = 0
    delay = 0.0

    def __init__(self, argv, **kwargs):
        FakeChild.calls.append((argv, kwargs))
        self.killed = False

    def communicate(self, timeout=None):
        if FakeChild.delay:
            time.sleep(FakeChild.delay)
        return FakeChild.printed, b""

    def kill(self):
        self.killed = True


@pytest.fixture
def uvx(monkeypatch, tmp_path):
    uv = tmp_path / ("uv.exe" if os.name == "nt" else "uv")
    uv.write_bytes(b"")
    monkeypatch.setenv("UV", str(uv))
    monkeypatch.setattr(uf, "install_route", lambda: "uvx")
    FakeChild.calls, FakeChild.printed, FakeChild.returncode, FakeChild.delay = \
        [], b"kitchensink4xl 999.0.1\n", 0, 0.0
    monkeypatch.setattr(uf, "_spawn", FakeChild)
    return str(uv)


def test_update_now_asks_first_and_not_now_runs_nothing(uvx):
    _cache()
    handler = Dialogs({"choice": uf.NOT_NOW})
    (out, err), = _session([("xl_app", {"op": "update_now"})], handler)
    assert not err and out["status"] == "not_now" and FakeChild.calls == []
    assert len(handler.seen) == 1


def test_update_now_asks_its_own_dialog_when_the_cache_did_not_know_the_version(
        uvx, monkeypatch):
    """The cache is due and names the installed version; the refresh finds a newer
    one. A client that can ask and answers Not now gets one dialog and no child."""
    _cache(uc.current_version(), age_days=8)
    monkeypatch.setattr(uc, "_fetch", lambda url=uc.PYPI_URL, timeout=uc.TIMEOUT_SECONDS:
                        _payload(uc.current_version(), NEW))
    handler = Dialogs({"choice": uf.NOT_NOW})
    (out, err), = _session([("xl_app", {"op": "update_now"})], handler)
    assert not err and out["status"] == "not_now", out
    assert out["new_version"] == NEW
    assert len(handler.seen) == 1 and FakeChild.calls == []


def test_update_now_acts_only_on_the_version_its_dialog_step_decided(uvx, monkeypatch):
    """A slow background check still running when the dialog step looks finds no
    newer version, so no dialog is asked. If that check finishes before update_now
    acts, update_now must not install the version no dialog asked about."""
    seen = []

    def newer(refresh=False):
        seen.append(refresh)
        return None if len(seen) == 1 else (uc.current_version(), NEW)

    monkeypatch.setattr(uf, "_newer_for_update_now", newer)
    handler = Dialogs({"choice": uf.NOT_NOW})
    (out, err), = _session([("xl_app", {"op": "update_now"})], handler)
    assert FakeChild.calls == [] and handler.seen == []
    assert "NOT_FOUND" in str(out) + str(err), (out, err)


def test_update_now_runs_the_fixed_argv_and_reports_the_reload_line(uvx):
    _cache()
    handler = Dialogs({"choice": uf.UPDATE_NOW})
    (out, err), = _session([("xl_app", {"op": "update_now"})], handler)
    assert not err and out["status"] == "updated", out
    (argv, kwargs), = FakeChild.calls
    assert argv == [uvx, "tool", "run", "--refresh-package", uc.PACKAGE,
                    "--refresh-package", uc.CORE_PACKAGE, uc.PACKAGE, "--version"]
    assert kwargs["shell"] is False and kwargs["stdin"] is subprocess.DEVNULL
    assert kwargs["creationflags"] == getattr(subprocess, "CREATE_NO_WINDOW", 0)
    assert "restart your AI app" in out["message"]          # fastmcp's client name
    assert f"such as {uc.PACKAGE}==X, change it to {uc.PACKAGE}" in out["message"]
    assert uf.read_state()["installed_version"] == NEW
    assert uf.due() is None                                   # installed, awaiting restart


@pytest.mark.parametrize("printed,code", [(b"kitchensink4xl 2.0.0\n", 0), (b"", 0),
                                          (b"kitchensink4xl 999.0.1\n", 2)])
def test_uvx_succeeds_only_when_the_new_version_runs(uvx, printed, code):
    _cache()
    FakeChild.printed, FakeChild.returncode = printed, code
    (out, _), = _session([("xl_app", {"op": "update_now"})])
    assert out["status"] == "failed" and "Nothing changed" in out["message"]


def test_update_now_never_blocks_long_and_a_second_call_attaches(uvx, monkeypatch):
    _cache()
    monkeypatch.setattr(uf, "UPDATE_NOW_WAIT_S", 0.2)
    FakeChild.delay = 1.0
    results = _session([("xl_app", {"op": "update_now"}), ("xl_app", {"op": "update_now"})])
    assert [o["status"] for o, _ in results] == ["updating", "updating"]
    assert len(FakeChild.calls) == 1
    uf._JOB.done.wait(5)
    (out, _), = _session([HELP])
    assert out["update"]["status"] == "updated"               # the next result carries it
    (out, _), = _session([HELP])
    assert "update" not in out                                # told once


def test_the_dialog_update_now_installs_in_the_background(uvx, doc):
    _cache()
    FakeChild.delay = 0.5
    handler = Dialogs({"choice": uf.UPDATE_NOW})
    (first, _), = _session([_read(doc)], handler)
    assert first["update"]["status"] == "updating in the background"
    uf._JOB.done.wait(5)
    (later, _), = _session([HELP])
    assert later["update"]["status"] == "updated"


def test_a_hard_timeout_kills_the_child(uvx, monkeypatch):
    killed = []

    class Hangs(FakeChild):
        def communicate(self, timeout=None):
            if not killed:
                raise subprocess.TimeoutExpired("uv", timeout)
            return b"", b""

        def kill(self):
            killed.append(True)

    monkeypatch.setattr(uf, "_spawn", Hangs)
    _cache()
    (out, _), = _session([("xl_app", {"op": "update_now"})])
    assert killed and out["status"] == "failed" and "longer than 120 seconds" in out["message"]


def test_update_now_refuses_without_a_newer_version_and_takes_no_arguments():
    _cache(uc.current_version())
    (out, err), = _session([("xl_app", {"op": "update_now"})])
    assert err and out["error"]["code"] == "NOT_FOUND"
    (out, err), = _session([("xl_app", {"op": "update_now", "args": {"version": "9"}})])
    assert err and out["error"]["code"] == "BAD_PARAMS"
    token = uf._CONSENT.set("unasked")
    try:
        _cache()
        with pytest.raises(uf.UpdateRefused):
            uf.update_now()
    finally:
        uf._CONSENT.reset(token)


@pytest.mark.parametrize("route,needle", [("pip", f"pip install -U {uc.PACKAGE}"),
                                          ("uv-tool", f"uv tool upgrade {uc.PACKAGE}")])
def test_pip_and_uv_tool_return_the_command_and_run_nothing(route, needle, monkeypatch):
    monkeypatch.setattr(uf, "install_route", lambda: route)
    monkeypatch.setattr(uf, "_spawn", _landmine)
    _cache()
    (out, _), = _session([("xl_app", {"op": "update_now"})])
    assert out["status"] == "command" and needle in out["message"]


def test_route_detection(monkeypatch, tmp_path):
    route = _REAL_ROUTE
    archive = tmp_path / "cache" / "archive-v0" / "abc"
    archive.mkdir(parents=True)
    uv = tmp_path / ("uv.exe" if os.name == "nt" else "uv")
    uv.write_bytes(b"")
    monkeypatch.setattr(sys, "prefix", str(archive))
    assert route() == "pip"                       # no UV variable
    monkeypatch.setenv("UV", "uv")
    assert route() == "pip"                       # relative: never a PATH lookup
    monkeypatch.setenv("UV", str(tmp_path / "python.exe"))
    assert route() == "pip"                       # not named uv
    monkeypatch.setenv("UV", str(tmp_path / "missing" / uv.name))
    assert route() == "pip"                       # does not exist
    monkeypatch.setenv("UV", str(uv))
    assert route() == "uvx"
    (archive / "uv-receipt.toml").write_text("", encoding="utf-8")
    assert route() == "uv-tool"                   # a uv tool install
    monkeypatch.setenv(uc.INSTALL_ROUTE_ENV, "mcpb")
    assert route() == "mcpb"                      # the bundle manifest setting only


def test_reload_lines():
    assert uf.reload_line("claude-code") == \
        f"run /mcp, select the {uc.PRODUCT} server, choose Reconnect"
    assert uf.reload_line("cursor-vscode") == "turn the server off and on in Settings > MCP"
    assert uf.reload_line("Visual Studio Code") == "run MCP: Restart Server"
    assert uf.reload_line("codex-mcp-client") == "restart the app"
    assert uf.reload_line("evil\nname") == "restart your AI app"


# --------------------------------------------------------- 7. Desktop bundle


def _bundle(version=NEW, **changes) -> bytes:
    manifest = json.loads((ROOT / "bundle" / "manifest.json").read_text(encoding="utf-8"))
    manifest["version"] = version
    manifest["server"]["mcp_config"]["args"] = [f"{uc.PACKAGE}{uc.BUNDLE_EXTRAS}=={version}"]
    for key, value in changes.items():
        if key == "env":
            manifest["server"]["mcp_config"]["env"] = value
        elif key == "command":
            manifest["server"]["mcp_config"]["command"] = value
        else:
            manifest[key] = value
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("manifest.json", json.dumps(manifest))
        z.write(ROOT / "bundle" / "icon.png", "icon.png")
        if changes.get("extra"):
            z.writestr("run.ps1", "x")
    return buf.getvalue()


def test_the_bundle_check():
    assert uf.check_bundle(_bundle(), NEW)["name"] == uc.PACKAGE
    for bad in (_bundle(extra=True), _bundle(name="other"), _bundle("1.0.0"),
                _bundle(command="powershell"), _bundle(env={"PATH": "x"}),
                _bundle(env={uc.INSTALL_ROUTE_ENV: "pip"}),
                _bundle(env={**uc.BUNDLE_USER_ENV, "KS4XL_VERIFY_COM": "1"}), b"not a zip", b"x" * 1_000_001):
        with pytest.raises(uf.BundleRefused):
            uf.check_bundle(bad, NEW)
    assert uf.bundle_url("2.0.1") == ("https://github.com/KitchenSink4AI/KitchenSink4XL/"
                                      "releases/download/v2.0.1/kitchensink4xl-2.0.1.mcpb")


@pytest.mark.parametrize("tamper", [None, "hash", "bundle"])
def test_the_desktop_route_checks_before_it_opens(tamper, monkeypatch, tmp_path):
    data = _bundle(extra=True) if tamper == "bundle" else _bundle()
    digest = hashlib.sha256(b"other" if tamper == "hash" else data).hexdigest()
    fetched, opened = [], []

    def download(url, limit):
        fetched.append(url)
        if url.endswith(".sha256"):
            return f"{digest}  {uc.PACKAGE}-{NEW}.mcpb\n".encode()
        return data

    monkeypatch.setattr(uf, "_download", download)
    monkeypatch.setattr(uf, "_open_bundle", opened.append)
    monkeypatch.setattr(uf, "install_route", lambda: "mcpb")
    monkeypatch.setattr(uf.sys, "platform", "win32")
    _cache()
    (out, _), = _session([("xl_app", {"op": "update_now"})])
    assert fetched[0] == uf.bundle_url(NEW) + ".sha256"
    if tamper is None:
        assert out["status"] == "opened"
        assert opened == [uc.cache_path().parent / f"{uc.PACKAGE}-update.mcpb"]
    else:
        assert out["status"] == "failed" and opened == []
        assert "Nothing was installed" in out["message"]


def test_redirects_leave_github_never():
    handler = uf._GitHubOnly()
    with pytest.raises(Exception):
        handler.redirect_request(None, None, 302, "", {}, "https://example.com/x.mcpb")
    with pytest.raises(Exception):
        handler.redirect_request(None, None, 302, "", {}, "http://github.com/x.mcpb")


def test_the_release_files_pass_the_same_check(tmp_path):
    sys.path.insert(0, str(ROOT / "scripts"))
    import bundle_release

    packed = tmp_path / f"{uc.PACKAGE}.mcpb"
    packed.write_bytes(_bundle(uc.current_version()))
    bundle, digest = bundle_release.release_files(packed)
    assert bundle.name == f"{uc.PACKAGE}-{uc.current_version()}.mcpb"
    text = digest.read_text(encoding="ascii")
    assert uf.expected_sha256(text, bundle.name) == hashlib.sha256(bundle.read_bytes()).hexdigest()
    manifest = json.loads((ROOT / "bundle" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["server"]["mcp_config"]["env"] == {
        **uc.BUNDLE_USER_ENV, uc.INSTALL_ROUTE_ENV: "mcpb"}
    assert manifest["server"]["mcp_config"]["args"] == [f"{uc.PACKAGE}{uc.BUNDLE_EXTRAS}==2.0.0"]


# ------------------------------------------------------ 8. stdio round trips


def _stdio(tmp_path, declare: bool, answer: str = "Not now") -> dict:
    """The Community launcher over real stdio with a raw client that may
    declare elicitation and answers each dialog with `answer`."""
    import family_harness as fh

    state = tmp_path / "stdio-state"
    # The bundle route acts, so a client with the dialog is asked; the answer
    # is Not now, so nothing is downloaded.
    env = fh._env({uc.CACHE_DIR_ENV: str(state), uc.OFF_ENV: "on",
                   uc.INSTALL_ROUTE_ENV: "mcpb"})
    stamp = datetime.now(timezone.utc).isoformat()
    state.mkdir(parents=True, exist_ok=True)
    (state / "update-check.json").write_text(json.dumps({
        "last_check": stamp, "latest_version": NEW, "ok": True, "last_success": stamp,
        "disclosed": True}), encoding="utf-8")
    proc = subprocess.Popen(fh.launch_cmd(), stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, env=env,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    killer = threading.Timer(90, proc.kill)
    killer.start()
    seen: dict = {"dialogs": [], "results": []}

    def send(msg):
        proc.stdin.write((json.dumps(msg) + "\n").encode())
        proc.stdin.flush()

    def recv(want):
        while True:
            msg = json.loads(proc.stdout.readline())
            if msg.get("method") == "elicitation/create":
                seen["dialogs"].append(msg["params"])
                if answer == "-32601":
                    send({"jsonrpc": "2.0", "id": msg["id"],
                          "error": {"code": -32601, "message": "Method not found"}})
                else:
                    send({"jsonrpc": "2.0", "id": msg["id"],
                          "result": {"action": "accept", "content": {"choice": answer}}})
                continue
            if msg.get("id") == want and "method" not in msg:
                return msg

    try:
        caps = {"elicitation": {}} if declare else {}
        send({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {
            "protocolVersion": min(uf.DIALOG_PROTOCOLS), "capabilities": caps,
            "clientInfo": {"name": "claude-code", "version": "0"}}})
        recv(1)
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        for rid in (2, 3):
            send({"jsonrpc": "2.0", "id": rid, "method": "tools/call",
                  "params": {"name": "xl_help", "arguments": {}}})
            seen["results"].append(json.loads(recv(rid)["result"]["content"][0]["text"]))
    finally:
        killer.cancel()
        proc.stdin.close()
        try:
            proc.wait(timeout=20)
        except subprocess.TimeoutExpired:
            proc.kill()
        seen["stderr"] = proc.stderr.read().decode("utf-8", "replace")
    seen["state"] = json.loads((state / "update-notices.json").read_text(encoding="utf-8"))
    return seen


def test_stdio_a_client_with_the_dialog_answers_not_now(tmp_path):
    seen = _stdio(tmp_path, declare=True)
    assert len(seen["dialogs"]) == 1
    assert seen["dialogs"][0]["requestedSchema"]["properties"]["choice"]["enum"] == \
        ["Update now", "Not now"]
    assert all("update" not in r for r in seen["results"])
    assert "last_attempt" in seen["state"]
    assert f"{uc.PRODUCT} {NEW} is available (installed" in seen["stderr"]


def test_stdio_a_method_not_found_answer_falls_back(tmp_path):
    """A client that answers -32601 (method not found)."""
    seen = _stdio(tmp_path, declare=True, answer="-32601")
    assert len(seen["dialogs"]) == 1
    assert seen["results"][0]["update"]["relay"].startswith("Tell the user once")
    assert "update" not in seen["results"][1]
    assert seen["state"]["last_dialog"] == "unavailable:error"


def test_stdio_a_client_without_the_dialog_gets_the_key_once(tmp_path):
    seen = _stdio(tmp_path, declare=False)
    assert seen["dialogs"] == []
    assert seen["results"][0]["update"]["relay"].startswith("Tell the user once")
    assert "update" not in seen["results"][1]


# ------------------------------------------- 9. the launcher and a real uv


def test_the_launcher_prints_its_version():
    import family_harness as fh

    cmd = fh.launch_cmd()
    cmd[-1] = cmd[-1].replace("main()", "sys.argv[1:] = ['--version']; main()")
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=60, env=fh._env(None),
                         stdin=subprocess.DEVNULL,
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    assert res.returncode == 0 and res.stdout.strip() == f"{uc.PACKAGE} {uc.current_version()}"


def _wheel(folder: Path, version: str) -> None:
    """A dummy kitchensink4xl wheel (with its [com] extra) whose console
    script answers --version."""
    name = f"{uc.PACKAGE}-{version}"
    files = {
        f"{uc.PACKAGE}_probe/__init__.py":
            f"def main():\n    print('{uc.PACKAGE} {version}')\n",
        f"{name}.dist-info/METADATA":
            f"Metadata-Version: 2.1\nName: {uc.PACKAGE}\nVersion: {version}\n"
            "Provides-Extra: com\n",
        f"{name}.dist-info/WHEEL":
            "Wheel-Version: 1.0\nGenerator: probe\nRoot-Is-Purelib: true\nTag: py3-none-any\n",
        f"{name}.dist-info/entry_points.txt":
            f"[console_scripts]\n{uc.PACKAGE} = {uc.PACKAGE}_probe:main\n",
    }
    record = "".join(f"{p},,\n" for p in files) + f"{name}.dist-info/RECORD,,\n"
    files[f"{name}.dist-info/RECORD"] = record
    with zipfile.ZipFile(folder / f"{name}-py3-none-any.whl", "w") as z:
        for path, text in files.items():
            z.writestr(path, text)


def _index(root: Path) -> None:
    links = "".join(
        f'<a href="/files/{w.name}#sha256={hashlib.sha256(w.read_bytes()).hexdigest()}">'
        f"{w.name}</a><br/>" for w in sorted((root / "files").glob("*.whl")))
    page = root / "simple" / uc.PACKAGE
    page.mkdir(parents=True, exist_ok=True)
    (page / "index.html").write_text(f"<html><body>{links}</body></html>", encoding="utf-8")


@pytest.mark.skipif(shutil.which("uv") is None, reason="uv is not installed")
def test_uvx_probe_a_local_index_with_two_versions(tmp_path, monkeypatch):
    """After the update's own command, plain `uvx PACKAGE` runs
    the new version; the manual hint's command does the same. Local index
    on 127.0.0.1 only."""
    root = tmp_path / "idx"
    (root / "files").mkdir(parents=True)
    _wheel(root / "files", "2.0.0")
    _index(root)

    class Quiet(http.server.SimpleHTTPRequestHandler):
        def __init__(self, *a, **k):
            super().__init__(*a, directory=str(root), **k)

        def end_headers(self):
            self.send_header("Cache-Control", "max-age=600, public")
            super().end_headers()

        def log_message(self, *a):
            pass

    httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), Quiet)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    uv = shutil.which("uv")
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    def env(cache):
        e = {k: v for k, v in os.environ.items() if not k.startswith("UV_")}
        e.update(UV_CACHE_DIR=str(tmp_path / cache), UV_TOOL_DIR=str(tmp_path / "tools"),
                 UV_DEFAULT_INDEX=f"http://127.0.0.1:{httpd.server_port}/simple",
                 UV_PYTHON=sys.executable, UV_PYTHON_DOWNLOADS="never", UV_NO_CONFIG="1",
                 UV=uv)
        return e

    def uvx(cache, *args):
        res = subprocess.run([uv, "tool", "run", *args], capture_output=True, text=True,
                             timeout=120, env=env(cache), stdin=subprocess.DEVNULL,
                             creationflags=flags)
        assert res.returncode == 0, res.stderr[-600:]
        return res.stdout.strip()

    try:
        for cache in ("c1", "c2"):
            assert uvx(cache, uc.PACKAGE, "--version") == f"{uc.PACKAGE} 2.0.0"
        _wheel(root / "files", "2.0.1")
        _index(root)
        assert uvx("c1", uc.PACKAGE, "--version") == f"{uc.PACKAGE} 2.0.0"   # stays cached
        # the update's own child, for real
        monkeypatch.setattr(os, "environ", env("c1"))
        monkeypatch.setattr(uf, "_spawn", _REAL_SPAWN)
        job = uf._Job("2.0.1", "uvx", "claude-code")
        uf._run_uvx(job)
        assert job.status == "updated", job.message
        assert uvx("c1", uc.PACKAGE, "--version") == f"{uc.PACKAGE} 2.0.1"
        # the Desktop bundle and the README launch with the [com] extra: the
        # update's refresh must move that spelling too
        assert uvx("c1", f"{uc.PACKAGE}[com]", "--version") == f"{uc.PACKAGE} 2.0.1"
        # the manual hint: uvx --refresh PACKAGE --version
        assert uvx("c2", "--refresh", uc.PACKAGE, "--version") == f"{uc.PACKAGE} 2.0.1"
        assert uvx("c2", uc.PACKAGE, "--version") == f"{uc.PACKAGE} 2.0.1"
        # a pinned launch command stays where it is
        assert uvx("c1", f"{uc.PACKAGE}==2.0.0", "--version") == f"{uc.PACKAGE} 2.0.0"
    finally:
        httpd.shutdown()
