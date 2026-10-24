"""The pack start-up settings are gone, and an upgrade keeps working.

Released 1.x servers read KS4XL_MODE (the startup pack list) and
KS4XL_PACK_POLICY (auto/locked). Every operation is now listed from the
first message, so neither does anything. Users who upgrade keep their old
launch configuration, so the old DEFAULTS are accepted and ignored with one
stderr line, and only a value that limited the tools in 1.x stops the start
(a lock that silently stopped locking would fail open). Excel never had the
.mcpb all_tools / lock_tools tick boxes, so there is no bundle toggle and
those names are not read. What this file pins:

1. THE BUNDLE maps no legacy variable and offers no tool-loading option.
2. AN OLD DEFAULT IS IGNORED: MODE lite / full / everything and PACK_POLICY
   auto, any case, any surrounding whitespace; a blank value is the same as
   unset and is not listed. The ignored names, never their values, are in
   Settings.legacy_set, the one stderr line, and get_server_info.
3. A VALUE THAT LIMITED THE TOOLS STOPS THE SERVER, alone and combined with
   an accepted MODE, in either order (copy slot LEGACY_SETTING_REFUSAL).
4. Present-day settings (READ_ONLY, FORBID) keep their validation and
   enforcement when an old default is present.
5. Over real stdio, a start with MODE=lite lists the full tool set.
"""

from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest

import family_harness as fh
from xlsx_mcp import server
from xlsx_mcp.settings import LEGACY_SETTINGS, Settings, SettingsError
from xlsx_mcp.surface import copy, plan

MANIFEST = Path(__file__).resolve().parents[2] / "bundle" / "manifest.json"
MODE, POLICY = "KS4XL_MODE", "KS4XL_PACK_POLICY"
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
ACCEPTED = [(MODE, "lite"), (MODE, "full"), (MODE, "everything"), (MODE, "LITE"),
            (MODE, " Full "), (MODE, "\tEVERYTHING  "), (POLICY, "auto"), (POLICY, " AUTO ")]
REFUSED = [(MODE, "design,com"), (MODE, "lite,design"), (MODE, "locked"), (MODE, "fulll"),
           (MODE, "true"), (MODE, "1"), (MODE, "auto"),
           (POLICY, "locked"), (POLICY, " LOCKED "), (POLICY, "lockedd"), (POLICY, "full"),
           (POLICY, "true"), (POLICY, "0")]


def test_the_bundle_maps_no_legacy_setting():
    m = json.loads(MANIFEST.read_text("utf-8"))
    for key in ("all_tools", "lock_tools", "mode", "pack_policy"):
        assert key not in m.get("user_config", {}), key
    env = m["server"]["mcp_config"].get("env", {})
    for name in LEGACY_SETTINGS:
        assert name not in env, name
    # What the bundle does keep.
    assert "verify_with_excel" in m.get("user_config", {})
    assert "allowed_root" in json.dumps(m.get("user_config", {}))


def test_no_environment_ignores_nothing():
    s = Settings.from_env({})
    assert s.legacy_set == ()
    assert plan.build(s).op_count() == 69


@pytest.mark.parametrize("name,value", ACCEPTED)
def test_an_old_default_is_accepted_and_named(name, value):
    s = Settings.from_env({name: value})
    assert s.legacy_set == (name,)
    assert s.read_only is False and s.forbid == ()
    p = plan.build(s)
    assert p.op_count() == 69 and len(p.tools) == 7


def test_both_old_defaults_are_named_in_order():
    s = Settings.from_env({POLICY: "auto", MODE: "lite"})
    assert s.legacy_set == (MODE, POLICY)


@pytest.mark.parametrize("blank", ["", " ", "  \t "])
@pytest.mark.parametrize("name", LEGACY_SETTINGS)
def test_a_blank_retired_variable_is_accepted_and_not_listed(name, blank):
    s = Settings.from_env({name: blank})
    assert s.legacy_set == ()
    p = plan.build(s)
    assert p.op_count() == 69 and len(p.tools) == 7


@pytest.mark.parametrize("name,value", REFUSED)
def test_a_value_that_limited_the_tools_refuses_to_start(name, value):
    alone = {name: value}
    beside = {name: value, (POLICY if name == MODE else MODE): ("auto" if name == MODE else "lite")}
    for env in (alone, beside):
        with pytest.raises(SettingsError) as ei:
            Settings.from_env(env)
        msg = str(ei.value)
        assert msg.startswith(name + " is set to a value that limited the tools"), msg
        assert "KS4XL_READ_ONLY" in msg and "KS4XL_FORBID" in msg
        assert "{name}" not in msg and chr(10) not in msg
        assert chr(0x2014) not in msg


def test_a_refusal_names_the_variable_that_is_refused_not_the_accepted_one():
    with pytest.raises(SettingsError) as ei:
        Settings.from_env({MODE: "lite", POLICY: "locked"})
    assert str(ei.value).startswith(POLICY + " ")
    with pytest.raises(SettingsError) as ei:
        Settings.from_env({MODE: "design,com", POLICY: "auto"})
    assert str(ei.value).startswith(MODE + " ")


@pytest.mark.parametrize("name,value", [(MODE, "lite"), (POLICY, "auto")])
def test_present_day_restrictions_still_apply_beside_an_old_default(name, value):
    ro = plan.build(Settings.from_env({name: value, "KS4XL_READ_ONLY": "1"}))
    plain = plan.build(Settings.from_env({"KS4XL_READ_ONLY": "1"}))
    assert ro.hidden_by_read_only and ro.hidden_by_read_only == plain.hidden_by_read_only
    assert list(ro.tools) == list(plain.tools) and ro.op_count() < 69
    fb = plan.build(Settings.from_env({name: value, "KS4XL_FORBID": "set_cell,xl_app"}))
    assert "set_cell" in fb.forbidden_ops and "xl_app" in fb.forbidden_tools
    with pytest.raises(SettingsError):
        Settings.from_env({name: value, "KS4XL_READ_ONLY": "maybe"})
    with pytest.raises(SettingsError):
        fh.start(Settings.from_env({name: value, "KS4XL_FORBID": "no_such_thing"}))


def test_forbid_and_read_only_are_enforced_in_the_call_beside_mode_lite():
    h = fh.start(Settings.from_env({MODE: "lite", "KS4XL_READ_ONLY": "1"}))
    assert [t.name for t in h.list_tools()] == ["xl_read", "xl_help"]
    out, err, _ = h.call("xl_read", {"op": "set_cell"})
    assert err and out["error"]["code"] == "NOT_AVAILABLE"
    info, _, _ = h.call("xl_read", {"op": "get_server_info"})
    assert info["surface"]["legacy_settings_ignored"] == [MODE]
    h = fh.start(Settings.from_env({MODE: "lite", "KS4XL_FORBID": "audit_styles"}))
    out, err, _ = h.call("xl_read", {"op": "audit_styles"})
    assert err and out["error"]["code"] == "NOT_AVAILABLE"


def test_status_shows_the_ignored_names_and_nothing_else():
    h = fh.start(Settings.from_env({MODE: " LITE ", POLICY: "auto"}))
    info, err, _ = h.call("xl_read", {"op": "get_server_info"})
    assert not err
    assert info["surface"]["legacy_settings_ignored"] == [MODE, POLICY]
    assert "LITE" not in json.dumps(info)
    info, _, _ = fh.start(Settings.from_env({})).call("xl_read", {"op": "get_server_info"})
    assert info["surface"]["legacy_settings_ignored"] == []
    info, _, _ = fh.start(Settings.from_env({MODE: ""})).call(
        "xl_read", {"op": "get_server_info"})
    assert info["surface"]["legacy_settings_ignored"] == []


def test_the_notice_is_one_tagged_stderr_line_and_never_echoes_a_value(capsys):
    server._legacy_notice(Settings.from_env({MODE: "  EVERYTHING ", POLICY: "AUTO"}))
    cap = capsys.readouterr()
    assert cap.out == ""
    assert cap.err.count(chr(10)) == 1 and cap.err.startswith("[kitchensink4xl] ")
    assert ("Ignored old settings: KS4XL_MODE, KS4XL_PACK_POLICY. This version lists every "
            "tool from the start, and the server started normally. To limit the tools, use "
            "KS4XL_READ_ONLY=1 or KS4XL_FORBID=<tool or op names>. Remove these settings to "
            "stop this notice.") in cap.err
    assert "EVERYTHING" not in cap.err and "AUTO" not in cap.err
    assert "{names}" not in cap.err and chr(0x2014) not in cap.err


@pytest.mark.parametrize("env", [{}, {MODE: ""}, {POLICY: "  "}])
def test_no_notice_when_nothing_was_ignored(env, capsys):
    server._legacy_notice(Settings.from_env(env))
    cap = capsys.readouterr()
    assert cap.out == "" and cap.err == ""


def test_the_refusal_is_one_line_from_main(capsys, monkeypatch):
    monkeypatch.setenv(POLICY, "locked")
    monkeypatch.setenv(MODE, "lite")
    with pytest.raises(SystemExit) as ei:
        server.main()
    assert ei.value.code == 2
    cap = capsys.readouterr()
    assert cap.out == ""
    assert cap.err.count(chr(10)) == 1 and cap.err.startswith("[kitchensink4xl] " + POLICY)
    assert "did not start" in cap.err and "Ignored old settings" not in cap.err


def test_a_refused_start_writes_no_notice_even_when_mode_is_accepted():
    code, err = fh.stdio_start_only({MODE: "lite", POLICY: "locked"})
    assert code == 2 and "Ignored old settings" not in err
    assert err.strip().count(chr(10)) == 0 and err.startswith("[kitchensink4xl] " + POLICY)


def test_a_real_start_with_old_defaults_writes_one_stderr_line_and_no_stdout():
    res = subprocess.run(fh.launch_cmd(), stdin=subprocess.DEVNULL, capture_output=True,
                         env=fh._env({MODE: " LITE ", POLICY: "auto"}), timeout=60,
                         creationflags=NO_WINDOW)
    err = res.stderr.decode("utf-8", "replace")
    assert res.returncode == 0, err
    assert res.stdout == b""
    notice = [ln for ln in err.splitlines() if "Ignored old settings" in ln]
    assert len(notice) == 1, err
    assert notice[0].startswith("[kitchensink4xl] Ignored old settings: KS4XL_MODE, "
                                "KS4XL_PACK_POLICY.")
    assert "LITE" not in err


def test_over_real_stdio_mode_lite_lists_the_full_tool_set():
    res = fh.stdio_session([("xl_read", {"op": "get_server_info"})],
                           env={MODE: "lite", POLICY: "auto"})
    assert [t["name"] for t in res["tools"]] == [
        "xl_read", "xl_edit", "xl_format", "xl_structure", "xl_workbook", "xl_app", "xl_help"]
    info = json.loads(res["results"][0]["text"])
    assert info["surface"]["tools"] == 7 and info["surface"]["ops"] == 69
    assert info["surface"]["legacy_settings_ignored"] == [MODE, POLICY]
    notice = [ln for ln in res["stderr"].splitlines() if "Ignored old settings" in ln]
    assert len(notice) == 1, res["stderr"]
    # the full op list, not a lite subset
    ops = {op for t in res["tools"] for op in
           t["inputSchema"]["properties"]["op"].get("enum", [])}
    assert len(ops) == 69 and "set_cell" in ops and "com_goal_seek" in ops


def test_the_copy_slots_carry_the_final_wording():
    assert copy.LEGACY_SETTING_NOTICE.count("{names}") == 1
    assert copy.LEGACY_SETTING_REFUSAL.count("{name}") == 2
    for text in (copy.LEGACY_SETTING_NOTICE, copy.LEGACY_SETTING_REFUSAL):
        assert chr(0x2014) not in text and "KS4XL_READ_ONLY=1" in text
