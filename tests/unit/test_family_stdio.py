"""The family surface over REAL stdio through the Community launcher: the
fixed list, no listChanged, the compact single emission on the wire,
launch settings enforced inside the call, legacy settings, and the setup
refusals. The same design as the Word editor's family-surface tests."""

from __future__ import annotations

import json

import pytest

import family_harness as fh

EXPECTED = ["xl_read", "xl_edit", "xl_format", "xl_structure", "xl_workbook", "xl_app",
            "xl_help"]


def _body(r):
    return json.loads(r["text"])


def test_wire_fixed_list_every_tool_once_and_no_list_changed(tmp_path):
    book = str(tmp_path / "wire.xlsx")
    copy_to = str(tmp_path / "wire-copy.xlsx")
    calls = [
        ("xl_help", {}),
        ("xl_read", {"op": "get_server_info"}),
        ("xl_workbook", {"op": "create_workbook", "args": {"path": book}}),
        ("xl_edit", {"op": "write_range", "args": {"path": book, "location": "A1",
                                                   "data": [["Name", "Qty"], ["a", 1]]}}),
        ("xl_format", {"op": "format_cells", "args": {"path": book, "location": "A1:B1",
                                                      "font": {"bold": True}}}),
        ("xl_structure", {"op": "manage_worksheet", "args": {"path": book, "action": "add",
                                                             "new_name": "Two"}}),
        ("xl_workbook", {"op": "copy_workbook", "args": {"src": book, "dst": copy_to}}),
        ("xl_read", {"op": "get_workbook_metadata", "args": {"path": book}}),
        # xl_app is called with arguments its op refuses before starting any
        # Excel, so this test never reaches an Excel at all.
        ("xl_app", {"op": "com_goal_seek", "args": {}}),
        ("xl_read", {"op": "get_workbook_metadata", "args": {"path": book, "bogus": 1}}),
    ]
    res = fh.stdio_session(calls)
    assert not (res["capabilities"].get("tools") or {}).get("listChanged")
    assert [t["name"] for t in res["tools"]] == EXPECTED
    assert json.dumps(res["tools_raw"], sort_keys=True) == \
        json.dumps(res["tools_after_raw"], sort_keys=True)
    for t in res["tools"]:
        assert "outputSchema" not in t
    for r in res["results"]:
        assert r["structured"] is None and r["blocks"] == 1, r
        json.loads(r["text"])
    info = _body(res["results"][1])
    assert info["version"] == res["server_version"] and info["edition"] == "community"
    assert info["surface"]["tools"] == 7 and info["surface"]["ops"] == 69
    for r in res["results"][2:8]:
        assert not r["is_error"], r["text"][:300]
    assert _body(res["results"][7])["ok"] is True
    assert _body(res["results"][-2])["error"]["code"] == "BAD_PARAMS"
    bad = res["results"][-1]
    assert bad["is_error"] and _body(bad)["error"]["code"] == "BAD_PARAMS"
    assert "enable_tools" not in res["instructions"]


def test_wire_forbid_and_read_only_refuse_inside_the_call():
    res = fh.stdio_session(
        [("xl_read", {"op": "audit_styles", "args": {"path": "x.xlsx"}}),
         ("xl_edit", {"op": "set_cell", "args": {"path": "x.xlsx"}}),
         ("xl_read", {"op": "get_server_info"})],
        env={"KS4XL_READ_ONLY": "1", "KS4XL_FORBID": "audit_styles"})
    assert [t["name"] for t in res["tools"]] == ["xl_read", "xl_help"]
    a, b, info = res["results"]
    assert a["is_error"] and _body(a)["error"]["code"] == "NOT_AVAILABLE"
    assert _body(a)["error"]["details"]["decided_by"]
    # xl_edit is not listed, but a client calling it by name still reaches
    # the server; the refusal is decided inside the call.
    assert b["is_error"] and _body(b)["error"]["code"] == "NOT_AVAILABLE"
    body = _body(info)
    assert body["surface"]["read_only"] is True and body["surface"]["forbidden"] == ["audit_styles"]


def test_unknown_forbid_name_stops_the_server():
    code, err = fh.stdio_start_only({"KS4XL_FORBID": "no_such_op"})
    assert code == 2 and "KS4XL_FORBID" in err


def test_bad_setting_stops_the_server_with_one_line():
    code, err = fh.stdio_start_only({"KS4XL_READ_ONLY": "maybe"})
    assert code == 2 and "KS4XL_READ_ONLY" in err and "Traceback" not in err


@pytest.mark.parametrize("env", [{"KS4XL_PACK_POLICY": "locked"},
                                 {"KS4XL_MODE": "design,com"},
                                 {"KS4XL_PACK_POLICY": "lockedd"}])
def test_a_retired_lock_setting_stops_the_server_with_one_line(env):
    """KS4XL_PACK_POLICY=locked was the 1.x administrator
    lock. Starting the full surface anyway failed open; now a retired
    setting with a value that limited the tools stops the server and names
    the 2.0 replacements."""
    code, err = fh.stdio_start_only(env)
    name = next(iter(env))
    assert code == 2, err
    assert err.count(name) >= 1 and "KS4XL_READ_ONLY" in err and "KS4XL_FORBID" in err
    assert "Traceback" not in err and err.strip().count(chr(10)) == 0, err


def test_core_refuses_to_start_without_a_policy():
    import subprocess
    import sys
    code = "from xlsx_mcp.server import main; main()"
    r = subprocess.run([sys.executable, "-X", "utf8", "-c", code], stdin=subprocess.DEVNULL,
                       capture_output=True, timeout=60, env=fh._env(None),
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    # The core raises at import (edition split), before any op registers.
    assert r.returncode != 0
    assert b"No license policy is installed" in r.stderr


# ------------------------------------------------------------ load pin

#: The measured surface over real stdio (scripts/measure_surface.py, Excel's
#: counter) was 6,127 tokens with the copy slots UNFILLED
#: (placeholder facts), pinned with a 5% margin (6,433). Filling the slots
#: changes the figure; a raise is a deliberate, reviewed change.
SURFACE_TOKEN_CEILING = 6433


def _measure():
    import importlib.util
    from pathlib import Path
    root = Path(__file__).resolve().parents[2]
    spec = importlib.util.spec_from_file_location("measure_surface",
                                                  root / "scripts" / "measure_surface.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.measure()


def within_ceiling(report: dict, ceiling: int) -> bool:
    return report["approx_tokens"] <= ceiling


def test_dg4_surface_load_is_within_the_pinned_ceiling():
    report = _measure()
    assert report["tools"] == 7 and report["ops"] == 69
    assert within_ceiling(report, SURFACE_TOKEN_CEILING), report["approx_tokens"]
    for row in report["rows"]:
        assert row["description_chars"] <= 2048 and row["schema_bytes"] <= 5000, row


def test_dg4_red_the_pin_fails_over_its_ceiling():
    assert not within_ceiling({"approx_tokens": SURFACE_TOKEN_CEILING + 1},
                              SURFACE_TOKEN_CEILING)
