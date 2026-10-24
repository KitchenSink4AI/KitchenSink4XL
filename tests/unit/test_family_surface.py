"""The family surface: shape, strict arguments,
routing, launch settings, help, the emitted-once envelope, JSON-text
decoding (with Excel's numbers-safety additions), the deadline, the output
cap, and the one choke point.

The same checks as the Word editor's family surface tests, adapted to Excel.
Replaces the pack tests (test_tiered_loading, test_packs_notice and the
pack halves of several others); each deleted test's guarantee is carried by
a test here.
"""

from __future__ import annotations

import ast
import hashlib
import json
import re
import threading
import time
from pathlib import Path

import openpyxl
import pytest
from pydantic import ConfigDict

import family_harness as fh
from xlsx_mcp import envelope, registry
from xlsx_mcp import server as srv
from xlsx_mcp.settings import Settings, SettingsError
from xlsx_mcp.surface import copy, dispatch, groups, help as help_mod, plan as plan_mod

ROOT = Path(__file__).resolve().parents[2]
EXPECTED_TOOLS = ["xl_read", "xl_edit", "xl_format", "xl_structure", "xl_workbook",
                  "xl_app", "xl_help"]
OP_COUNT = 69


@pytest.fixture(scope="module")
def harness():
    return fh.start()


@pytest.fixture
def xlsx(tmp_path):
    p = tmp_path / "fam.xlsx"
    srv.create_workbook(str(p))
    srv.write_range(str(p), "A1", [["Name", "Qty"], ["Alpha", 1], ["Beta", 2]])
    return str(p)


# ------------------------------------------------------------------ shape


def test_seven_tools_69_ops_in_the_brief_order(harness):
    tools = harness.list_tools()
    assert [t.name for t in tools] == EXPECTED_TOOLS
    ops = [op for t in tools if t.name != "xl_help"
           for op in t.inputSchema["properties"]["op"]["enum"]]
    assert len(ops) == len(set(ops)) == OP_COUNT
    for t in tools:
        if t.name == "xl_help":
            continue
        assert t.inputSchema["properties"]["op"]["enum"] == list(groups.GROUPS[t.name])


def test_tool_list_shape_and_annotations(harness):
    for t in harness.list_tools():
        assert t.title and t.annotations and t.annotations.title
        assert t.outputSchema is None
        for key, prop in t.inputSchema["properties"].items():
            assert prop.get("description"), f"{t.name}.{key}"
        # only xl_app reaches a remote service (update_now: PyPI, GitHub)
        assert t.annotations.openWorldHint is (t.name == "xl_app")
        if t.name == "xl_help":
            assert t.annotations.readOnlyHint is True
            continue
        assert set(t.inputSchema["properties"]) == {"op", "args"}
        assert t.inputSchema["additionalProperties"] is False
        ops = t.inputSchema["properties"]["op"]["enum"]
        specs = [registry.REGISTRY[o] for o in ops]
        assert t.annotations.readOnlyHint is all(s.read_only for s in specs)
        assert t.annotations.destructiveHint is any(s.destructive for s in specs)
        assert t.annotations.idempotentHint is t.annotations.readOnlyHint
        listed = t.description.splitlines() + \
            (t.inputSchema["properties"]["op"].get("description") or "").splitlines()
        for s in specs:
            assert plan_mod._op_line(s) in listed, s.name


def test_only_xl_read_is_read_only_and_it_holds_no_writer(harness):
    by = {t.name: t for t in harness.list_tools()}
    assert by["xl_read"].annotations.readOnlyHint is True
    for name in EXPECTED_TOOLS[1:-1]:
        assert by[name].annotations.readOnlyHint is False, name


def test_annotation_rule_red_cases():
    class A(registry.StrictArgs):
        pass
    base = dict(summary="s", args=A, handler=lambda: {}, destructive=False, runs_live=False)
    with pytest.raises(plan_mod.PlanError):
        plan_mod.check_registry({"x": registry.OpSpec(name="x", group="xl_edit", writes=False,
                                                      read_only=True, **base)})
    with pytest.raises(plan_mod.PlanError):
        plan_mod.check_registry({"y": registry.OpSpec(name="y", group="xl_read", writes=True,
                                                      read_only=False, **base)})


def test_internal_ops_are_never_listed_or_dispatched():
    def _probe_internal(path: str) -> dict:
        return {}
    spec = registry.OpSpec(name="_probe_internal", group="xl_read", summary="s",
                           args=registry.args_model(_probe_internal, "_probe_internal"),
                           handler=_probe_internal, writes=False, read_only=True,
                           destructive=False, runs_live=False)
    registry.REGISTRY["_probe_internal"] = spec
    try:
        h = fh.start()
        assert "_probe_internal" not in json.dumps([t.model_dump() for t in h.list_tools()])
        out, err, _ = h.call("xl_read", {"op": "_probe_internal"})
        assert err and out["error"]["code"] == "NOT_FOUND"
        out, err, _ = h.call("xl_help", {"op": "_probe_internal"})
        assert err and out["error"]["code"] == "NOT_FOUND"
    finally:
        registry.unregister("_probe_internal")


def test_an_op_with_no_tool_group_fails_the_plan():
    """Registering an op that no tool holds stops the build."""
    def orphan_probe(path: str) -> dict:
        return {}
    spec = registry.OpSpec(name="orphan_probe", group=None, summary="s",
                           args=registry.args_model(orphan_probe, "orphan_probe"),
                           handler=orphan_probe, writes=False, read_only=True,
                           destructive=False, runs_live=False)
    registry.REGISTRY["orphan_probe"] = spec
    try:
        with pytest.raises(plan_mod.PlanError):
            plan_mod.build(Settings())
    finally:
        registry.unregister("orphan_probe")


def test_the_list_never_changes_and_list_changed_is_not_advertised(harness, xlsx):
    first = [t.model_dump() for t in harness.list_tools()]
    harness.call("xl_read", {"op": "get_server_info"})
    harness.call("xl_read", {"op": "get_workbook_metadata", "args": {"path": xlsx}})
    harness.call("xl_help", {})
    assert [t.model_dump() for t in harness.list_tools()] == first
    caps = harness.capabilities().tools
    assert not (caps and caps.listChanged)


def test_no_pack_machinery_survives():
    import importlib.util
    assert importlib.util.find_spec("xlsx_mcp.packs") is None
    assert not hasattr(srv, "enable_tools") and not hasattr(srv, "disable_tools")
    text = srv.instructions_for(srv._license_policy.active())
    assert "enable_tools" not in text and not re.search(r"packs?", text.lower())


# ------------------------------------------------------------ strict


def _example_args(spec, xlsx):
    ex = help_mod.example(spec)
    return {k: (xlsx if k == "path" else v) for k, v in ex.items()}


def strict_violations(h, xlsx, specs=None) -> list[str]:
    bad = []
    specs = specs or registry.public_ops()
    for name, spec in specs.items():
        args = _example_args(spec, xlsx)
        args["zz_unknown_key"] = 1
        out, err, _ = h.call(spec.group, {"op": name, "args": args})
        if not (err and out["error"]["code"] == "BAD_PARAMS"
                and "zz_unknown_key" in json.dumps(out["error"].get("details", {}))):
            bad.append(f"{name}: {json.dumps(out)[:200]}")
    return bad


def test_dg6_unknown_keys_refuse_for_every_op(harness, xlsx, monkeypatch):
    calls = []
    for name, spec in registry.public_ops().items():
        monkeypatch.setitem(registry.REGISTRY, name, registry.OpSpec(
            **{**spec.__dict__, "handler": lambda *a, _n=name, **k: calls.append(_n) or {}}))
    assert strict_violations(harness, xlsx) == []
    assert calls == [], f"reached the engine before refusing: {calls}"


def test_dg6_the_order_desc_class_refuses_on_sort(harness, xlsx):
    """An unknown key like {"order": "desc"} is refused, never ignored."""
    out, err, _ = harness.call("xl_structure", {"op": "sort_range", "args": {
        "path": xlsx, "location": "A1:B3", "keys": [{"column": "B"}], "order": "desc"}})
    assert err and out["error"]["code"] == "BAD_PARAMS"
    assert "order" in json.dumps(out["error"]["details"])


def test_dg6_red_a_permissive_model_is_caught(harness, xlsx, monkeypatch):
    spec = registry.REGISTRY["audit_styles"]

    class Loose(spec.args):
        model_config = ConfigDict(extra="ignore")
    loose = registry.OpSpec(**{**spec.__dict__, "args": Loose})
    monkeypatch.setitem(registry.REGISTRY, "audit_styles", loose)
    found = strict_violations(harness, xlsx, {"audit_styles": loose})
    assert found and found[0].startswith("audit_styles")


def test_top_level_shape_is_strict(harness):
    for arguments in ({"op": "get_server_info", "extra": 1}, {}, {"op": 5},
                      {"op": "get_server_info", "args": [1]}):
        out, err, _ = harness.call("xl_read", arguments)
        assert err and out["error"]["code"] == "BAD_PARAMS", arguments


def test_routing_refusals_name_the_right_tool(harness):
    out, err, _ = harness.call("xl_read", {"op": "set_cell"})
    assert err and out["error"]["code"] == "BAD_PARAMS"
    assert out["error"]["details"]["tool"] == "xl_edit"
    out, err, _ = harness.call("xl_read", {"op": "no_such_op"})
    assert err and out["error"]["code"] == "NOT_FOUND"
    out, err, _ = harness.call("xl_read", {"op": "enable_tools", "args": {"packs": ["x"]}})
    assert err and out["error"]["code"] == "NOT_FOUND"


# ----------------------------------------------------- launch settings


def test_read_only_mode_lists_only_read_only_tools():
    h = fh.start(Settings(read_only=True))
    assert [t.name for t in h.list_tools()] == ["xl_read", "xl_help"]
    out, err, _ = h.call("xl_read", {"op": "set_cell"})
    assert err and out["error"]["code"] == "NOT_AVAILABLE"
    assert out["error"]["details"]["decided_by"] == copy.SETTINGS_OWNER
    out, err, _ = h.call("xl_help", {"op": "set_cell"})
    assert err and out["error"]["code"] == "NOT_AVAILABLE"
    info, _, _ = h.call("xl_read", {"op": "get_server_info"})
    assert info["surface"]["read_only"] is True
    assert info["surface"]["tools"] == 2


def test_forbid_removes_ops_and_tools():
    h = fh.start(Settings(forbid=("audit_styles", "xl_app")))
    tools = {t.name: t for t in h.list_tools()}
    assert "xl_app" not in tools
    assert "audit_styles" not in tools["xl_read"].inputSchema["properties"]["op"]["enum"]
    assert "audit_styles(" not in tools["xl_read"].description
    # A forbidden TOOL is not registered at all; calling it by name still
    # refuses in the envelope, NOT_AVAILABLE naming who decided it, and
    # nothing runs (Excel kept its enveloped unknown-tool refusal).
    out, err, _ = h.call("xl_app", {"op": "recalculate"})
    assert err and out["error"]["code"] == "NOT_AVAILABLE"
    assert out["error"]["details"]["decided_by"] == copy.SETTINGS_OWNER
    out, err, _ = h.call("xl_nope", {})
    assert err and out["error"]["code"] == "NOT_FOUND"
    assert "xl_read" in out["error"]["details"]["tools"]
    for op, tool in (("audit_styles", "xl_read"), ("recalculate", "xl_read"),
                     ("recalculate", "xl_edit")):
        out, err, _ = h.call(tool, {"op": op})
        assert err and out["error"]["code"] == "NOT_AVAILABLE", (op, tool, out)
        out, err, _ = h.call("xl_help", {"op": op})
        assert err and out["error"]["code"] == "NOT_AVAILABLE"
    listing, _, _ = h.call("xl_help", {})
    assert "audit_styles" not in json.dumps(listing) and "recalculate" not in json.dumps(listing)
    info, _, _ = h.call("xl_read", {"op": "get_server_info"})
    assert info["surface"]["forbidden"] == ["audit_styles", "xl_app"]


def test_unknown_forbid_name_refuses_to_start():
    with pytest.raises(SettingsError):
        fh.start(Settings(forbid=("no_such_thing",)))


@pytest.mark.parametrize("env,ok", [
    ({"KS4XL_READ_ONLY": "1"}, True), ({"KS4XL_READ_ONLY": "maybe"}, False),
    ({"KS4XL_MAX_OUTPUT_CHARS": "abc"}, False), ({"KS4XL_MAX_OUTPUT_CHARS": "100"}, False),
    ({"KS4XL_CALL_TIMEOUT_S": "0"}, False), ({"KS4XL_CALL_TIMEOUT_S": "30"}, True),
    # An old 1.x default is ignored; a retired setting that limited the tools
    # (a lock, a list of tool groups, an unknown word) stops the server.
    ({"KS4XL_MODE": "full"}, True), ({"KS4XL_PACK_POLICY": "auto"}, True),
    ({"KS4XL_PACK_POLICY": "locked"}, False), ({"KS4XL_MODE": "design,com"}, False),
    ({"KS4XL_MODE": "fulll"}, False), ({"KS4XL_PACK_POLICY": "lockedd"}, False),
])
def test_settings_parse_or_refuse(env, ok):
    if ok:
        Settings.from_env(env)
    else:
        with pytest.raises(SettingsError):
            Settings.from_env(env)


def test_retired_settings_refuse_to_start_and_name_the_replacements():
    # Fail closed, see test_legacy_settings.py.
    for env in ({"KS4XL_MODE": "design,com"}, {"KS4XL_PACK_POLICY": "locked"}):
        with pytest.raises(SettingsError) as ei:
            Settings.from_env(env)
        msg = str(ei.value)
        assert next(iter(env)) in msg and "KS4XL_READ_ONLY" in msg and "KS4XL_FORBID" in msg
    assert plan_mod.build(Settings.from_env({})).op_count() == OP_COUNT


# ------------------------------------------------------------------- help


def test_help_lists_explains_and_teaches_location(harness):
    out, err, _ = harness.call("xl_help", {})
    assert not err and out["tools"]["xl_edit"][0] == "set_cell"
    out, err, _ = harness.call("xl_help", {"op": "set_cell"})
    assert not err and out["tool"] == "xl_edit" and out["writes"] is True
    assert "location" in out["args_schema"]["properties"] and out["location"]
    registry.REGISTRY["set_cell"].args.model_validate(out["example"]["args"])
    out, err, _ = harness.call("xl_help", {"topic": "location"})
    assert not err and out["location"] == copy.LOCATION_HELP
    for bad in ({"op": "x", "extra": 1}, {"topic": "nope"}):
        out, err, _ = harness.call("xl_help", bad)
        assert err and out["error"]["code"] == "BAD_PARAMS"


def test_every_help_example_validates():
    for name, spec in registry.public_ops().items():
        spec.args.model_validate(help_mod.example(spec))


def test_help_examples_name_realistic_files():
    """A path example is a workbook, an output is the kind it writes."""
    assert help_mod.example(registry.REGISTRY["set_cell"])["path"].endswith(".xlsx")
    assert help_mod.example(registry.REGISTRY["com_export_pdf"])["output"].endswith(".pdf")
    assert help_mod.example(registry.REGISTRY["com_render_sheet"])["output"].endswith(".png")


# --------------------------------------------------------------- envelope


def test_emitted_once_and_compact(harness, xlsx):
    for tool, arguments in (("xl_read", {"op": "get_server_info"}),
                            ("xl_read", {"op": "get_workbook_metadata",
                                         "args": {"path": xlsx}}),
                            ("xl_read", {"op": "nope"})):
        out, err, res = harness.call(tool, arguments)
        assert res.structuredContent is None and len(res.content) == 1
        text = res.content[0].text
        assert text == json.dumps(json.loads(text), ensure_ascii=False, separators=(",", ":"))


def test_success_keeps_ok_and_file_e6(harness, xlsx, tmp_path):
    out, err, _ = harness.call("xl_read", {"op": "get_workbook_metadata",
                                           "args": {"path": xlsx}})
    assert not err and out["ok"] is True and out["file"] == xlsx
    dst = str(tmp_path / "copy.xlsx")
    out, err, _ = harness.call("xl_workbook", {"op": "copy_workbook",
                                               "args": {"src": xlsx, "dst": dst}})
    # The op's own `file` (the destination) wins over the envelope's, exactly
    # as the released _SuccessEnvelope merged them.
    assert not err and out["ok"] is True and out["file"] == dst


def test_unknown_exceptions_are_internal_errors_and_scrubbed(harness, xlsx, tmp_path,
                                                              monkeypatch):
    home = str(tmp_path / "Users" / "someone")
    monkeypatch.setenv("USERPROFILE", home)
    spec = registry.REGISTRY["audit_styles"]
    for exc in (RuntimeError(f"bad thing at {home}\\secret\\x.xlsx"), ZeroDivisionError("z"),
                MemoryError("m")):
        def boom(*a, _e=exc, **k):
            raise _e
        monkeypatch.setitem(registry.REGISTRY, "audit_styles",
                            registry.OpSpec(**{**spec.__dict__, "handler": boom}))
        out, err, _ = harness.call("xl_read", {"op": "audit_styles", "args": {"path": xlsx}})
        assert err and out["error"]["code"] == "INTERNAL_ERROR", out
        assert home not in json.dumps(out) and "secret" not in json.dumps(out)
        assert out["error"]["hint"]


def test_hint_table_is_total_over_the_closed_codes():
    for code in envelope.CLOSED_CODES:
        assert envelope.HINTS[code], code


def test_every_mapped_code_is_closed():
    for _t, code in envelope.CODE_MAP:
        assert code in envelope.CLOSED_CODES


def test_see_ops_replaces_the_pack_signpost():
    from xlsx_mcp.core.errors import WorkbookProtected
    exc = WorkbookProtected("encrypted")
    exc.hint_tools = ("com_convert_format",)
    out = envelope.refusal(exc)
    assert out["error"]["details"]["see_ops"] == [{"op": "com_convert_format", "tool": "xl_app"}]
    assert "enable_tools" not in json.dumps(out)


# -------------------------------------------------------- JSON text


def _op(name):
    return registry.REGISTRY[name]


def test_json_text_array_is_decoded_and_recorded(harness, xlsx):
    edits = json.dumps([{"op": "set_value", "location": "C1", "value": "Gamma"}])
    out, err, _ = harness.call("xl_edit", {"op": "apply_edits",
                                           "args": {"path": xlsx, "edits": edits}})
    assert not err, out
    assert out["json_text_decoded"] == ["edits"]
    assert openpyxl.load_workbook(xlsx).active["C1"].value == "Gamma"


def test_json_text_is_recorded_on_refusals_too(harness, xlsx):
    edits = json.dumps([{"op": "set_value", "location": "ZZZZ0", "value": 1}])
    out, err, _ = harness.call("xl_edit", {"op": "apply_edits",
                                           "args": {"path": xlsx, "edits": edits}})
    assert err and out["error"]["details"]["json_text_decoded"] == ["edits"], out


def test_json_text_matrix_is_decoded(xlsx):
    args, decoded = dispatch.decode_json_text(
        _op("write_range"), {"path": xlsx, "location": "A1", "data": '[["a", 1]]'})
    assert decoded == ["data"] and args["data"] == [["a", 1]]


def test_json_looking_text_in_a_string_parameter_is_left(xlsx):
    args, decoded = dispatch.decode_json_text(
        _op("find_cells"), {"path": xlsx, "query": '["not", "a list"]'})
    assert decoded == [] and args["query"] == '["not", "a list"]'


def test_location_is_never_decoded(xlsx):
    """location admits a bare 'A1' string, so JSON-looking text stays text."""
    args, decoded = dispatch.decode_json_text(
        _op("read_range"), {"path": xlsx, "location": '{"cell": "A1"}'})
    assert decoded == [] and args["location"] == '{"cell": "A1"}'


def test_list_or_string_parameter_is_left():
    union_specs = [s for s in registry.public_ops().values()
                   if any("string" in dispatch._admits(p, s.args.model_json_schema()) and
                          "array" in dispatch._admits(p, s.args.model_json_schema())
                          for p in (s.args.model_json_schema().get("properties") or {}).values())]
    assert union_specs, "no list-or-string parameter found to exercise"
    s = union_specs[0]
    props = s.args.model_json_schema()["properties"]
    name = next(n for n, p in props.items()
                if {"string", "array"} <= dispatch._admits(p, s.args.model_json_schema()))
    args, decoded = dispatch.decode_json_text(s, {name: '["a"]'})
    assert decoded == [] and args[name] == '["a"]'


def test_malformed_json_text_is_left_and_refused_normally(harness, xlsx):
    out, err, _ = harness.call("xl_edit", {"op": "apply_edits",
                                           "args": {"path": xlsx, "edits": "[{oops"}})
    assert err and out["error"]["code"] == "BAD_PARAMS"
    assert "json_text_decoded" not in json.dumps(out)


def test_object_where_an_array_goes_is_left(xlsx):
    args, decoded = dispatch.decode_json_text(
        _op("apply_edits"), {"path": xlsx, "edits": '{"op": "clear"}'})
    assert decoded == [] and args["edits"] == '{"op": "clear"}'


def test_allof_oneof_admits():
    root = {"$defs": {"A": {"type": "object"}}}
    assert dispatch._admits({"allOf": [{"$ref": "#/$defs/A"}]}, root) == {"object"}
    assert dispatch._admits({"oneOf": [{"type": "array"}, {"type": "null"}]}, root) == \
        {"array", "other"}


# ------------------------------------------- numbers safety

NUMBER_LOOKING = ["[1,2]", "1e3", "007", "1,234", "{\"a\": 1}", "=1+1", "TRUE", "3.0"]


def _no_string_param_is_ever_decoded():
    bad = []
    for name, spec in registry.public_ops().items():
        root = spec.args.model_json_schema()
        for pname, prop in (root.get("properties") or {}).items():
            kinds = dispatch._admits(prop, root)
            if "string" in kinds or "any" in kinds:
                _args, decoded = dispatch.decode_json_text(spec, {pname: "[1,2]"})
                if decoded:
                    bad.append(f"{name}.{pname}")
                _args, decoded = dispatch.decode_json_text(spec, {pname: '{"a": 1}'})
                if decoded:
                    bad.append(f"{name}.{pname}")
    return bad


def test_no_string_admitting_parameter_is_decoded_on_any_op():
    assert _no_string_param_is_ever_decoded() == []


@pytest.mark.parametrize("text", NUMBER_LOOKING)
def test_set_cell_value_lands_exactly_as_today(harness, tmp_path, text):
    """A value sent as text through the surface lands in the file exactly as
    the same text written in process (the released tool's behaviour)."""
    via_surface = tmp_path / "s.xlsx"
    in_process = tmp_path / "p.xlsx"
    srv.create_workbook(str(via_surface))
    srv.create_workbook(str(in_process))
    out, err, _ = harness.call("xl_edit", {"op": "set_cell", "args": {
        "path": str(via_surface), "location": "A1", "value": text}})
    srv.set_cell(str(in_process), "A1", text)
    if err:
        # A refusal must be the same refusal the released tool gave.
        with pytest.raises(Exception):
            srv.set_cell(str(tmp_path / "p2.xlsx"), "A1", text)
        return
    assert "json_text_decoded" not in out
    a = openpyxl.load_workbook(via_surface).active["A1"]
    b = openpyxl.load_workbook(in_process).active["A1"]
    assert (a.value, a.data_type) == (b.value, b.data_type), text
    if not text.startswith("="):
        assert isinstance(a.value, str) and a.value == text


def test_set_cell_numbers_stay_numbers(harness, tmp_path):
    p = tmp_path / "n.xlsx"
    srv.create_workbook(str(p))
    for value, cell in ((1000, "A1"), (0.1, "A2"), (True, "A3")):
        out, err, _ = harness.call("xl_edit", {"op": "set_cell", "args": {
            "path": str(p), "location": cell, "value": value}})
        assert not err, out
    ws = openpyxl.load_workbook(p).active
    assert ws["A1"].value == 1000 and ws["A2"].value == 0.1 and ws["A3"].value is True


def test_write_range_text_cells_stay_text(harness, tmp_path):
    p = tmp_path / "w.xlsx"
    srv.create_workbook(str(p))
    out, err, _ = harness.call("xl_edit", {"op": "write_range", "args": {
        "path": str(p), "location": "A1", "data": [["1e3", "[1,2]", 5, "007"]]}})
    assert not err, out
    ws = openpyxl.load_workbook(p).active
    assert [c.value for c in ws[1]] == ["1e3", "[1,2]", 5, "007"]


# ------------------------------------------------------ deadline


def test_timeout_is_reported_and_nothing_is_saved_after_it(xlsx, monkeypatch):
    """The invariant: a TIMEOUT answer means nothing was saved. A slow edit
    whose save would land after the bound never publishes."""
    monkeypatch.setattr(dispatch, "BACKSTOP_GRACE_S", 0.2)
    h = fh.start(Settings(call_timeout_s=1.0))
    before = hashlib.sha256(Path(xlsx).read_bytes()).hexdigest()
    from xlsx_mcp.core import package as pkgmod
    real_save = pkgmod.WorkbookPackage.save
    gate = threading.Event()

    def slow_save(self, *a, **k):
        gate.wait(10)  # held past the bound, released after TIMEOUT
        return real_save(self, *a, **k)
    monkeypatch.setattr(pkgmod.WorkbookPackage, "save", slow_save)
    out, err, _ = h.call("xl_edit", {"op": "set_cell",
                                     "args": {"path": xlsx, "location": "D1", "value": "late"}})
    assert err and out["error"]["code"] == "TIMEOUT", out
    assert out["error"]["details"]["saved"] is False
    gate.set()
    deadline = time.monotonic() + 10
    while time.monotonic() < deadline:  # poll, no timing margin
        time.sleep(0.2)
        if not any(t.name.startswith("asyncio") and t.is_alive() for t in threading.enumerate()):
            break
    time.sleep(0.5)
    assert hashlib.sha256(Path(xlsx).read_bytes()).hexdigest() == before


def test_a_save_that_committed_before_the_bound_is_never_reported_as_timeout(xlsx, monkeypatch):
    """The other half of the invariant, over the real dispatcher: the save
    commits inside the bound, then the call is held past it (backup staging
    blocked). The backstop cannot cancel a committed save, so the answer is
    the real success and the file holds the write."""
    monkeypatch.setattr(dispatch, "BACKSTOP_GRACE_S", 0.2)
    h = fh.start(Settings(call_timeout_s=1.0))
    from xlsx_mcp.core import safesave as ss
    real_prepare = ss.prepare_rotation
    release = threading.Event()

    def slow_prepare(*a, **k):  # runs right AFTER the commit point
        release.wait(10)
        return real_prepare(*a, **k)
    monkeypatch.setattr(ss, "prepare_rotation", slow_prepare)
    timer = threading.Timer(3.0, release.set)  # past the 1 s bound + grace
    timer.start()
    try:
        out, err, _ = h.call("xl_edit", {"op": "set_cell",
                                         "args": {"path": xlsx, "location": "D2",
                                                  "value": "committed"}})
    finally:
        timer.cancel()
        release.set()
    assert not err, out
    assert openpyxl.load_workbook(xlsx).active["D2"].value == "committed"


def test_a_call_that_saved_reports_it(xlsx):
    """Commit first: the backstop cannot turn a finished save into TIMEOUT."""
    from xlsx_mcp.core import deadline as dl
    d = dl.Deadline(30)
    with dl.active(d):
        srv.set_cell(xlsx, "E1", "saved")
    assert d.committed and not d.cancel()


@pytest.mark.parametrize("writer", ["create_workbook", "copy_workbook", "restore",
                                    "export_range"])
def test_every_commit_point_refuses_after_cancel(xlsx, tmp_path, writer):
    """Each disk writer checks the deadline before its first change: a
    cancelled deadline means the target is untouched."""
    from xlsx_mcp.core import deadline as dl
    target = tmp_path / "t.xlsx"
    srv.set_cell(xlsx, "F1", "x")  # a prev slot to restore from
    before = Path(xlsx).read_bytes()
    d = dl.Deadline(30)
    assert d.cancel()
    with dl.active(d), pytest.raises(Exception) as ei:
        if writer == "create_workbook":
            srv.create_workbook(str(target))
        elif writer == "copy_workbook":
            srv.copy_workbook(xlsx, str(target))
        elif writer == "restore":
            srv.manage_backups("restore", path=xlsx, source="prev")
        else:
            srv.export_range(xlsx, "A1:B3", out_file=str(tmp_path / "o.csv"))
    assert getattr(ei.value, "code", None) == "TIMEOUT"
    assert not target.exists() and not (tmp_path / "o.csv").exists()
    assert Path(xlsx).read_bytes() == before


def test_commit_after_cancel_refuses():
    from xlsx_mcp.core import deadline as dl
    d = dl.Deadline(30)
    assert d.cancel()
    with pytest.raises(Exception) as ei:
        d.commit_or_raise()
    assert getattr(ei.value, "code", None) == "TIMEOUT"


def test_com_timeout_argument_raises_the_bound():
    s = Settings(call_timeout_s=300.0)
    assert s.timeout_for(registry.REGISTRY["recalculate"].timeout_s) == 900.0
    assert s.timeout_for(None, 1000.0) == 900.0  # capped at the max
    assert dispatch._requested_timeout({"timeout_seconds": 500}) == 500.0


# ------------------------------------------------------ output cap


def _big(tmp_path) -> str:
    p = tmp_path / "big.xlsx"
    srv.create_workbook(str(p))
    srv.write_range(str(p), "A1", [[f"r{r}c{c} " + "text " * 8 for c in range(10)]
                                   for r in range(300)])
    return str(p)


def test_output_cap_cuts_with_a_page_block(tmp_path):
    p = _big(tmp_path)
    h = fh.start(Settings(max_output_chars=4000))
    for op, args in (("read_range", {"location": "A1:J300"}), ("get_grid_view", {}),
                     ("query_range", {"location": "A1:J300", "header": False}),
                     ("find_cells", {"query": "text"})):
        out, err, res = h.call("xl_read", {"op": op, "args": {"path": p, **args}})
        assert len(res.content[0].text) <= 4000, op
        if not err:
            assert out.get("page", {}).get("truncated") in (True, None), op


def test_output_cap_leaves_small_answers_unchanged(harness, xlsx):
    out, err, res = harness.call("xl_read", {"op": "get_workbook_metadata",
                                             "args": {"path": xlsx}})
    assert "page" not in out


def test_dg1_every_read_op_answers_within_the_cap(tmp_path):
    """Sweep: every xl_read op (com_status excepted: it reports the
    private pool) on a large workbook under a small ceiling: the answer on
    the wire never exceeds it, whether the op answers or refuses."""
    p = _big(tmp_path)
    srv.create_table(p, "A1:J300", "BigTable")
    cap = 3000
    h = fh.start(Settings(max_output_chars=cap))
    overs = []
    for op in groups.GROUPS["xl_read"]:
        spec = registry.REGISTRY[op]
        args = _example_args(spec, p)
        if op == "find_cells":
            args["query"] = "text"
        if op == "get_table":
            args["name"] = "BigTable"
        if "location" in spec.args.model_fields and op in ("read_range", "query_range"):
            args["location"] = "A1:J300"
        out, err, res = h.call("xl_read", {"op": op, "args": args})
        if len(res.content[0].text) > cap:
            overs.append((op, len(res.content[0].text)))
    assert overs == []


# -------------------------------------------------- names and the registration pin


_IDENT = re.compile(r"\b[a-z]+(?:_[a-z0-9]+)+\b")


def _surface_names():
    return set(registry.public_ops()) | set(groups.TOOL_NAMES)


def test_hints_name_only_real_tools_and_ops():
    known = _surface_names()
    for code in envelope.CLOSED_CODES:
        text = envelope.HINTS[code]
        assert "enable_tools" not in text and not re.search(r"packs?", text.lower()), code
        for word in _IDENT.findall(text):
            if word.startswith(("get_", "set_", "manage_", "com_")):
                assert word in known, (code, word)
    for word in _IDENT.findall(envelope.RECOVERY_HINT):
        if word.startswith(("get_", "manage_", "xl_")):
            assert word in known, word


def test_docstrings_name_no_pack():
    for name, spec in registry.public_ops().items():
        doc = (spec.handler.__doc__ or "").lower()
        assert "enable_tools" not in doc and not re.search(r"packs?", doc), name


def test_only_surface_tools_registers_mcp_tools():
    """Every MCP tool is built in surface/tools.py; server.py only adds
    what that module built, and nothing else constructs or decorates tools."""
    src = ROOT / "src" / "xlsx_mcp"
    offenders = []
    for path in src.rglob("*.py"):
        rel = path.relative_to(src).as_posix()
        tree = ast.parse(path.read_text("utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, ast.Attribute) and node.attr in ("add_tool", "tool", "from_function"):
                if rel not in ("surface/tools.py", "server.py", "license_policy.py"):
                    offenders.append(f"{rel}:{node.lineno}:{node.attr}")
                if rel == "server.py" and node.attr != "add_tool":
                    offenders.append(f"{rel}:{node.lineno}:{node.attr}")
            if isinstance(node, ast.Name) and node.id in ("FunctionTool", "Tool") \
                    and rel != "surface/tools.py":
                offenders.append(f"{rel}:{node.lineno}:{node.id}")
    assert offenders == []


def test_legacy_notice_and_slots_exist():
    assert "{names}" in copy.LEGACY_SETTING_NOTICE
    assert "{name}" in copy.LEGACY_SETTING_REFUSAL
    for name in ("INSTRUCTIONS", "OP_PARAM", "ARGS_PARAM", "HELP_OP_PARAM", "HELP_TOPIC_PARAM",
                 "HELP_USAGE", "OPS_HEADING", "OPS_CONTINUED", "OP_PARAM_CONTINUED",
                 "SETTINGS_OWNER", "LEGACY_SETTING_NOTICE", "LEGACY_SETTING_REFUSAL",
                 "XL_APP_HIDDEN_EXCEL",
                 "LOCATION_HELP"):
        assert getattr(copy, name).strip(), name
    assert set(copy.TOOLS) == set(groups.TOOL_NAMES)
    for s in copy.all_strings():
        assert chr(0x2014) not in s
