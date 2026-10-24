"""The op sweep: every registered op, called once through its
tool over REAL stdio (the Community launcher), on a canary workbook.

1. Every op either succeeds or refuses with an allowed STATE code; BAD_PARAMS
   or INTERNAL_ERROR from the sweep's own example fails ("the sweep cannot
   exercise this op; give it an example" in surface/examples.py). A small
   MUST_SUCCEED set may not even refuse.
2. No canary text appears in a refusal's code or hint, or on stderr (P8
   scrub). Whether canary text appears in each op's SUCCESS payload is
   recorded (the baseline map; not asserted) and
   written to $KS4XL_SWEEP_REPORT when that is set.
3. Outguard, per site: every op that writes an
   output path, given the source workbook itself, its backup slot or its
   backup store as the output WITH overwrite:true, refuses in the guard's own
   words and leaves both byte-identical. (Without overwrite an existing
   target refused CONFLICT first, so the old sweep-wide version stayed green
   with the guard removed.)
4. COM ops run in the server's private hidden Excel pool, never a user's
   Excel; where Excel is not installed (CI) they refuse APP_NOT_RUNNING,
   which is an allowed state.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
from pathlib import Path

import pytest

import family_harness as fh
from xlsx_mcp import registry
from xlsx_mcp import server as srv
from xlsx_mcp.core import safesave
from xlsx_mcp.surface import help as help_mod

CANARY = "K9QCANARY"

#: States a refusal may legitimately report for the sweep's example (the
#: fixture or the environment decides them, not the arguments).
READER_OK_REFUSALS = {"NOT_FOUND", "UNSUPPORTED_CONTENT", "APP_NOT_RUNNING", "APP_BUSY",
                      "APP_BLOCKED", "CALC_UNAVAILABLE"}
WRITER_OK_REFUSALS = READER_OK_REFUSALS | {"CONFLICT", "HAZARD_REFUSED", "VALIDATION_FAILED",
                                           "WORKBOOK_LOCKED"}
MUST_SUCCEED = {"create_workbook", "get_grid_view", "get_server_info", "get_workflows",
                "read_range", "set_cell"}

#: Output-path parameters, per op.
OUTPUT_FILE_ARGS = {
    "copy_workbook": ("dst", {}),
    "export_range": ("out_file", {"location": "A1:B4", "fmt": "csv"}),
    "export_file": ("out_file", {"fmt": "json"}),
    "com_export_pdf": ("output", {}),
    "com_render_sheet": ("output", {}),
    "com_convert_format": ("output", {}),
}
OUTPUT_DIR_ARGS = {
    "export_file": ("out_dir", {"fmt": "csv"}),
    "manage_image": ("out_dir", {"action": "extract"}),
}


def _canary_book(path: Path) -> str:
    p = str(path)
    srv.create_workbook(p)
    srv.write_range(p, "A1", [["Name", "Qty"], [f"{CANARY}_row", 1], ["b", 2], ["c", 3]])
    srv.create_table(p, "A1:B4", "Sales")
    srv.set_formula(p, "B5", "=SUM(B2:B4)")
    srv.write_range(p, "D1", [["Region", "Units"], ["North", 5], ["South", 7]])
    srv.manage_name(p, "add", name="Totals", refers_to="=Sheet1!$B$2:$B$4")
    srv.manage_comment(p, "add", location="A2", text=f"{CANARY}_comment")
    srv.manage_hyperlink(p, "add", location="A3", target="https://example.com",
                         display=f"{CANARY}_link")
    srv.set_header_footer(p, header={"center": f"{CANARY}_header"})
    srv.set_workbook_properties(p, title=f"{CANARY}_title")
    srv.manage_worksheet(p, "add", new_name=f"{CANARY}_sheet")
    return p


def _args_for(name: str, work: Path, source: Path) -> dict:
    spec = registry.REGISTRY[name]
    args = help_mod.example(spec)
    book = work / f"{name}.xlsx"
    if name != "create_workbook":
        shutil.copy2(source, book)
    fields = spec.args.model_fields
    if "path" in args or "path" in fields and fields["path"].is_required():
        args["path"] = str(book)
    if name == "copy_workbook":
        args["src"], args["dst"] = str(book), str(work / f"{name}-copy.xlsx")
    if "output" in args:
        args["output"] = str(work / f"{name}-{Path(args['output']).name}")
    for key in ("out_file", "out_dir", "dest_dir"):
        if key in args:
            args[key] = str(work / f"{name}-{key}")
    return args


@pytest.fixture(scope="module")
def sweep(tmp_path_factory):
    work = tmp_path_factory.mktemp("sweep")
    source = work / "canary-source.xlsx"
    _canary_book(source)
    names = sorted(registry.public_ops())
    calls = [(registry.REGISTRY[n].group, {"op": n, "args": _args_for(n, work, source)})
             for n in names]
    res = fh.stdio_session(calls, timeout=1800)
    return names, calls, res


def test_every_op_succeeds_or_refuses_with_an_allowed_state(sweep):
    names, calls, res = sweep
    bad = []
    for name, (_tool, call), r in zip(names, calls, res["results"]):
        body = json.loads(r["text"])
        if not r["is_error"]:
            continue
        code = body.get("error", {}).get("code")
        spec = registry.REGISTRY[name]
        allowed = READER_OK_REFUSALS if spec.read_only else WRITER_OK_REFUSALS
        if name in MUST_SUCCEED or code not in allowed:
            bad.append(f"{name}: {code}: {body.get('error', {}).get('message', '')[:160]}")
    assert bad == [], ("the sweep cannot exercise these ops; give each an example "
                       "in surface/examples.py:\n" + "\n".join(bad))


def test_no_canary_in_refusal_codes_hints_or_stderr(sweep):
    names, _calls, res = sweep
    leaks = []
    for name, r in zip(names, res["results"]):
        body = json.loads(r["text"])
        if r["is_error"]:
            err = body.get("error", {})
            if CANARY in str(err.get("code")) or CANARY in str(err.get("hint")):
                leaks.append(name)
    assert leaks == []
    assert CANARY not in res["stderr"]


def test_dg10_baseline_map_is_recorded(sweep):
    """Not a gate: which ops return workbook text."""
    names, _calls, res = sweep
    baseline = {name: {"canary_in_success": (not r["is_error"]) and CANARY in r["text"],
                       "outcome": "ok" if not r["is_error"]
                       else json.loads(r["text"]).get("error", {}).get("code")}
                for name, r in zip(names, res["results"])}
    out = os.environ.get("KS4XL_SWEEP_REPORT")
    if out:
        Path(out).write_text(json.dumps(baseline, indent=1, sort_keys=True), "utf-8")
    assert len(baseline) == len(registry.public_ops())


def _digest(p: Path) -> str | None:
    return hashlib.sha256(p.read_bytes()).hexdigest() if p.exists() else None


#: Per-site cases: (op, output key, extra args,
#: target, the words only the GUARD says). Every case passes overwrite:true
#: where the op takes it, so an existing target cannot refuse CONFLICT
#: first: only the output guard stands between the call and the file. The
#: render's source case is not reachable (its .png rule refuses any
#: workbook name first), so its guard is driven with a .png in the store.
_GUARD_SOURCE = "source workbook itself"
_GUARD_STORE = "backup store"
DG12_CASES = [
    ("copy_workbook", "dst", {}, "source", "same file"),  # its own same-file rule
    ("copy_workbook", "dst", {}, "slot", _GUARD_STORE),
    ("export_range", "out_file", {"location": "A1:B4", "fmt": "csv"}, "source", _GUARD_SOURCE),
    ("export_range", "out_file", {"location": "A1:B4", "fmt": "csv"}, "slot", _GUARD_STORE),
    ("export_file", "out_file", {"fmt": "json"}, "source", _GUARD_SOURCE),
    ("export_file", "out_file", {"fmt": "json"}, "slot", _GUARD_STORE),
    ("export_file", "out_dir", {"fmt": "csv"}, "store", _GUARD_STORE),
    ("manage_image", "out_dir", {"action": "extract"}, "store", _GUARD_STORE),
    ("com_export_pdf", "output", {}, "source", _GUARD_SOURCE),
    ("com_export_pdf", "output", {}, "slot", _GUARD_STORE),
    ("com_render_sheet", "output", {}, "store_png", _GUARD_STORE),
    ("com_convert_format", "output", {}, "source", _GUARD_SOURCE),
    ("com_convert_format", "output", {}, "slot", _GUARD_STORE),
]


@pytest.mark.parametrize("op,key,extra,target,words", DG12_CASES,
                         ids=[f"{c[0]}.{c[1]}={c[3]}" for c in DG12_CASES])
def test_dg12_no_output_path_can_clobber_the_source_or_its_backup(op, key, extra, target,
                                                                   words, tmp_path,
                                                                   monkeypatch):
    """Each output site, given the source workbook, its prev slot, or its
    backup store as the output, refuses in the output guard's OWN words and
    leaves both byte-identical. Excel is never started: the guard fires
    before any COM call, and a stand-in fails the test if one is made."""
    from xlsx_mcp.com import session as com_session
    from xlsx_mcp.ops import comtier

    def no_com(*_a, **_k):
        raise AssertionError("reached COM: the output guard did not refuse")
    monkeypatch.setattr(com_session, "run_com", no_com)
    monkeypatch.setattr(comtier, "_clipboard_available", lambda: True)
    source = tmp_path / "guarded.xlsx"
    _canary_book(source)
    srv.set_cell(str(source), "C1", "x")  # guarantees a prev slot
    slot = safesave.slot_dir(str(source)) / safesave.PREV_SLOT
    assert slot.exists()
    dest = {"source": source, "slot": slot, "store": slot.parent,
            "store_png": slot.parent / "render.png"}[target]
    args = {"path": str(source), **extra, key: str(dest)}
    if "overwrite" in registry.REGISTRY[op].args.model_fields:
        args["overwrite"] = True
    if op == "copy_workbook":
        other = tmp_path / "other.xlsx"
        srv.create_workbook(str(other))
        args = {"src": str(source if target == "source" else other), "dst": str(dest),
                "overwrite": True}
    before = (_digest(source), _digest(slot))
    payload, err, _ = fh.start().call(registry.REGISTRY[op].group, {"op": op, "args": args})
    assert err, f"output guard let {op}.{key}={target} through: {payload}"
    assert words in payload["error"]["message"], payload["error"]
    assert (_digest(source), _digest(slot)) == before
    assert not (slot.parent / "render.png").exists()
