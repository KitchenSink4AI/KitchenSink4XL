"""The old-tool -> new-op map: complete, real, and not stale."""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path

import pytest

from xlsx_mcp import registry
from xlsx_mcp.settings import Settings
from xlsx_mcp.surface import plan

ROOT = Path(__file__).resolve().parents[2]
RELEASED = ROOT / "migration" / "released_tools_1.2.5.json"
MAP = ROOT / "migration" / "v1_to_v2.json"
TABLE = ROOT / "docs" / "surface_map.md"


def _gen():
    spec = importlib.util.spec_from_file_location("gen_map", ROOT / "scripts" /
                                                  "generate_surface_map.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_a_every_op_is_on_exactly_one_visible_tool():
    p = plan.build(Settings())
    seen: dict[str, str] = {}
    for name, t in p.tools.items():
        if name == plan.HELP_TOOL:
            continue
        for op in t.ops:
            assert op not in seen, (op, seen.get(op), name)
            seen[op] = name
    assert set(seen) == set(registry.public_ops())


def test_b_every_released_tool_has_an_entry_that_names_something_real():
    released = json.loads(RELEASED.read_text("utf-8"))["tools"]
    mapping = json.loads(MAP.read_text("utf-8"))["tools"]
    assert set(released) == set(mapping)
    p = plan.build(Settings())
    for old, entry in mapping.items():
        if entry is None or entry.get("tool") is None:
            assert old in ("enable_tools", "disable_tools"), old
            continue
        for dest in [entry, *entry.get("also", [])]:
            assert dest["op"] in p.tools[dest["tool"]].ops, (old, dest)


def test_c_checked_in_table_and_json_are_fresh():
    gen = _gen()
    table, data = gen.render()
    assert TABLE.read_text("utf-8") == table, "docs/surface_map.md is stale; rerun the script"
    assert json.loads(MAP.read_text("utf-8")) == data, "migration map is stale"


def test_d_red_an_op_with_no_group_breaks_the_map():
    def mapless_probe(path: str) -> dict:
        return {}
    registry.REGISTRY["mapless_probe"] = registry.OpSpec(
        name="mapless_probe", group=None, summary="s",
        args=registry.args_model(mapless_probe, "mapless_probe"), handler=mapless_probe,
        writes=False, read_only=True, destructive=False, runs_live=False)
    try:
        with pytest.raises(plan.PlanError):
            _gen().render()
    finally:
        registry.unregister("mapless_probe")
