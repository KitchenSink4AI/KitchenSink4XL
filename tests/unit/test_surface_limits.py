"""Surface limits, measured on the real catalog.

* Claude Code cuts every MCP tool description at 2,048 characters, so no
  description may be longer; op lines that do not fit move, whole and in
  order, into the `op` parameter's description.
* Codex drops EVERY description inside an input schema larger than 5,000
  bytes (compact JSON), so no schema may be larger.
* Nothing is lost: every visible op's line appears exactly once, whole.
"""

from __future__ import annotations

import json

import pytest

from xlsx_mcp import registry
from xlsx_mcp.settings import Settings
from xlsx_mcp.surface import copy, plan


def _plans():
    ops = registry.public_ops()
    return [plan.build(Settings()),
            plan.build(Settings(read_only=True)),
            plan.build(Settings(forbid=("xl_app", sorted(ops)[0])))]


def test_every_tool_fits_both_clients():
    for p in _plans():
        for name, t in p.tools.items():
            assert len(t.description) <= plan.MAX_DESCRIPTION_CHARS, (name, len(t.description))
            size = plan.schema_bytes(t.input_schema)
            assert size <= plan.MAX_SCHEMA_BYTES, (name, size)
            assert len(name) <= 64 and t.title


def test_every_op_line_is_listed_once_and_whole():
    ops = registry.public_ops()
    for p in _plans():
        for name, t in p.tools.items():
            if name == plan.HELP_TOOL:
                continue
            op_desc = t.input_schema["properties"]["op"]["description"]
            text = t.description + "\n" + op_desc
            for op in t.ops:
                line = plan._op_line(ops[op])
                assert text.count(line) == 1, (name, op)
            overflow = op_desc != copy.OP_PARAM
            assert (copy.OPS_CONTINUED in t.description) == overflow, name
            assert t.input_schema["properties"]["op"]["enum"] == t.ops


def test_red_the_limits_catch_an_oversized_catalog():
    head = copy.TOOLS["xl_read"]["description"]
    lines = [f"- op_{i:03d}: " + "x" * 100 for i in range(80)]
    desc, rest = plan._layout(head, lines)
    assert len(desc) <= plan.MAX_DESCRIPTION_CHARS and rest is not None
    schema = plan._group_schema([f"op_{i:03d}" for i in range(80)], rest)
    assert plan.schema_bytes(schema) > plan.MAX_SCHEMA_BYTES
    assert json.dumps(schema).count("op_079") == 2


@pytest.mark.parametrize("n", [0, 1, 5])
def test_layout_is_plain_when_everything_fits(n):
    head = "Head."
    lines = [f"- op{i}: short" for i in range(n)]
    desc, rest = plan._layout(head, lines)
    assert rest is None and desc == "\n".join([head, copy.OPS_HEADING, *lines])
