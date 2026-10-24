"""get_workflows tests: the task listing, recipe shapes, the registry
guarantee (every non-forthcoming step names a real op on the tool that
holds it; family surface: no step, note or summary names a pack or
enable_tools), the honest forthcoming marking on COM-tier steps, and the
refusal path.
"""

from __future__ import annotations

import json
import re

import pytest

from xlsx_mcp import registry, server  # noqa: F401  (the import registers ops)
from xlsx_mcp.core.errors import XlMcpError
from xlsx_mcp.settings import Settings
from xlsx_mcp.surface import plan as _plan
from xlsx_mcp.ops import workflows as _workflows

EXPECTED_TASKS = {
    "recover-workbook", "merge-workbooks", "report-build", "data-cleanup",
    "formatting-audit-and-fix", "safe-edit-of-rich-workbook",
    "migrate-from-incumbent",
    "migrate-from-v1",  # the migrate recipe
}


def _registered_names() -> set[str]:
    return set(registry.public_ops())


def test_listing_names_every_task():
    out = _workflows.get_workflows()
    assert {t["task"] for t in out["tasks"]} == EXPECTED_TASKS
    for t in out["tasks"]:
        assert t["summary"]


def test_recipe_shape():
    out = _workflows.get_workflows("report-build")
    assert out["task"] == "report-build"
    assert out["summary"]
    assert isinstance(out["notes"], list)
    for step in out["steps"]:
        assert step["tool"] and step["op"] and step["why"]
        assert "pack" not in step


def test_every_step_tool_is_registered_or_forthcoming():
    registered = _registered_names()
    for task, wf in _workflows.WORKFLOWS.items():
        for step in wf["steps"]:
            if step.get("forthcoming"):
                assert step["tool"] in _workflows.FORTHCOMING_TOOLS, \
                    f"{task}: {step['tool']} marked forthcoming but not " \
                    "declared"
                assert step["tool"] not in registered, \
                    f"{task}: {step['tool']} is registered; drop the " \
                    "forthcoming mark"
            else:
                assert step["tool"] in registered, \
                    f"{task}: step names unregistered tool {step['tool']!r}"


def test_every_served_step_names_a_real_tool_and_op():
    """Each served step is a call a client can make exactly as
    written, on the default plan."""
    p = _plan.build(Settings())
    for task in _workflows.WORKFLOWS:
        for step in _workflows.get_workflows(task)["steps"]:
            if step.get("forthcoming"):
                continue
            assert step["op"] in p.tools[step["tool"]].ops, (task, step)


def test_no_recipe_mentions_a_pack_or_enable_tools():
    # migrate-from-v1's generated map names
    # every released 1.2.5 tool, the two removed pack switches included; the
    # recipe's own words still name no pack.
    recipes = [_workflows.get_workflows(t) for t in _workflows.WORKFLOWS]
    for r in recipes:
        r.pop("map", None)
    text = json.dumps([_workflows.get_workflows()] + recipes)
    assert "enable_tools" not in text
    assert not re.search(r"\bpacks?\b", text.lower())


def test_com_steps_say_so_in_the_why():
    for wf in _workflows.WORKFLOWS.values():
        for step in wf["steps"]:
            if step.get("forthcoming"):
                assert "xl_app" in step["why"], \
                    "a forthcoming step must say xl_app is coming"


def test_migrate_recipe_is_note_free():
    out = _workflows.get_workflows("migrate-from-incumbent")
    assert out["notes"] == []
    # every step's rationale names what it replaces or adds
    assert all("replaces" in s["why"] or "no incumbent" in s["why"]
               for s in out["steps"])


def test_unknown_task_refuses():
    with pytest.raises(XlMcpError):
        _workflows.get_workflows("world-domination")
