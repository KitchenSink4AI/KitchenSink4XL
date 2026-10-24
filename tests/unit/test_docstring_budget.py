"""Docstring budgets and the no-em-dash rule, the same as in the Word and PowerPoint tests.

Two rules, enforced mechanically (the Glama Tool-Definition-
Quality-A discipline):

- NO EM DASHES anywhere in a tool description.
- DOCSTRING BUDGET: every description in 60-130 tokens (chars/4), multiplex
  tools to ~350. Binding across the whole registered surface
  (SURFACE_READY below).

The MULTIPLEX list is an exemption, so it is audited like one. It now holds
26 of the 69 registered tools, and those 26 carry 54% of the description
budget; a planning entry for manage_power_query sat in it for months after
the tool was dropped. Adding a name here
is a decision to spend up to 220 extra tokens on it, every session it is
loaded.

MULTIPLEX pre-lists the planned high-traffic / action-parameter tools
so later additions inherit the ~350 cap without touching
this file.
"""

from __future__ import annotations

import inspect

import pytest

from xlsx_mcp import registry, server  # noqa: F401  (the import registers ops)
from xlsx_mcp.settings import Settings
from xlsx_mcp.surface import plan

# The core data-plane families have landed, so the
# docstring budget binds across the whole registered surface.
SURFACE_READY = True

# The planned multiplex / high-traffic tools that earn the ~350-token cap
# (pre-listed so later additions inherit the cap). The
# later additions (query_range, the token-shaped read) join here.
MULTIPLEX = {
    "get_grid_view", "apply_edits", "validate", "get_workflows",
    "diagnose_workbook", "manage_worksheet", "manage_backups",
    "manage_table", "manage_name", "manage_comment", "manage_hyperlink",
    "manage_conditional_format", "manage_data_validation",
    "modify_grid_structure", "audit_formulas",
    "com_manage_pivot",
    "query_range",
    # set_merge is an action-parameter tool (merge/unmerge/list);
    # replace_cells is high-traffic with dry-run, plan-first, and
    # injection-lint semantics that do not fit the standard cap.
    "set_merge", "replace_cells",
    # The objects pair and set_protection are action-parameter
    # multiplexes with honesty stories (round-trip, advisory protection)
    # that must be stated; set_page_layout carries eleven orthogonal
    # settings whose units and exclusions need naming.
    "manage_image", "manage_chart", "set_protection", "set_page_layout",
    # 1.1: set_dimensions crossed the line when grouping landed. It now
    # carries ten orthogonal settings across three families (sizing, hiding,
    # and the outline), and the outline half needs its own vocabulary
    # stated: spans, levels 1-7, the collapse flag, and which side the
    # summary row sits on. That is the same shape as set_page_layout.
    "set_dimensions",
}


class _Op:
    """An op's shipped docstring as xl_help serves it (family surface: the
    per-op description is no longer a wire tool description, but it is
    still what a client reads about the op, and its first sentence is the
    op's line in the tool description)."""

    def __init__(self, name, spec):
        self.name = name
        self.description = inspect.getdoc(spec.handler) or ""


def _tools():
    return [_Op(n, s) for n, s in registry.public_ops().items()]


def test_no_em_dashes_in_descriptions():
    """Standing rule: no em dashes anywhere public.
    Checked on every op docstring and on the seven wire tools."""
    for tool in _tools():
        desc = tool.description or ""
        assert "\u2014" not in desc, f"{tool.name} description has an em dash"
    for name, tool in plan.build(Settings()).tools.items():
        assert "\u2014" not in tool.description + str(tool.input_schema), name


def test_docstring_budget():
    """Every tool description inside 60-130 tokens (chars/4; multiplex tools
    to ~350). Skipped until SURFACE_READY: the full surface does not exist in
    a bare scaffold, so enforcing the lower bound would be a false
    failure. The gate is wired; flip SURFACE_READY to bind it."""
    if not SURFACE_READY:
        pytest.skip(
            "docstring budget gated until the engine surface lands "
            "(flip SURFACE_READY in this file to enforce)"
        )
    for tool in _tools():
        desc = tool.description or ""
        assert desc, f"{tool.name} has no description"
        tokens = len(desc) / 4
        cap = 350 if tool.name in MULTIPLEX else 130
        assert 60 <= tokens <= cap, (
            f"{tool.name} description is ~{tokens:.0f} tokens, "
            f"outside [60, {cap}]"
        )
