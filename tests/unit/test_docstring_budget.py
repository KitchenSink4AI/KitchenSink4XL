"""Docstring budgets and the no-em-dash rule.

Two rules, enforced mechanically:

- NO EM DASHES anywhere in a tool description.
- DOCSTRING BUDGET: every description in 60-130 tokens (chars/4), multiplex
  tools to ~350. Binding across the whole registered surface.

The MULTIPLEX list is an exemption, so it is reviewed like one. It holds
26 of the 69 registered tools, and those 26 carry 54% of the description
budget. Adding a name here
is a decision to spend up to 220 extra tokens on it, every session it is
loaded.

MULTIPLEX pre-lists the planned high-traffic / action-parameter tools
 so later additions inherit the ~350 cap without touching
this file.
"""

from __future__ import annotations

import asyncio

import pytest

from xlsx_mcp import server

# The docstring budget binds across the whole registered surface.
SURFACE_READY = True

# The multiplex / high-traffic tools that earn the ~350-token cap.
MULTIPLEX = {
    "get_grid_view", "apply_edits", "validate", "get_workflows",
    "diagnose_workbook", "manage_worksheet", "manage_backups",
    "manage_table", "manage_name", "manage_comment", "manage_hyperlink",
    "manage_conditional_format", "manage_data_validation",
    "modify_grid_structure", "audit_formulas",
    "com_manage_pivot", "enable_tools", "disable_tools",
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


def _tools():
    """The shipped surface only: tools whose function lives in an xlsx_mcp
    module (staging/integration snippets that re-register on the shared
    FastMCP instance are filtered out)."""
    tools = asyncio.run(server.mcp.list_tools())
    return [
        t for t in tools
        if getattr(getattr(t, "fn", None), "__module__", "").startswith(
            "xlsx_mcp"
        )
    ]


def test_no_em_dashes_in_descriptions():
    """Standing rule: no em dashes anywhere public."""
    for tool in _tools():
        desc = tool.description or ""
        assert "—" not in desc, f"{tool.name} description has an em dash"


def test_docstring_budget():
    """Every tool description inside 60-130 tokens (chars/4; multiplex tools
    to ~350). Skipped when SURFACE_READY is False."""
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
