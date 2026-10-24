"""Scaffold gates: the package imports, the server lists cleanly over
an in-process fastmcp Client, and the shared infrastructure is wired.

These are the scaffold gate checks captured as tests: the server import +
tool-count check and the envelope wiring. Family surface (2.0): the
pack registry and the startup-mode route are gone; the op-to-tool map they
used to pin is pinned here against surface/groups.py, and the surface itself
is covered by test_family_surface.py.
"""

from __future__ import annotations

import asyncio

from fastmcp import Client

from xlsx_mcp import envelope, registry, server
from xlsx_mcp.core.errors import (
    CalcUnavailable,
    HazardRefused,
    WorkbookLocked,
)
from xlsx_mcp.surface.groups import OP_GROUP


# The core data plane (the daily driver).
PHASE_3A = {
    "create_workbook", "copy_workbook", "get_workbook_metadata",
    "diagnose_workbook", "manage_worksheet",
    "read_range", "set_cell", "write_range", "clear_range",
    "copy_range", "move_range", "query_range",
    "get_grid_view", "apply_edits",
    "format_cells", "set_dimensions",
}


def test_imports_and_lists_over_client():
    """The server imports and lists the fixed family tools via an
    in-process fastmcp Client; every core data-plane op is on the surface."""

    async def run():
        async with Client(server.create_server()) as c:
            tools = await c.list_tools()
            return {t.name: t.inputSchema["properties"].get("op", {}).get("enum", [])
                    for t in tools}

    listed = asyncio.run(run())
    assert "xl_help" in listed and "get_server_info" in listed["xl_read"]
    ops = {op for members in listed.values() for op in members}
    assert PHASE_3A <= ops


def test_every_op_has_a_tool():
    """Every registered op sits on one family tool (no strays); the former
    design/io/com pack members land on the family tool that owns them."""
    ops = set(registry.public_ops())
    assert ops == set(OP_GROUP), sorted(ops ^ set(OP_GROUP))
    design = {"manage_conditional_format", "manage_data_validation",
              "apply_style", "copy_format", "audit_styles",
              "manage_image", "manage_chart", "manage_table", "manage_name"}
    io = {"manage_comment", "set_protection", "set_page_layout",
          "set_header_footer", "get_external_links", "inspect_vba",
          "export_file", "get_pivot", "get_connections"}
    com = {"recalculate", "com_manage_pivot", "com_export_pdf",
           "com_render_sheet", "com_convert_format", "com_save_with_password",
           "com_autofit", "com_goal_seek", "com_set_sparkline",
           "com_validate_opens_clean", "com_status"}
    assert (design | io | com) <= ops
    # com_status is a read; every other former com-pack member drives the
    # private hidden Excel from xl_app.
    assert {OP_GROUP[n] for n in com - {"com_status"}} == {"xl_app"}
    assert OP_GROUP["com_status"] == "xl_read"
    # com_run_macro is DELIBERATELY absent: macro execution is deferred from
    # v1 (VBA is preserve/inspect only).
    assert "com_run_macro" not in ops


def test_grid_error_codes_present():
    """The grid-specific closed codes are in the vocabulary and map from
    their exception types."""
    for code in ("HAZARD_REFUSED", "FORMULA_REJECTED", "CALC_UNAVAILABLE",
                 "WORKBOOK_LOCKED"):
        assert code in envelope.CLOSED_CODES
    assert envelope.classify(HazardRefused("x")) == "HAZARD_REFUSED"
    assert envelope.classify(CalcUnavailable("x")) == "CALC_UNAVAILABLE"
    assert envelope.classify(WorkbookLocked("x")) == "WORKBOOK_LOCKED"


def test_refusal_shape_and_iserror():
    """A typed exception becomes the {ok: false, error: {...}} refusal, sent
    once with isError=true."""
    payload = envelope.refusal(HazardRefused("slicers would be lost"))
    assert payload["ok"] is False
    assert payload["error"]["code"] == "HAZARD_REFUSED"
    assert payload["error"]["hint"]  # a HAZARD_REFUSED hint is populated
    r = envelope.tool_result(payload, is_error=True)
    assert r.is_error is True and r.structured_content is None
