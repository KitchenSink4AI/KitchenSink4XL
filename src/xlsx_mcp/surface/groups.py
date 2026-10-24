"""Which tool each op lives on: the Excel grouping. Grouped by risk
and domain: one read-only tool for every read, one tool per kind of change,
and xl_app for the ops that drive the private hidden Excel.

Order within a tool is the order the op lines appear in the tool description.
xl_read leads with the cheap orientation reads (metadata, the grid view, a
range read), then the targeted reads, the audits and the rare
inspections; the other tools keep their starting-map order. xl_help is
generated from the registry and holds no ops.
"""

from __future__ import annotations

HELP_TOOL = "xl_help"

GROUPS: dict[str, tuple[str, ...]] = {
    "xl_read": (
        "get_workbook_metadata", "get_grid_view", "read_range", "get_cells",
        "query_range", "find_cells", "get_table", "diagnose_workbook",
        "validate", "audit_formulas", "audit_styles", "get_pivot",
        "get_connections", "get_external_links", "inspect_vba",
        "get_workflows", "get_server_info", "com_status",
    ),
    "xl_edit": (
        "set_cell", "set_cells", "write_range", "clear_range", "copy_range",
        "move_range", "apply_edits", "replace_cells", "set_formula",
        "import_data",
    ),
    "xl_format": (
        "format_cells", "apply_style", "copy_format", "set_dimensions",
        "set_merge", "manage_conditional_format", "manage_comment",
        "manage_hyperlink", "manage_image", "manage_chart", "set_view",
    ),
    "xl_structure": (
        "manage_worksheet", "modify_grid_structure", "create_table",
        "manage_table", "manage_name", "sort_range", "set_filter",
        "clear_filter", "manage_data_validation",
    ),
    "xl_workbook": (
        "create_workbook", "copy_workbook", "manage_backups",
        "set_workbook_properties", "set_protection", "set_page_layout",
        "set_header_footer", "export_range", "export_file",
    ),
    "xl_app": (
        "recalculate", "com_manage_pivot", "com_export_pdf",
        "com_render_sheet", "com_convert_format", "com_save_with_password",
        "com_autofit", "com_goal_seek", "com_set_sparkline",
        "com_validate_opens_clean", "update_now", "update_notices",
    ),
}

READ_ONLY_GROUPS = frozenset({"xl_read"})

TOOL_NAMES: tuple[str, ...] = tuple(GROUPS) + (HELP_TOOL,)

OP_GROUP: dict[str, str] = {op: tool for tool, ops in GROUPS.items() for op in ops}


def group_of(op: str) -> str | None:
    return OP_GROUP.get(op)
