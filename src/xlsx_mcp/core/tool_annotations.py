"""Tool titles, declared by name.

Every registered tool ships a `title` annotation: the Anthropic Connectors
Directory requires one, and a client shows it to a human instead of the
raw tool name. `readOnlyHint` lives in core/readonly.py and is not repeated here.

The title is derived mechanically, so that a new tool gets one without
anyone inventing it: the tool name, split on underscores, Title Cased,
with known acronyms upper-cased, a leading `com_` replaced by the host
application's name, and short function words kept lowercase inside the
phrase. A handful of names too short or too generic to read as a title
take a phrase from the first clause of their own description instead;
those are the only hand-written strings in the table, and
docs/TOOL_TITLES.md marks each one.

Titles are unique within this server and none exceeds 40 characters. An
unknown name RAISES at registration, the same contract core/readonly.py
already holds this surface to.
"""

from __future__ import annotations

#: tool name -> the title a client shows. Unique, at most 40
#: characters, and mirrored into docs/TOOL_TITLES.md, which a test
#: checks against this table so the published English cannot drift
#: from what goes on the wire.
TITLES: dict[str, str] = {
    "apply_edits":               "Apply Edits",
    "apply_style":               "Apply Style",
    "audit_formulas":            "Audit Formulas",
    "audit_styles":              "Audit Styles",
    "clear_filter":              "Clear Filter",
    "clear_range":               "Clear Range",
    "com_autofit":               "Excel Autofit",
    "com_convert_format":        "Excel Convert Format",
    "com_export_pdf":            "Excel Export PDF",
    "com_goal_seek":             "Excel Goal Seek",
    "com_manage_pivot":          "Excel Manage Pivot",
    "com_render_sheet":          "Excel Render Sheet",
    "com_save_with_password":    "Excel Save with Password",
    "com_set_sparkline":         "Excel Set Sparkline",
    "com_status":                "Excel Status",
    "com_validate_opens_clean":  "Excel Validate Opens Clean",
    "copy_format":               "Copy Format",
    "copy_range":                "Copy Range",
    "copy_workbook":             "Copy Workbook",
    "create_table":              "Create Table",
    "create_workbook":           "Create Workbook",
    "diagnose_workbook":         "Diagnose Workbook",
    "disable_tools":             "Disable Tools",
    "enable_tools":              "Enable Tools",
    "export_file":               "Export File",
    "export_range":              "Export Range",
    "find_cells":                "Find Cells",
    "format_cells":              "Format Cells",
    "get_cells":                 "Get Cells",
    "get_connections":           "Get Connections",
    "get_external_links":        "Get External Links",
    "get_grid_view":             "Get Grid View",
    "get_pivot":                 "Get Pivot",
    "get_server_info":           "Get Server Info",
    "get_table":                 "Get Table",
    "get_workbook_metadata":     "Get Workbook Metadata",
    "get_workflows":             "Get Workflows",
    "import_data":               "Import Data",
    "inspect_vba":               "Inspect VBA",
    "manage_backups":            "Manage Backups",
    "manage_chart":              "Manage Chart",
    "manage_comment":            "Manage Comment",
    "manage_conditional_format": "Manage Conditional Format",
    "manage_data_validation":    "Manage Data Validation",
    "manage_hyperlink":          "Manage Hyperlink",
    "manage_image":              "Manage Image",
    "manage_name":               "Manage Name",
    "manage_table":              "Manage Table",
    "manage_worksheet":          "Manage Worksheet",
    "modify_grid_structure":     "Modify Grid Structure",
    "move_range":                "Move Range",
    "query_range":               "Query Range",
    "read_range":                "Read Range",
    "recalculate":               "Recalculate Formulas",
    "replace_cells":             "Replace Cells",
    "set_cell":                  "Set Cell",
    "set_cells":                 "Set Cells",
    "set_dimensions":            "Set Dimensions",
    "set_filter":                "Set Filter",
    "set_formula":               "Set Formula",
    "set_header_footer":         "Set Header Footer",
    "set_merge":                 "Set Merge",
    "set_page_layout":           "Set Page Layout",
    "set_protection":            "Set Protection",
    "set_view":                  "Set View",
    "set_workbook_properties":   "Set Workbook Properties",
    "sort_range":                "Sort Range",
    "validate":                  "Validate Workbook",
    "write_range":               "Write Range",
}


def title(name: str) -> str:
    """The MCP title for one tool. Unknown names RAISE, because a tool that
    reaches tools/list without a title fails the directory's annotation
    requirement and a name-shaped fallback would hide that from the test."""
    try:
        return TITLES[name]
    except KeyError:
        raise RuntimeError(
            f"tool {name!r} has no title in this module. Every registered "
            f"tool needs one: the Anthropic directory requires a title on "
            f"every tool, and a generated fallback would pass the check "
            f"while shipping 'Com Export Pdf' to a user."
        ) from None


def annotations(name: str, read_only: bool) -> dict:
    """The annotation dict for one tool, ready for registration."""
    return {
        "title": title(name),
        "readOnlyHint": read_only,
    }
