"""Tool titles and the write-shape annotations, declared by name.

Every registered tool ships four MCP annotations. `readOnlyHint` lives in
core/readonly.py and is not repeated here. The other three live here:

* **title** -- the human-readable name a client shows instead of the raw
  tool name. Derived mechanically: the tool name, split on underscores,
  Title Cased, with known acronyms upper-cased, a leading `com_` replaced
  by the host application's name, and short function words kept lowercase
  inside the phrase. A handful of names too short or too generic to read
  as a title take a phrase from the first clause of their own description
  instead; those are the only hand-written strings in the table and
  docs/TOOL_TITLES.md marks each one. Titles are unique within this
  server and none exceeds 40 characters.

* **destructiveHint** -- whether the tool may perform a destructive
  update. ONE rule decides it, and docs/TOOL_ANNOTATIONS.md states the
  rule alongside a per-tool reason so a reviewer can dispute any single
  row:

      false only when every code path either (a) adds new content or a new
      file without replacing anything that was already there, creation
      refusing an existing target, or (b) changes no user data at all
      (this session's tool surface, the viewport, a read performed through
      a hidden or already-running Office instance, a read taken through a
      temporary copy).

      true otherwise. That covers everything that deletes, replaces,
      overwrites, clears, reorders, moves, applies a batch of edits, saves
      over the document the user has open, or writes an output file it may
      silently overwrite. It also covers every tool that writes into an
      existing document, because the pre-write backup these servers take
      is defeatable by the tool's own `backup=False` argument and so
      cannot be claimed as guaranteed reversibility, and everything whose
      reversibility could not be PROVEN by reading the code.

  Read-only tools carry no destructiveHint: the field is meaningful only
  when readOnlyHint is false, and a value there would be noise.

* **idempotentHint** -- true only where repeating the identical call
  obviously lands the same state: the pack switches, whose own docstrings
  say idempotent, and the whole-value setters that take an address and a
  value and carry no action selector. Everything else is left unset rather
  than guessed.

* **openWorldHint** -- false on every tool. Nothing here reaches a remote
  service; the optional update check is not a tool.

An unclassified name RAISES at registration, the same contract
core/readonly.py already holds this surface to.
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

#: destructiveHint: true. Deletes, replaces, overwrites, clears,
#: reorders, moves, batch-edits, saves over the open document, writes
#: an output file it may overwrite, acts on a live page, or could not
#: be proven reversible. docs/TOOL_ANNOTATIONS.md carries the reason
#: for every row.
DESTRUCTIVE: frozenset[str] = frozenset({
    "apply_edits", "apply_style", "clear_filter", "clear_range",
    "com_autofit", "com_convert_format", "com_export_pdf",
    "com_goal_seek", "com_manage_pivot", "com_render_sheet",
    "com_save_with_password", "com_set_sparkline", "copy_format",
    "copy_range", "export_file", "export_range", "format_cells",
    "import_data", "manage_backups", "manage_chart", "manage_comment",
    "manage_conditional_format", "manage_data_validation",
    "manage_hyperlink", "manage_image", "manage_name", "manage_table",
    "manage_worksheet", "modify_grid_structure", "move_range",
    "recalculate", "replace_cells", "set_cell", "set_cells",
    "set_dimensions", "set_filter", "set_formula", "set_header_footer",
    "set_merge", "set_page_layout", "set_protection", "set_view",
    "set_workbook_properties", "sort_range", "write_range"
})

#: destructiveHint: false. Every code path either only ADDS, or
#: changes no user data at all. Each row's reason is in
#: docs/TOOL_ANNOTATIONS.md.
NON_DESTRUCTIVE: frozenset[str] = frozenset({
    "com_validate_opens_clean", "copy_workbook", "create_table",
    "create_workbook", "disable_tools", "enable_tools"
})

#: idempotentHint: true. Repeating the identical call lands the same
#: state. Anything absent here is left UNSET rather than guessed.
IDEMPOTENT: frozenset[str] = frozenset({
    "disable_tools", "enable_tools", "set_cell", "set_cells",
    "set_dimensions", "set_filter", "set_formula", "set_header_footer",
    "set_page_layout", "set_view", "set_workbook_properties"
})


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


def destructive_hint(name: str) -> bool | None:
    """The MCP destructiveHint, or None for a read-only tool, where the
    field carries no meaning. An unclassified mutating name returns None
    and `annotations` raises on it."""
    if name in NON_DESTRUCTIVE:
        return False
    if name in DESTRUCTIVE:
        return True
    return None


def idempotent_hint(name: str) -> bool | None:
    """True where repeating the call lands the same state, else None. Never
    false: an unlisted tool is unclassified, not proven non-idempotent."""
    return True if name in IDEMPOTENT else None


def open_world_hint(name: str) -> bool:
    """False for every tool: this server talks to local files and to a local
    Office installation, never to a remote service."""
    return False


def annotations(name: str, read_only: bool) -> dict:
    """The full annotation dict for one tool, ready for registration.
    destructiveHint is omitted on read-only tools and idempotentHint is
    omitted where it was not classified, so an absent field means
    "undeclared" rather than "false"."""
    ann: dict = {
        "title": title(name),
        "readOnlyHint": read_only,
        "openWorldHint": open_world_hint(name),
    }
    if not read_only:
        destructive = destructive_hint(name)
        if destructive is None:
            raise RuntimeError(
                f"tool {name!r} is not classified DESTRUCTIVE or "
                f"NON_DESTRUCTIVE in this module. Every tool that can "
                f"change something must declare whether the change may be "
                f"destructive before it can be registered."
            )
        ann["destructiveHint"] = destructive
    idempotent = idempotent_hint(name)
    if idempotent is not None:
        ann["idempotentHint"] = idempotent
    return ann
