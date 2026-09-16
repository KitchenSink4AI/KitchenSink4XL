# Tool annotations

Every tool this server registers ships four MCP annotations: `title`
(see [TOOL_TITLES.md](TOOL_TITLES.md)), `readOnlyHint`, `destructiveHint`
and `openWorldHint`, plus `idempotentHint` where it is obviously true.
The tables are GENERATED from `src/xlsx_mcp/core/tool_annotations.py` and a test in
`tests/unit/test_tool_annotations.py` fails if they disagree, so a row
here is what goes on the wire.

## The rule for `destructiveHint`

ONE rule decides every row, and it is stated here so a reviewer can
dispute any single one of them.

**`false`** only when every code path either

* **(a)** adds new content or a new file without replacing anything that
  was already there, with creation refusing an existing target, or
* **(b)** changes no user data at all: this session's tool surface, the
  viewport, a read performed through a hidden or already-running Office
  instance, or a read taken through a temporary copy.

**`true`** otherwise. That covers everything that deletes, replaces,
overwrites, clears, reorders, moves, applies a batch of edits, saves over
the document the user has open, writes an output file it may silently
overwrite, or acts on a live page. It also covers **every tool that writes
into an existing document**, because the pre-write backup these servers
take is defeatable by the tool's own `backup=False` argument and therefore
cannot be claimed as guaranteed reversibility. And it covers everything
whose reversibility could not be PROVEN by reading the code: an unproven
claim of safety is the one thing this annotation must not make.

Read-only tools carry no `destructiveHint`. The field is meaningful only
when `readOnlyHint` is false, and a value there would be noise.

`idempotentHint` is set true only where repeating the identical call
obviously lands the same state: the pack switches, whose own docstrings
say idempotent, and the whole-value setters that take an address and a
value and carry no action selector. Everything else is left unset rather
than guessed; an absent hint means undeclared, never false.

`openWorldHint` is `false` on every tool. Nothing here reaches a remote service; the
optional update check is not a tool.

## Tools that may perform destructive updates (45)

`destructiveHint: true`.

| Tool | Why |
|---|---|
| `apply_edits` | applies a batch of addressed edits whose ops include replace and delete |
| `apply_style` | writes over the styling or content already in place |
| `clear_filter` | removes the autofilter and unhides the rows it hid |
| `clear_range` | deletes existing content |
| `com_autofit` | rewrites stored column widths and row heights, then saves |
| `com_convert_format` | writes a converted file with no proven existing-file check |
| `com_export_pdf` | overwrite=true writes over an existing PDF |
| `com_goal_seek` | writes the solved value into the changing cell and saves |
| `com_manage_pivot` | action='delete' removes a pivot table |
| `com_render_sheet` | overwrite=true writes over an existing PNG |
| `com_save_with_password` | encrypts the workbook with a password that cannot be recovered |
| `com_set_sparkline` | action='clear' removes sparkline groups |
| `copy_format` | the destination's existing formatting is overwritten |
| `copy_range` | the destination rectangle is overwritten |
| `export_file` | an overwrite argument lets it write over an existing output file |
| `export_range` | an overwrite argument lets it write over an existing output file |
| `format_cells` | writes a new value over the one already stored |
| `import_data` | writes over whatever already sits under the anchor |
| `manage_backups` | action='restore' overwrites the workbook and action='purge' deletes backups |
| `manage_chart` | carries delete or remove actions alongside its read actions |
| `manage_comment` | carries delete or remove actions alongside its read actions |
| `manage_conditional_format` | carries delete or remove actions alongside its read actions |
| `manage_data_validation` | carries delete or remove actions alongside its read actions |
| `manage_hyperlink` | carries delete or remove actions alongside its read actions |
| `manage_image` | carries delete or remove actions alongside its read actions |
| `manage_name` | carries delete or remove actions alongside its read actions |
| `manage_table` | carries delete or remove actions alongside its read actions |
| `manage_worksheet` | carries delete or remove actions alongside its read actions |
| `modify_grid_structure` | deletes rows or columns and rewrites every reference to them |
| `move_range` | reorders or moves existing content |
| `recalculate` | opens the workbook in Excel, recalculates, and saves over it |
| `replace_cells` | rewrites matched cell values or formula text workbook-wide |
| `set_cell` | writes a new value over the one already stored |
| `set_cells` | writes a new value over the one already stored |
| `set_dimensions` | writes a new value over the one already stored |
| `set_filter` | hides the non-matching rows it evaluates |
| `set_formula` | writes a new value over the one already stored |
| `set_header_footer` | writes a new value over the one already stored |
| `set_merge` | a merge keeps only the top-left value; the absorbed cells are cleared |
| `set_page_layout` | writes a new value over the one already stored |
| `set_protection` | writes a new value over the one already stored |
| `set_view` | writes a new value over the one already stored |
| `set_workbook_properties` | writes a new value over the one already stored |
| `sort_range` | reorders or moves existing content |
| `write_range` | writes a block of values over the cells under the anchor |
## Tools that may not (6)

`destructiveHint: false`.

| Tool | Why |
|---|---|
| `com_validate_opens_clean` | drives a hidden or already-running Office instance to read; the file is not modified |
| `copy_workbook` | writes to a separate destination; an overwrite rotates the destination into its backup slot first |
| `create_table` | adds new content at a position; nothing existing is replaced or removed |
| `create_workbook` | creates a new object or file and refuses an existing target |
| `disable_tools` | changes only this session's tool surface, which the counterpart tool reverses |
| `enable_tools` | changes only this session's tool surface, which the counterpart tool reverses |
## Read-only tools (18)

`readOnlyHint: true`, and no `destructiveHint`: the field carries no
meaning for a tool that changes nothing. The classification itself lives
in `core/readonly.py` and is guarded by
`tests/unit/test_readonly_annotations.py`.

`audit_formulas`, `audit_styles`, `com_status`, `diagnose_workbook`, `find_cells`, `get_cells`, `get_connections`, `get_external_links`, `get_grid_view`, `get_pivot`, `get_server_info`, `get_table`, `get_workbook_metadata`, `get_workflows`, `inspect_vba`, `query_range`, `read_range`, `validate`

## Idempotent tools (11)

`idempotentHint: true`.

`disable_tools`, `enable_tools`, `set_cell`, `set_cells`, `set_dimensions`, `set_filter`, `set_formula`, `set_header_footer`, `set_page_layout`, `set_view`, `set_workbook_properties`
