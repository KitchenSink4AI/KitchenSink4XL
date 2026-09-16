# Tool titles

Every tool this server registers ships a `title` annotation, which is what
the Anthropic Connectors Directory requires and what a client shows in
place of the raw tool name. This table is GENERATED from
`src/xlsx_mcp/core/tool_annotations.py` and a test in `tests/unit/test_tool_annotations.py`
fails if the two ever disagree, so the English here is the English on the
wire.

## How a title is derived

The rule is mechanical, so that a new tool gets a title without anyone
inventing one:

1. split the tool name on underscores;
2. replace a leading `com_` with the host application's name and collapse
   an immediately repeated word (`com_word_status` -> `Word Status`);
3. upper-case known acronyms (`pdf` -> `PDF`, `svg` -> `SVG`), Title Case
   the rest, and keep short function words lowercase inside the phrase
   (`add_equation_to_shape` -> `Add Equation to Shape`);
4. a name too short or too generic to read as a title takes a phrase from
   the first clause of its own description instead. Those are the only
   hand-written strings here and the Source column marks them.

Titles are unique within this server and none exceeds 40 characters. No
title carries product or marketing language.

69 tools.

| Tool | Title | Source |
|---|---|---|
| `apply_edits` | Apply Edits | mechanical |
| `apply_style` | Apply Style | mechanical |
| `audit_formulas` | Audit Formulas | mechanical |
| `audit_styles` | Audit Styles | mechanical |
| `clear_filter` | Clear Filter | mechanical |
| `clear_range` | Clear Range | mechanical |
| `com_autofit` | Excel Autofit | mechanical |
| `com_convert_format` | Excel Convert Format | mechanical |
| `com_export_pdf` | Excel Export PDF | mechanical |
| `com_goal_seek` | Excel Goal Seek | mechanical |
| `com_manage_pivot` | Excel Manage Pivot | mechanical |
| `com_render_sheet` | Excel Render Sheet | mechanical |
| `com_save_with_password` | Excel Save with Password | mechanical |
| `com_set_sparkline` | Excel Set Sparkline | mechanical |
| `com_status` | Excel Status | mechanical |
| `com_validate_opens_clean` | Excel Validate Opens Clean | mechanical |
| `copy_format` | Copy Format | mechanical |
| `copy_range` | Copy Range | mechanical |
| `copy_workbook` | Copy Workbook | mechanical |
| `create_table` | Create Table | mechanical |
| `create_workbook` | Create Workbook | mechanical |
| `diagnose_workbook` | Diagnose Workbook | mechanical |
| `disable_tools` | Disable Tools | mechanical |
| `enable_tools` | Enable Tools | mechanical |
| `export_file` | Export File | mechanical |
| `export_range` | Export Range | mechanical |
| `find_cells` | Find Cells | mechanical |
| `format_cells` | Format Cells | mechanical |
| `get_cells` | Get Cells | mechanical |
| `get_connections` | Get Connections | mechanical |
| `get_external_links` | Get External Links | mechanical |
| `get_grid_view` | Get Grid View | mechanical |
| `get_pivot` | Get Pivot | mechanical |
| `get_server_info` | Get Server Info | mechanical |
| `get_table` | Get Table | mechanical |
| `get_workbook_metadata` | Get Workbook Metadata | mechanical |
| `get_workflows` | Get Workflows | mechanical |
| `import_data` | Import Data | mechanical |
| `inspect_vba` | Inspect VBA | mechanical |
| `manage_backups` | Manage Backups | mechanical |
| `manage_chart` | Manage Chart | mechanical |
| `manage_comment` | Manage Comment | mechanical |
| `manage_conditional_format` | Manage Conditional Format | mechanical |
| `manage_data_validation` | Manage Data Validation | mechanical |
| `manage_hyperlink` | Manage Hyperlink | mechanical |
| `manage_image` | Manage Image | mechanical |
| `manage_name` | Manage Name | mechanical |
| `manage_table` | Manage Table | mechanical |
| `manage_worksheet` | Manage Worksheet | mechanical |
| `modify_grid_structure` | Modify Grid Structure | mechanical |
| `move_range` | Move Range | mechanical |
| `query_range` | Query Range | mechanical |
| `read_range` | Read Range | mechanical |
| `recalculate` | Recalculate Formulas | first clause of its description |
| `replace_cells` | Replace Cells | mechanical |
| `set_cell` | Set Cell | mechanical |
| `set_cells` | Set Cells | mechanical |
| `set_dimensions` | Set Dimensions | mechanical |
| `set_filter` | Set Filter | mechanical |
| `set_formula` | Set Formula | mechanical |
| `set_header_footer` | Set Header Footer | mechanical |
| `set_merge` | Set Merge | mechanical |
| `set_page_layout` | Set Page Layout | mechanical |
| `set_protection` | Set Protection | mechanical |
| `set_view` | Set View | mechanical |
| `set_workbook_properties` | Set Workbook Properties | mechanical |
| `sort_range` | Sort Range | mechanical |
| `validate` | Validate Workbook | first clause of its description |
| `write_range` | Write Range | mechanical |
