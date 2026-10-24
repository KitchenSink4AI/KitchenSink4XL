"""A worked example call per op, served by xl_help and exercised
by the op sweep (tests/unit/test_op_sweep.py).

Data, not English: every value is an argument an op accepts. The examples
assume a workbook whose active sheet holds a small table in A1:B4 (a header
row Name/Qty and three data rows), a table named Sales over it, a total
formula in B5, a Region/Units block in D1:E3 and a defined name Totals; the
sweep builds exactly that fixture. Every op has an entry (no help example
is a schema placeholder), and tests hold that every example
validates against its op's strict argument model.
"""

from __future__ import annotations

from typing import Any

BOOK = "workbook.xlsx"

EXAMPLES: dict[str, dict[str, Any]] = {
    # xl_read
    "get_workbook_metadata": {"path": BOOK},
    "read_range": {"path": BOOK, "location": "A1:B4", "values": "both"},
    "get_cells": {"path": BOOK, "cells": ["A2", "B2"]},
    "query_range": {"path": BOOK, "location": "A1:B4",
                    "where": [{"column": "Qty", "op": "gt", "value": 1}]},
    "find_cells": {"path": BOOK, "query": "Name"},
    "get_table": {"path": BOOK, "name": "Sales"},
    "get_grid_view": {"path": BOOK},
    "audit_formulas": {"path": BOOK},
    "audit_styles": {"path": BOOK, "top": 5},
    "validate": {"path": BOOK},
    "diagnose_workbook": {"path": BOOK},
    "get_pivot": {"path": BOOK},
    "get_connections": {"path": BOOK},
    "get_external_links": {"path": BOOK},
    "inspect_vba": {"path": BOOK},
    "get_workflows": {"task": "report-build"},
    "get_server_info": {},
    "com_status": {},
    # xl_edit
    "set_cell": {"path": BOOK, "location": "C1", "value": "Total"},
    "set_cells": {"path": BOOK, "cells": [{"cell": "C2", "value": 10}]},
    "write_range": {"path": BOOK, "location": "G1", "data": [["Region", "Units"], ["North", 5]]},
    "clear_range": {"path": BOOK, "location": "D2:E2"},
    "copy_range": {"path": BOOK, "source": "A1:B2", "dest": "G5"},
    "move_range": {"path": BOOK, "source": "A4:B4", "dest": "G10"},
    "apply_edits": {"path": BOOK, "edits": [
        {"op": "set_value", "location": "C3", "value": 3},
        {"op": "set_formula", "location": "C5", "formula": "=SUM(B2:B4)"}]},
    "replace_cells": {"path": BOOK, "find": "Name", "replace": "Item", "dry_run": True},
    "set_formula": {"path": BOOK, "location": "B5", "formula": "=SUM(B2:B4)"},
    "import_data": {"path": BOOK, "source": "Region,Units\nNorth,5\n", "fmt": "csv",
                    "location": "J1"},
    # xl_format
    "format_cells": {"path": BOOK, "location": "A1:B1", "font": {"bold": True},
                     "fill": {"color": "FFFF00"}},
    "apply_style": {"path": BOOK, "style": "Good", "location": "B2"},
    "copy_format": {"path": BOOK, "source": "A1", "dest": "D1"},
    "set_dimensions": {"path": BOOK, "column_widths": {"A": 18}},
    "set_merge": {"path": BOOK, "action": "list"},
    "manage_conditional_format": {"path": BOOK, "action": "add", "location": "B2:B4",
                                  "cf_type": "cell_is",
                                  "params": {"operator": "greater_than", "formula": ["1"],
                                             "fill": "FFC7CE"}},
    "manage_comment": {"path": BOOK, "action": "add", "location": "B3", "text": "Check this"},
    "manage_hyperlink": {"path": BOOK, "action": "add", "location": "A3",
                         "target": "https://example.com"},
    "manage_image": {"path": BOOK, "action": "list"},
    "manage_chart": {"path": BOOK, "action": "create", "chart_type": "bar",
                     "data": "B1:B4", "categories": "A2:A4", "at_cell": "E2"},
    "set_view": {"path": BOOK, "freeze": "A2"},
    # xl_structure
    "manage_worksheet": {"path": BOOK, "action": "add", "new_name": "Notes"},
    "modify_grid_structure": {"path": BOOK, "action": "insert_rows", "at": 3, "count": 1},
    "create_table": {"path": BOOK, "location": "D1:E3", "name": "Regions"},
    "manage_table": {"path": BOOK, "name": "Sales", "action": "add_row",
                     "values": ["d", 4]},
    "manage_name": {"path": BOOK, "action": "add", "name": "QtyRange",
                    "refers_to": "=Sheet1!$B$2:$B$4"},
    "sort_range": {"path": BOOK, "location": "A1:B4",
                   "keys": [{"column": "B", "descending": True}]},
    "set_filter": {"path": BOOK, "location": "A1:B4"},
    "clear_filter": {"path": BOOK},
    "manage_data_validation": {"path": BOOK, "action": "add", "location": "C2:C4",
                               "dv_type": "list", "values": ["Yes", "No"]},
    # xl_workbook
    "create_workbook": {"path": BOOK, "sheets": ["Summary", "Data"]},
    "copy_workbook": {"src": BOOK, "dst": "workbook-copy.xlsx"},
    "manage_backups": {"action": "list", "path": BOOK},
    "set_workbook_properties": {"path": BOOK, "title": "Quarterly figures"},
    "set_protection": {"path": BOOK, "action": "sheet"},
    "set_page_layout": {"path": BOOK, "orientation": "landscape"},
    "set_header_footer": {"path": BOOK, "header": {"center": "Quarterly figures"}},
    "export_range": {"path": BOOK, "location": "A1:B4", "fmt": "csv"},
    "export_file": {"path": BOOK, "fmt": "json"},
    # xl_app
    "recalculate": {"path": BOOK, "engine": "auto"},
    "com_manage_pivot": {"path": BOOK, "action": "create", "source_range": "A1:B4",
                         "dest_cell": "E3", "rows": ["Name"],
                         "values": [{"field": "Qty", "func": "sum"}]},
    "com_export_pdf": {"path": BOOK, "output": "workbook.pdf"},
    "com_render_sheet": {"path": BOOK, "output": "sheet.png"},
    "com_convert_format": {"path": BOOK, "output": "workbook.xlsm"},
    "com_save_with_password": {"path": BOOK, "password": "example-password"},
    "com_autofit": {"path": BOOK, "columns": "A:B"},
    "com_goal_seek": {"path": BOOK, "target_cell": "B5", "target_value": 10,
                      "changing_cell": "B2", "save": False},
    "com_set_sparkline": {"path": BOOK, "action": "create", "location": "C2",
                          "source": "B2:B4"},
    "com_validate_opens_clean": {"path": BOOK},
    "update_now": {},
    "update_notices": {"action": "status"},
}
