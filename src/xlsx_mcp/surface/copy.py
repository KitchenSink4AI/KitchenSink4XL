"""Customer-facing copy: every string a model or a person reads on the MCP
surface that the family migration adds or moves.

Written from the code and the op docstrings, in the same register as the
Word family copy (the wording the Word family measured, adapted to Excel's
facts).

Strings that are NOT here are the product's own shipped English, carried
over unchanged: the refusal hints that existed before (envelope HINTS) and
each op's line, which is the first sentence of the op's own docstring
unless OP_LINES below overrides it.

The owner-test ship gate (scripts/build_editions.py --owner-test and
scripts/smoke_editions.py --owner-test) fails on any unfilled slot, so
`_slot` stays defined for a later surface to use and no value below calls
it.
"""

from __future__ import annotations

#: Prefix every unfilled slot carries. Written as two literals so this
#: definition is not itself a hit once the slots are filled.
SLOT = "COPY" "_SLOT:"
FACTS = "FACTS" " TO CONVEY:"


def _slot(facts: str) -> str:
    return f"{SLOT} {FACTS} {facts}"


# ------------------------------------------------------------ instructions

#: The server instructions. The licensing adapter puts its introduction
#: (product name, PyPI name, suite) in front of this text and its licensing
#: guidance after it, so neither is repeated here. Kept short enough that
#: the whole assembled text stays under Claude Code's 2,048-character cap on
#: server instructions (adapter introduction 75 + this + guidance 941; a unit
#: test pins the assembled length).
#: test_typed_schemas requires every location selector to be named here, and
#: test_family_surface forbids the letters "pack" (so no "package").
INSTRUCTIONS = (
    "Edits Excel .xlsx and .xlsm workbooks. Each tool takes {op, args}; args holds the op's "
    "arguments, including path (the workbook file). xl_help with an op returns its arguments "
    "and an example; xl_help with topic='location' explains the location object: one selector "
    "(cell, range, r1c1, name, table, used_range, region, search or anchor), an optional "
    "sibling sheet, or a bare 'B7' or 'A1:C10' as shorthand. A search or name that matches "
    "more than once is refused with every candidate; if a value or name may be on several "
    "sheets, run find_cells first and ask the user. Reads label each value cached, absent or "
    "formula, so a formula with no stored result never reads as blank. Every change backs the "
    "file up first (in a hidden .ks4xl-backups folder), saves atomically and verifies after "
    "writing. A workbook open in Excel can be read but not changed. xl_app works in a private "
    "hidden Excel, never the user's own. Launch settings may make this server read-only or "
    "remove ops; no tool call changes that. Not affiliated with Microsoft."
)

# ----------------------------------------------------------- the tool list

#: One entry per tool: title, and the description head (when to use it,
#: naming neighbouring tools; what it does to files). The per-op lines are
#: appended by the surface. xl_read's head is kept short: its op lines
#: already overflow into the op parameter, and a longer head grows its
#: schema toward the 5,000-byte limit.
TOOLS: dict[str, dict[str, str]] = {
    "xl_read": {
        "title": "Read Excel Workbook",
        "description": (
            "Read-only: look at a workbook before changing it (grid view, cells, queries, "
            "find, tables, audits, checks). Never changes a file, and works while it is open "
            "in Excel. To change anything use xl_edit, xl_format, xl_structure, xl_workbook "
            "or xl_app."),
    },
    "xl_edit": {
        "title": "Edit Excel Values and Formulas",
        "description": (
            "Values and formulas: set one or many cells, write, clear, copy or move ranges, "
            "the batch editor (apply_edits), find and replace, and CSV, TSV or JSON import. A "
            "string starting with '=' is always stored as a formula. Formatting is xl_format; "
            "sheets, rows and tables are xl_structure. Every change backs the file up first, "
            "saves atomically and verifies after writing. Refuses a workbook open in Excel, "
            "and a workbook whose fragile parts a save would drop unless allow_loss is true."),
    },
    "xl_format": {
        "title": "Excel Formatting and Objects",
        "description": (
            "How cells look and what sits on the sheet: number formats, fonts, fills, borders, "
            "named styles, the format painter, column widths and row heights, merges, "
            "conditional formats, comments, hyperlinks, images, charts and sheet view. Values "
            "are xl_edit. Every change backs the file up first; refuses a workbook open in "
            "Excel."),
    },
    "xl_structure": {
        "title": "Excel Sheets, Rows and Tables",
        "description": (
            "Workbook structure: add, rename, delete or hide sheets, insert or delete rows and "
            "columns (references are rewritten), tables, named ranges, sort, filters and data "
            "validation. Values are xl_edit. Every change backs the file up first; refuses a "
            "workbook open in Excel."),
    },
    "xl_workbook": {
        "title": "Excel Workbook Files",
        "description": (
            "Whole files: create, copy, backups, document properties, protection, page layout, "
            "headers and footers, and export of a range or sheets to CSV, TSV or JSON. Exports "
            "never overwrite the workbook or its backups. Edits inside a sheet belong to the "
            "other tools; manage_backups lists, restores and snapshots the automatic backups."),
    },
    "xl_app": {
        "title": "Excel Application (Hidden Instance)",
        "description": (
            "Needs Microsoft Excel on Windows: full recalculation, real pivot tables, PDF "
            "export, sheet render to PNG, format conversion, password encryption, exact "
            "autofit, goal seek, sparklines and the opens-clean check. Reads go through "
            "xl_read, and cell, format and structure edits through xl_edit, xl_format and "
            "xl_structure. Ops that change the workbook back it up first."),
    },
    "xl_help": {
        "title": "Excel Tool Help",
        "description": (
            "Read-only help. Call it when an op's arguments are unclear or after a BAD_PARAMS "
            "refusal. With no arguments it lists every tool and its ops; with op it returns "
            "that op's full arguments, an example call and the tool that runs it; "
            "topic='location' explains the location object. Reads no file."),
    },
}

#: xl_app head addition (Excel has no live editing).
XL_APP_HIDDEN_EXCEL = (
    "These ops run in a private hidden Excel that this server starts and closes. They never "
    "attach to or change the user's own Excel, and a workbook the user has open in Excel is "
    "refused, not edited.")

#: Per-op line suffixes (facts added to a docstring-derived line).
OP_LINE_SUFFIX: dict[str, str] = {}

#: The one-line schema description of the `op` parameter on every group tool.
OP_PARAM = "The operation to run: one of the ops listed in this tool's description."

#: The one-line schema description of the `args` parameter.
ARGS_PARAM = (
    "The op's arguments as an object. Include path (the workbook file) unless the op's line "
    "names other paths or the op reads no workbook; unknown keys are refused. xl_help lists "
    "every argument.")

#: xl_help's parameter descriptions.
HELP_OP_PARAM = ("An op name, for that op's arguments and an example; leave it out to list "
                 "every tool and op.")
HELP_TOPIC_PARAM = "'location' explains the location object that positional ops take."

#: The heading line above the per-op list in each tool description.
OPS_HEADING = (
    "Ops, one per line, as name(required; opt optional): what it does. path, backup and "
    "verify_com are not repeated; xl_help shows them:")

#: Layout: the description ends with OPS_CONTINUED when some op lines
#: moved to the op parameter, whose description starts with OP_PARAM_CONTINUED.
OPS_CONTINUED = "More ops follow, in the same format, in the op parameter's description."
OP_PARAM_CONTINUED = (
    "The operation to run: an op listed in this tool's description, or one of these, in the "
    "same format:")

#: Per-op line overrides (target at most about 110 characters, verb first,
#: with the one surprise a caller needs). An op not listed keeps its own
#: docstring first sentence verbatim. Each override is written from the op's
#: full docstring.
OP_LINES: dict[str, str] = {
    # xl_read
    "audit_formulas": (
        "Audit formulas in a range, sheet or workbook: external references, volatile "
        "functions, missing cached values, errors."),
    "audit_styles": (
        "Count the workbook's cell formats against Excel's 64,000 limit and rate the bloat risk "
        "(ok, elevated or critical)."),
    "com_status": (
        "Report whether Excel is available to this server and what its hidden Excel is doing; "
        "never starts Excel."),
    "diagnose_workbook": (
        "Scan for fragile parts (pivots, slicers, shapes, macros, queries) a file save would "
        "drop, and recommend the route."),
    "find_cells": (
        "Find every cell whose value or formula matches (exact, contains or regex), with its "
        "address; pages with limit and offset."),
    "get_connections": (
        "List data connections and Power Query presence; connection strings can hold "
        "credentials, so treat them as sensitive."),
    "get_external_links": (
        "List links to other workbooks with their cached values, which are copies of other "
        "files' data, so treat them as sensitive."),
    "get_grid_view": (
        "Show a sheet or range as a compact A1 table with formula and merge markers, plus an "
        "anchor token for later edits."),
    "get_pivot": (
        "Describe existing pivot tables (source, fields, last refresh); creating or refreshing "
        "one is xl_app com_manage_pivot."),
    "get_server_info": (
        "Report version, edition, license status, tool and op counts, launch-settings narrowing "
        "and platform; needs no workbook."),
    "get_workbook_metadata": (
        "Read sheets, true used ranges, names, tables and a hazard summary; the call to orient "
        "before editing."),
    "inspect_vba": (
        "Report whether the workbook carries VBA, with module names and sizes; this server "
        "never writes or runs macros."),
    "query_range": (
        "Filter, sort, page and aggregate a range on the server, so only the rows and columns "
        "needed come back."),
    "read_range": (
        "Read a cell or range by location; values='displayed' gives the text Excel shows, "
        "'both' labels each value."),
    "validate": (
        "Run read-only checks (default structure, references, calc_staleness) and "
        "return one report."),
    # xl_edit
    "import_data": (
        "Import CSV, TSV or JSON text or a file at a location; codes such as IDs and zip codes "
        "stay text, and the result says which columns."),
    "set_cell": (
        "Write one cell by location; a string starting with '=' is always stored as a "
        "formula."),
    "write_range": (
        "Write a rectangular 2D block of values and formulas from the location's top-left cell; "
        "ragged rows are refused."),
    "move_range": (
        "Move a rectangle on the same sheet, rewriting every reference into it so references "
        "follow the cells."),
    "replace_cells": (
        "Find and replace across a workbook, sheet or range (exact, contains or regex); "
        "dry_run=true changes nothing."),
    "set_formula": (
        "Write a formula to a cell, or fill a range with relative references shifted per cell; "
        "results stay stale until recalculated."),
    # xl_format
    "copy_format": (
        "Copy one cell's complete format onto every cell of a destination range, leaving values "
        "untouched."),
    "set_dimensions": (
        "Set column widths and row heights, hide or group rows and columns; autofit here is an "
        "estimate (xl_app com_autofit is exact)."),
    "manage_conditional_format": (
        "Add, list or delete conditional-formatting rules (cell_is, formula, color scales, data "
        "bars, icon sets and more)."),
    "manage_hyperlink": (
        "Add, remove or list cell hyperlinks (a URL or an in-workbook reference); remove keeps "
        "the cell's text."),
    "manage_image": "Insert, list, delete or extract cell-anchored images.",
    "manage_comment": (
        "Add, edit, delete or list cell notes; list also reads threaded comments. Leave author "
        "out to sign with the user's own name, or pass author=\"assistant\" to sign as "
        "\"AI Assistant via KS4XL\"."),
    "manage_chart": (
        "Create, list or delete charts (bar, bar_horizontal, line, pie, doughnut, area, "
        "scatter); at_cell places one."),
    "set_view": (
        "Set freeze or split panes, gridlines, headings, zoom, selection and tab color for a "
        "sheet; unset values stay."),
    # xl_structure
    "manage_worksheet": (
        "Add, delete, rename, copy, reorder, hide or unhide a sheet; rename rewrites references "
        "to it and delete turns them into #REF!, as Excel does."),
    "modify_grid_structure": (
        "Insert or delete rows or columns and rewrite every reference on every sheet; "
        "references inside a deleted band become #REF!."),
    "manage_table": (
        "Add or delete rows and columns, rename, resize, toggle totals, restyle or unlist the "
        "table named name (action)."),
    "manage_name": (
        "Add, delete, rename, update or list defined names; an ambiguous name refuses until "
        "scope is given."),
    "set_filter": (
        "Apply an autofilter and hide the rows that do not meet all criteria; rows with "
        "uncalculated formulas stay visible."),
    "manage_data_validation": (
        "Add, list or delete data-validation rules (list, whole, decimal, date, time, "
        "text_length, custom); Excel enforces them."),
    # xl_workbook
    "copy_workbook": (
        "Copy a workbook byte for byte from src to dst; an existing dst is replaced only with "
        "overwrite=true, after a backup."),
    "manage_backups": (
        "List, restore, purge or snapshot the automatic backups (prev and anchor slots, and "
        "the copies deletes and clears keep); purge is a dry run unless dry_run=false."),
    "set_workbook_properties": (
        "Set document properties and calculation settings (mode, full calc on load, iteration); "
        "with none it reports them."),
    "set_protection": (
        "Protect or unprotect sheets and workbook structure; advisory only, not encryption "
        "(xl_app com_save_with_password encrypts)."),
    "export_range": (
        "Export a range, table or sheet to CSV, TSV or JSON, inline or to a file; "
        "values='displayed' writes what the cells show."),
    "export_file": (
        "Export whole sheets to CSV or TSV files, one per sheet, or to one JSON bundle."),
    "set_page_layout": (
        "Set print settings: orientation, paper size, margins, scale or fit, print area and "
        "repeated titles; unset values stay."),
    # xl_app
    "recalculate": (
        "Recalculate every formula in a hidden Excel and save the results; a circular "
        "reference is named."),
    "com_manage_pivot": "Create, refresh, delete or list real pivot tables through Excel.",
    "com_render_sheet": (
        "Render a sheet or range to a PNG exactly as Excel shows it; the workbook is not "
        "changed."),
    "com_convert_format": (
        "Convert the workbook to xlsx, xlsm, xlsb, xls, csv or ods with Excel's own Save As; "
        "the source is never overwritten."),
    "com_save_with_password": (
        "Encrypt the workbook with Excel's real password encryption; the password cannot be "
        "recovered, so keep the backup."),
    "com_validate_opens_clean": (
        "Open the file in a hidden Excel and report whether it opens without a repair "
        "prompt."),
}

#: The usage line xl_help returns when called with no op.
HELP_USAGE = (
    "Call a tool as {op, args}. xl_help with op gives that op's arguments and an example; "
    "topic='location' explains the location object.")

#: xl_help(topic='location'): the location-object explainer. Facts from
#: core/locate.py and the released server instructions.
LOCATION_HELP = (
    "A location object has exactly one selector, plus an optional sibling \"sheet\" (default "
    "the active sheet). Selectors: {\"cell\": \"B7\"}, {\"range\": \"A1:C10\"}, {\"r1c1\": "
    "\"R7C2\"}, {\"name\": \"Totals\"} (a defined name), {\"table\": \"Sales\"} (optionally "
    "with \"column\"), {\"used_range\": true} (the true used range of the sheet), "
    "{\"region\": {\"near\": \"B7\"}} (the block of data around a cell), {\"search\": "
    "{\"text\": \"Revenue\", \"occurrence\": 1}}, and {\"anchor\": \"...\"} (the token from "
    "xl_read get_grid_view, refused as STALE_ANCHOR if that area changed). A bare \"B7\" or "
    "\"A1:C10\" string is shorthand for a cell or range, for example {\"cell\": \"B7\", "
    "\"sheet\": \"Q3\"}. A search or name that matches more than once is refused with every "
    "candidate; add occurrence or use the address.")

# ------------------------------------------------------------------ hints

#: Recovery lines for the refusal codes the surface adds or newly reaches,
#: and for the codes that had none before (the hint table is total).
#: envelope.HINTS keeps the shipped hints for the older codes. No hint may
#: contain the letters "pack" (test_family_surface).
NEW_HINTS: dict[str, str] = {
    "NOT_AVAILABLE": (
        "This server's launch settings turned this off, usually by IT; no tool call can turn "
        "it back on."),
    # Only the call deadline raises TIMEOUT, and a file save refuses once the
    # deadline is cancelled. Work handed to Excel (the xl_app ops) is
    # not under that commit check, so the hint does not promise more.
    "TIMEOUT": (
        "The call ran past its time limit and was stopped. A file edit that times out saves "
        "nothing; work handed to Excel may still finish, so check the workbook before retrying."),
    "INTERNAL_ERROR": (
        "Something inside the server failed; the arguments were not the cause. Try once more, "
        "and if it repeats, report it with the version from xl_read get_server_info."),
    "LIMIT_EXCEEDED": (
        "The answer would pass the output limit; narrow the request (a smaller range, "
        "xl_read query_range with a limit, or fewer columns)."),
    "BAD_PARAMS": (
        "The request did not fit this op; the message says what. xl_help with this op gives "
        "its arguments and an example; when the op belongs to another tool, details.tool names "
        "it."),
    "UNSUPPORTED_CONTENT": (
        "The workbook holds something this op does not handle, or is not a usable .xlsx or "
        ".xlsm; the message names it, and the file was not changed."),
    "CONFLICT": (
        "The target already exists (pass overwrite where the op offers it), or Excel "
        "disconnected during the call; the message says which. Check it and retry."),
    # The Word editor's codes, in its wording adapted to Excel.
    "PROTECTED": (
        "The sheet's or the workbook's protection does not allow this change; cells that are "
        "not locked can still be edited. If you may lift it, xl_workbook set_protection with "
        "action='remove' does, and nothing else changes."),
    "GUARD_TRIPPED": (
        "A safety check you passed did not hold, so nothing was changed. Read the workbook "
        "again and resend with values that match it."),
}

#: When an argument fails validation, the BAD_PARAMS message gives the
#: form that works instead of only naming the field. Keyed (op, argument); the
#: op "*" stands for any op. {op} is the op's name. These are shown only on a
#: refusal, so they cost the tool list nothing.
WRONG_TYPE_FORMS: dict[tuple[str, str], str] = {
    ("set_cells", "cells"): (
        "{op}: cells must be a list of {{cell, value}} objects, such as "
        "[{{\"cell\": \"A1\", \"value\": 5}}, {{\"cell\": \"B2\", \"value\": "
        "\"=A1*2\"}}]. An object keyed by cell address is not accepted."),
    ("set_filter", "criteria"): (
        "{op}: criteria must be a list of {{column, op, value}} objects, such as "
        "[{{\"column\": \"Region\", \"op\": \"eq\", \"value\": \"West\"}}]. An "
        "object keyed by column name is not accepted."),
}

#: The same, for an argument that was left out.
MISSING_FORMS: dict[tuple[str, str], str] = {
    ("read_range", "location"): (
        "{op} needs location: where to read, such as \"A1:D5\", or "
        "{{\"used_range\": true}} for the whole sheet. The sheet goes in sheet."),
    ("*", "location"): (
        "{op} needs location: a cell or range such as \"B7\" or \"A1:C10\", or an "
        "object like {{\"table\": \"Sales\"}}. The sheet goes in sheet."),
}


def bad_params_message(op: str, errors: list[dict]) -> str | None:
    """The message for a validation failure when one of its errors has a form
    above; None when none does (the caller keeps its generic message)."""
    for e in errors:
        loc = e.get("loc") or ()
        field = str(loc[0]) if loc else ""
        table = MISSING_FORMS if e.get("type") == "missing" else WRONG_TYPE_FORMS
        form = table.get((op, field)) or table.get(("*", field))
        if form:
            return form.format(op=op)
    return None


#: The hint on NOT_FOUND for a workbook path that does not exist: the
#: workbook hint (re-read the metadata) cannot help there.
MISSING_FILE_HINT = (
    "No workbook exists at that path. Check the path, or make a new one with xl_workbook "
    "create_workbook.")

#: A valid range over a write, scan or read ceiling.
CEILING_HINT = (
    "The range is valid but holds more cells than one call takes; the message names the "
    "limit and how to split the work.")

#: Hint for a call to a tool name this server does not list (the refusal's
#: details carry the tool names that do exist; envelope.unknown_tool_middleware).
UNKNOWN_TOOL_HINT = (
    "No tool has that name on this server; details.tools lists the ones that do. Operations "
    "are not tools: call the tool with {op, args}, and xl_help lists every tool and op.")

#: The hint on NOT_FOUND for an op name that does not exist: the workbook
#: hint cannot help there; xl_help lists every op.
UNKNOWN_OP_HINT = (
    "No op has that name. xl_help with no arguments lists every tool and its ops.")

#: The `decided_by` text for a forbidden or read-only refusal and in
#: get_server_info.
SETTINGS_OWNER = (
    "This server's launch settings, usually set by IT; they are fixed until the server "
    "restarts, and no tool call changes them.")

#: The one stderr line at a successful start when retired 1.x settings are set
#: to an old default (settings.py). {names} is the ignored names joined by
#: ", "; no value is ever echoed.
LEGACY_SETTING_NOTICE = (
    "Ignored old settings: {names}. This version lists every tool from the start, and the "
    "server started normally. To limit the tools, use KS4XL_READ_ONLY=1 or "
    "KS4XL_FORBID=<tool or op names>. Remove these settings to stop this notice.")

#: The one stderr line when a retired 1.x setting is set to a value that
#: limited the tools; the server then does not start (it fails closed;
#: settings.py). {name} is the variable's name, and the launcher adds "The
#: server did not start." after it.
LEGACY_SETTING_REFUSAL = (
    "{name} is set to a value that limited the tools in an earlier version. This version "
    "does not read it, so the server stops here instead of starting without that limit. "
    "Remove {name} and use KS4XL_READ_ONLY=1 (reading tools only) or KS4XL_FORBID=<tool or "
    "op names> instead.")


def slots_in(text: str) -> bool:
    return SLOT in text


def all_strings() -> list[str]:
    """Every copy string, for the ship gate and the slot test."""
    out = [INSTRUCTIONS, XL_APP_HIDDEN_EXCEL, OP_PARAM, ARGS_PARAM, HELP_OP_PARAM,
           HELP_TOPIC_PARAM, OPS_HEADING, OPS_CONTINUED, OP_PARAM_CONTINUED, HELP_USAGE,
           LOCATION_HELP, SETTINGS_OWNER, LEGACY_SETTING_NOTICE, LEGACY_SETTING_REFUSAL,
           UNKNOWN_TOOL_HINT,
           UNKNOWN_OP_HINT, MISSING_FILE_HINT, CEILING_HINT,
           *NEW_HINTS.values(), *WRONG_TYPE_FORMS.values(), *MISSING_FORMS.values(),
           *OP_LINES.values(), *OP_LINE_SUFFIX.values()]
    for slot in TOOLS.values():
        out += [slot["title"], slot["description"]]
    return out
