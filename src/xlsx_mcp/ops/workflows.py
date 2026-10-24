"""ops/workflows.py: recommended tool sequences for common multi-step tasks.

The Word editor's discoverability pattern: curated recipes as plain data, served by
get_workflows. Every step is SERVED as the tool and the op to call plus the
rationale (family surface: a recipe names only tools and ops that exist;
tests/unit/test_workflows_tool.py holds it). The recipe data below still keys
each step by its op name under "tool"; _step() turns it into {tool, op}.
Steps whose tool belongs to the COM tier that has not shipped yet carry
forthcoming=true and say so honestly; the registry test asserts every
non-forthcoming step names a tool that actually exists on the server.

The sequences are recommendations, not scripts: optional steps can be
skipped, and the notes carry the judgment calls a caller should know.
"""

from __future__ import annotations

from ..core.errors import XlMcpError

#: Tools named in steps that are DECLARED but not yet registered. The COM
#: tier shipped, so the set is empty; the registry test treats members as
#: pending rather than missing.
FORTHCOMING_TOOLS: set[str] = set()

# task -> {summary, steps: [{tool (the op name), why, optional?, forthcoming?}],
#          notes: [...]}
WORKFLOWS: dict[str, dict] = {
    "recover-workbook": {
        "summary": (
            "Your file will not open or an edit went wrong: get a good "
            "copy back from the automatic backups. Written for the worst "
            "moment, not for developers."
        ),
        "steps": [
            {"tool": "manage_backups", "why": "action='list' with path set to your file. Every edit "
                    "this server made kept two automatic copies next to "
                    "the file (in a hidden .ks4xl-backups folder): 'prev' "
                    "is the file as it was just before the most recent "
                    "change, 'anchor' is the file as it was when this "
                    "work session started. This shows you both, with "
                    "sizes and times"},
            {"tool": "manage_backups", "why": "action='restore', source='prev'. This puts the "
                    "before-the-last-change copy back as your file. It is "
                    "safe to try: restoring is itself undoable (the "
                    "current state is saved to prev first), and a backup "
                    "that is itself damaged is refused rather than "
                    "installed"},
            {"tool": "diagnose_workbook", "why": "confirm the recovered file reads as healthy"},
            {"tool": "manage_backups", "optional": True,
             "why": "if prev was not good enough, action='restore', "
                    "source='anchor' goes further back, to the state at "
                    "the start of the session"},
            {"tool": "manage_backups", "optional": True,
             "why": "if the damage came from a delete, a clear or a replace "
                    "and other edits followed it, the list from the first "
                    "step also shows snapshots labelled 'pre-<op> <time>', "
                    "each the file just before that change; "
                    "action='restore' with source set to the label puts "
                    "it back"},
            {"tool": "manage_backups", "optional": True,
             "why": "going forward: action='snapshot' makes a permanent, "
                    "date-stamped copy the automatic rotation never "
                    "touches. Take one before anything risky"},
        ],
        "notes": [
            "What the backups can NOT hold: prev is the state before the "
            "LAST change made through this server. If the file was "
            "damaged by something else AFTER the last save (a crash, a "
            "bad disk, another program), that final edit is not in any "
            "slot; only a snapshot habit covers it.",
            "Nothing here deletes anything: restore always saves the "
            "current state to prev first, and snapshots are never "
            "overwritten.",
        ],
    },
    "merge-workbooks": {
        "summary": (
            "Combine data from several workbooks into one target without "
            "silently losing anything either side holds."
        ),
        "steps": [
            {"tool": "get_workbook_metadata", "why": "orient on every source: sheets, true used ranges, "
                    "tables, and the hazard summary"},
            {"tool": "diagnose_workbook", "why": "confirm the TARGET is safe for file-based edits "
                    "before bulk writes"},
            {"tool": "copy_workbook", "why": "branch the target byte-for-byte so the merge is "
                    "disposable"},
            {"tool": "manage_worksheet", "why": "action='add' one sheet per incoming source"},
            {"tool": "read_range", "why": "pull each source's used range; values='both' labels "
                    "which cells are formula-backed"},
            {"tool": "write_range", "why": "land the data in the target sheet, one verified save "
                    "per source"},
            {"tool": "audit_formulas", "why": "check nothing landed as a broken or external-reference "
                    "formula"},
            {"tool": "validate", "why": "checks=['structure', 'references'] before delivering "
                    "the merged file"},
        ],
        "notes": [
            "Range copies carry values and formulas, not charts, images, "
            "or pivots; those stay with their source parts.",
            "A formula referencing its old workbook becomes an external "
            "reference in the target; audit_formulas surfaces them.",
        ],
    },
    "report-build": {
        "summary": (
            "Build a report workbook from raw data: import, table, "
            "formulas, chart, then the page and view polish."
        ),
        "steps": [
            {"tool": "create_workbook", "why": "a fresh file with named sheets for data and report"},
            {"tool": "import_data", "why": "CSV/JSON into the data sheet with type control and "
                    "the formula-injection lint"},
            {"tool": "create_table", "why": "promote the range to a real table for structured "
                    "references and styling"},
            {"tool": "set_formula", "why": "summary metrics over the table; modern functions are "
                    "normalized automatically"},
            {"tool": "recalculate", "why": "populate cached results so non-Excel readers see "
                    "numbers; needs xl_app (Windows + Excel), "
                    "otherwise the file recalculates on its next Excel "
                    "open"},
            {"tool": "format_cells", "why": "number formats and header emphasis"},
            {"tool": "manage_chart", "why": "action='create' a chart over the table range"},
            {"tool": "set_view", "why": "freeze the header row and set the opening zoom"},
            {"tool": "set_page_layout", "why": "print area, orientation, and fit for the report page"},
            {"tool": "set_workbook_properties", "why": "title and author so the file is self-describing"},
            {"tool": "export_file", "optional": True,
             "why": "ship a CSV/JSON copy of the data alongside the "
                    "workbook"},
        ],
        "notes": [
        ],
    },
    "data-cleanup": {
        "summary": (
            "Clean a messy data sheet: find and replace, then validation "
            "rules, then sort, verified at each end."
        ),
        "steps": [
            {"tool": "diagnose_workbook", "why": "the hazard verdict decides whether file-based edits "
                    "are safe at all"},
            {"tool": "manage_backups", "why": "action='snapshot': a permanent DTG-stamped keeper "
                    "before surgery"},
            {"tool": "find_cells", "why": "locate the dirty values and patterns first"},
            {"tool": "replace_cells", "why": "dry_run=true to preview the change list, then apply"},
            {"tool": "manage_data_validation", "why": "add rules so the cleaned columns stay clean"},
            {"tool": "sort_range", "why": "multi-key sort that moves whole rows atomically"},
            {"tool": "validate", "why": "checks=['structure', 'references', 'calc_staleness'] "
                    "to close out"},
        ],
        "notes": [
            "query_range answers 'what is still dirty' without pulling the "
            "sheet into context.",
        ],
    },
    "formatting-audit-and-fix": {
        "summary": (
            "Diagnose style bloat against the 64,000-format ceiling and "
            "consolidate formatting."
        ),
        "steps": [
            {"tool": "audit_styles", "why": "the counters, the heaviest formats, and the risk "
                    "verdict against the ceiling"},
            {"tool": "get_grid_view", "why": "see what the heavy areas actually look like before "
                    "clearing anything"},
            {"tool": "clear_range", "why": "what='formats' on runaway areas resets them to "
                    "default"},
            {"tool": "apply_style", "why": "named styles instead of thousands of per-cell "
                    "one-offs"},
            {"tool": "format_cells", "why": "reapply the intended formatting deliberately"},
            {"tool": "validate", "why": "checks=['formatting_bloat'] confirms the risk verdict "
                    "moved"},
        ],
        "notes": [
        ],
    },
    "safe-edit-of-rich-workbook": {
        "summary": (
            "Edit a workbook holding fragile parts (pivots, charts, "
            "macros, queries) without silent loss: diagnose, hazard-aware "
            "edit, verify."
        ),
        "steps": [
            {"tool": "diagnose_workbook", "why": "the hazard scan names what a file-based save would "
                    "drop and recommends the route"},
            {"tool": "manage_backups", "why": "action='snapshot' before touching a rich file; the "
                    "automatic prev/anchor slots cover per-edit undo"},
            {"tool": "get_grid_view", "why": "see the region and its formula markers before "
                    "editing"},
            {"tool": "apply_edits", "why": "one atomic batch: every anchor validated first, one "
                    "backup, one save, one verify-after-write"},
            {"tool": "audit_formulas", "why": "confirm no reference broke and note which results "
                    "are stale"},
            {"tool": "validate", "why": "checks=['structure', 'references', 'hazards'] as the "
                    "exit gate"},
            {"tool": "manage_backups", "optional": True,
             "why": "action='restore' source='prev' is the undo if review "
                    "rejects the edit"},
        ],
        "notes": [
            "A would-lose verdict means mutating tools refuse. There are "
            "two routes: pass allow_loss:true to accept the named loss "
            "(the backup is taken first), or make the change in Excel "
            "itself. xl_app cannot make cell, format or structure edits, "
            "so it is not a route for this.",
            "Verify-after-write already runs on every save; validate is "
            "the independent read-back on top of it.",
            "For one risky save, pass verify_com:true to any mutating "
            "tool: after the write the file is opened in a hidden Excel "
            "and a repair-free open is required, or the backup is "
            "restored and the save refuses. It needs Excel and costs a "
            "COM round trip, so it is off by default; KS4XL_VERIFY_COM=1 "
            "turns it on for every save instead.",
        ],
    },
    "migrate-from-v1": {
        "summary": (
            "Translate 1.x calls (one tool per operation) to 2.0 (seven tools, "
            "each taking {op, args}); map lists every released 1.2.5 tool."
        ),
        "steps": [],
        "notes": [
            "Every 1.x tool is now an op: call its 2.0 tool with "
            "{\"op\": <the 1.x tool name>, \"args\": <the same arguments>}. "
            "map gives the tool for each name and any argument that changed.",
            "Every op is available from the start: the two 1.x tools that "
            "switched groups of tools on and off are gone (the map rows with "
            "no op). An old KS4XL_MODE or KS4XL_PACK_POLICY default is ignored "
            "with a one-line notice, and one that locked the tools stops the "
            "server; use KS4XL_READ_ONLY or KS4XL_FORBID to limit it.",
            "Values are snake_case: conditional-format operators and time "
            "periods (greater_than, last_7_days), validation types and "
            "operators (text_length, not_between) and calc_mode "
            "(auto_no_table). The camelCase spellings still work.",
            "manage_chart's cell argument is at_cell (anchor still works). "
            "Results name a cell at_cell or top_left and a chart's type "
            "chart_type; anchor means only a get_grid_view token.",
            "A 1.x call keeps working when its op name is sent with other "
            "casing or separators, and Word's file_path names the path.",
        ],
        "generated_map": True,
    },
    "migrate-from-incumbent": {
        "summary": (
            "Tool-by-tool mapping from the common file-based Excel MCP "
            "server surface onto KS4XL."
        ),
        "steps": [
            {"tool": "read_range", "why": "replaces read_data_from_excel; values='both' labels "
                    "cached vs formula so an uncalculated cell is never a "
                    "silent blank"},
            {"tool": "query_range", "why": "no incumbent analog: filter, sort, and aggregate "
                    "server-side instead of dumping whole sheets"},
            {"tool": "write_range", "why": "replaces write_data_to_excel; every save is atomic, "
                    "backed up, and verified"},
            {"tool": "set_formula", "why": "replaces apply_formula; modern functions are "
                    "normalized so they do not land as #NAME?"},
            {"tool": "validate", "why": "replaces validate_formula_syntax and "
                    "validate_excel_range; whole-workbook checks"},
            {"tool": "format_cells", "why": "replaces format_range"},
            {"tool": "set_merge", "why": "replaces merge_cells, unmerge_cells, and "
                    "get_merged_cells in one tool"},
            {"tool": "manage_worksheet", "why": "replaces create_worksheet, rename_worksheet, "
                    "delete_worksheet, and copy_worksheet"},
            {"tool": "modify_grid_structure", "why": "replaces insert_rows, insert_columns, "
                    "delete_sheet_rows, and delete_sheet_columns, and "
                    "rewrites every affected reference"},
            {"tool": "create_table", "why": "replaces create_table like for like, adding "
                    "name-collision and overlap checks"},
            {"tool": "manage_chart", "why": "replaces create_chart"},
            {"tool": "com_manage_pivot", "why": "replaces create_pivot_table with a real refreshable "
                    "pivot through Excel (xl_app; Windows + Excel)"},
        ],
        "notes": [],
    },
}


#: A panicked user types "recover", "undo", or "disaster", not the exact
#: task name; those words must land on the recipe.
_TASK_ALIASES: dict[str, str] = {
    "migrate": "migrate-from-v1",
    "migration": "migrate-from-v1",
    "upgrade": "migrate-from-v1",
    "recover": "recover-workbook",
    "recovery": "recover-workbook",
    "restore": "recover-workbook",
    "undo": "recover-workbook",
    "disaster": "recover-workbook",
    "corrupt": "recover-workbook",
}


def _step(step: dict) -> dict:
    """A recipe step on the family surface: the recipe data names the op
    (its old tool name); the served step names the tool that holds it and
    the op, so a caller can make the call exactly as written."""
    from ..surface.groups import OP_GROUP

    op = step["tool"]
    out = {"tool": OP_GROUP.get(op), "op": op}
    out.update({k: v for k, v in step.items() if k != "tool"})
    return out


def get_workflows(task: str | None = None) -> dict:
    """List available tasks, or return one task's recipe."""
    if task is None:
        return {
            "tasks": [{"task": name, "summary": wf["summary"]}
                      for name, wf in WORKFLOWS.items()],
            "note": "call xl_read get_workflows with task='<name>' for the recipe.",
        }
    key = _TASK_ALIASES.get(str(task).strip().lower(), task)
    wf = WORKFLOWS.get(key)
    if wf is None:
        raise XlMcpError(
            f"unknown task {task!r}; tasks: " + ", ".join(WORKFLOWS))
    out = {"task": key, **{k: v for k, v in wf.items() if k != "generated_map"},
           "steps": [_step(step) for step in wf["steps"]]}
    if wf.get("generated_map"):
        out["map"] = migration_map_v1()
    return out


#: Released tool names with no 2.0 op.
REMOVED_IN_V2 = ("enable_tools", "disable_tools")


def migration_map_v1() -> list[dict]:
    """The 1.x -> 2.0 map, generated from the live registry and the
    surface's own rename tables, so it cannot drift from what the server
    accepts (scripts/generate_surface_map.py writes the same facts to
    migration/v1_to_v2.json; a test checks both against the released 1.2.5
    tool list)."""
    from .. import registry
    from ..surface.groups import OP_GROUP

    rows: list[dict] = []
    for op in sorted(registry.public_ops()):
        row: dict = {"old": op, "tool": OP_GROUP.get(op), "op": op}
        changed = args_changed(op)
        if changed:
            row["changed"] = changed
        rows.append(row)
    rows += [{"old": name, "tool": None, "op": None} for name in REMOVED_IN_V2]
    return rows


def args_changed(op: str) -> dict:
    """What a 1.x call of op must spell differently in 2.0 (both forms are
    accepted): renamed arguments and snake_case values."""
    from ..surface import normalize

    changed = dict(normalize.RENAMED_ARGS.get(op, {}))
    for (vop, arg), table in normalize.VALUE_TABLES.items():
        if vop == op:
            changed.update({f"{arg}={xml}": f"{arg}={snake}"
                            for snake, xml in table.items() if snake != xml})
    return changed


__all__ = ["get_workflows", "WORKFLOWS", "FORTHCOMING_TOOLS", "migration_map_v1",
           "args_changed"]
