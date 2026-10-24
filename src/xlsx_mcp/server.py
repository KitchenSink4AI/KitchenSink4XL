"""kitchensink4xl: the consolidated FastMCP surface.

The surface: the infrastructure (envelope, op registry, dispatcher,
tiered loading, the KS4XL_MODE startup route) plus the file-tier domain
families over the core engine (core/package, hazard, verify, locate, refs,
calc): lifecycle and discovery, cells and ranges with the token-shaped
query_range, the grid view and batch layer, structural edits with full
reference rewriting (modify_grid_structure), merges, workbook-wide
find/replace, scatter cell reads/writes, formatting and the style layer
(named styles, format painter, style-bloat audit), tables and names,
conditional formatting and data validation, sort and filter, comments and
hyperlinks, import and export.

Contract (carried from the family):
- Positional tools take the grid location object, resolved
  through core.locate (ambiguity is a loud refusal carrying every match).
- The refusal envelope applies AT THE MCP BOUNDARY: the registered tool
  wraps the module-level function, converting typed exceptions into
  structured {ok: false, error: {...}} refusals with isError=true.
- Every mutating tool auto-backs-up before the mutation and verifies after
  the write (the safety core); saves are atomic and validated. Each also
  takes verify_com:true, the DEEP check: the produced file must open in a
  real hidden Excel without a repair prompt or the backup is restored and
  the save refuses. It is off by default (a COM round trip, and Excel is
  not there in headless CI); KS4XL_VERIFY_COM=1 makes it the default for
  every save, and the per-call parameter wins over that either way.
- FAMILY SURFACE: the module-level functions are
  OPS, not MCP tools. `_op` records each in xlsx_mcp.registry and leaves the
  function untouched (in-process callers keep raw returns and typed
  exceptions); create_server() builds the fixed seven-tool list from the
  registry (xlsx_mcp.surface) and every call goes through the one
  dispatcher. There are no packs and no runtime tool-list changes.
"""

from __future__ import annotations

import os as _os
import platform as _platform
import sys as _sys
from typing import Literal

from . import __version__
from . import license_policy as _license_policy
from . import registry as _registry
from . import runtime as _runtime
from .core import package as _package
from .core import schemas as _schemas
from .core import sandbox as _sandbox
from .core import update_check as _upd
from .core import star_nudge as _star_nudge
from .ops import annotations as _annotations
from .ops import backups as _backups
from .ops import cells as _cells
from .ops import comtier as _comtier
from .ops import inspectors as _inspectors
from .ops import objects as _objects
from .ops import pagelayout as _pagelayout
from .ops import protection as _protection
from .ops import condformat as _condformat
from .ops import datavalidation as _datavalidation
from .ops import dataio as _dataio
from .ops import format as _format
from .ops import formulas as _formulas
from .ops import lifecycle as _lifecycle
from .ops import names as _names
from .ops import properties as _properties
from .ops import search as _search
from .ops import sheetview as _sheetview
from .ops import sortfilter as _sortfilter
from .ops import structure as _structure
from .ops import tables as _tables
from .ops import validation as _validation
from .ops import view as _view
from .ops import workflows as _workflows

# The injected policy, checked at import (edition split): a core imported
# without one raises PolicySetupError HERE, before a single op registers and
# long before a workbook is opened. That is a setup error and reads as one; a
# missing or unverifiable receipt is a different thing entirely and never
# reaches this line. create_server() reads the policy again for the server
# it builds.
_license_policy.active()


# ------------------------------------------------------- op registration


def _op(fn):
    """Register a module-level function as an op of the family surface.

    The function itself is returned untouched, so in-process callers (the
    test suite's srv.<name>(...) convention) keep raw returns and typed
    exceptions. The registry records the op's tool (surface/groups.py), its
    generated strict args model, and its read-only and destructive
    classification (core/readonly.py, core/tool_annotations.py); an
    unclassified op raises HERE, at import, before any list is built."""
    from .surface.groups import group_of

    _registry.register(fn, group_of=group_of)
    return fn


# ------------------------------------------------------- placeholder reader


def _update_check_status() -> dict:
    """The update report, without its cache path.

    The on-demand check: it may ask PyPI, at most once every seven days,
    with a two-second timeout (the weekly background check, started after a
    tool call, shares the same cache). Nothing runs at startup, and a check
    that fails or is switched off reports that fact rather than going quiet.
    state is one of: disabled, update_available, current, unknown. Never
    raises.
    """
    try:
        return _upd.status()
    except Exception:
        return {"state": "unknown", "note": _upd.NOTE_UNKNOWN}


def _surface_facts() -> dict:
    """The fixed surface this server lists: tool and op counts, read-only
    mode, and what the launch settings removed."""
    from .surface import copy as _copy

    plan = _runtime.plan()
    settings = plan.settings
    return {
        "tools": len(plan.tools),
        "ops": plan.op_count(),
        "read_only": settings.read_only,
        "forbidden": sorted(set(plan.forbidden_tools) | set(plan.forbidden_ops)),
        "decided_by": _copy.SETTINGS_OWNER,
        "legacy_settings_ignored": list(settings.legacy_set),
        # What the tool list costs a client to load, by the
        # counter scripts/measure_surface.py uses (compact JSON of each
        # listed tool / 4); Word reports the same field.
        "approx_active_tokens": _approx_active_tokens(plan),
    }


def _approx_active_tokens(plan) -> int:
    import json as _json

    total = 0
    for tp in plan.tools.values():
        tool = {"name": tp.name, "title": tp.title, "description": tp.description,
                "inputSchema": tp.input_schema, "annotations": tp.annotations}
        total += round(len(_json.dumps(tool, ensure_ascii=False, separators=(",", ":"))) / 4)
    return total


def _installed_build() -> dict:
    """Which bundle this server came out of, read from the environment
    stamp the installer wrote. Three nulls on a Community install, which is
    how a support bundle tells the two editions apart at a glance."""
    try:
        return _upd.installed_build()
    except Exception:
        return {"bundle_version": None, "release_id": None, "channel": None}


def _os_line() -> str:
    """The operating system, for get_server_info, WITHOUT asking WMI (the
    same design as the Word editor).

    platform.platform() on Windows reaches platform._wmi_query (Python 3.12+:
    win32_ver and the machine lookup), and one such query can sit in
    _wmi.exec_query for minutes. uname is cached
    only on success, so every retry repeated the stuck query. On Windows the
    line is built from sys.getwindowsversion() and the process environment,
    which answer instantly; elsewhere platform.platform() uses os.uname and
    is kept."""
    if _sys.platform == "win32":
        import os as _os

        v = _sys.getwindowsversion()
        arch = _os.environ.get("PROCESSOR_ARCHITEW6432") or _os.environ.get(
            "PROCESSOR_ARCHITECTURE") or "unknown"
        return f"Windows-{v.major}.{v.minor}.{v.build}-{arch}"
    return _platform.platform()


@_op
def get_server_info(include_licensee: bool = False) -> dict:
    """Report KitchenSink4XL version, edition, licensing status, active tool
    count, op count and launch-settings narrowing, platform, Python, and sandbox/
    Excel verification settings (root count, never paths). Needs no workbook.
    include_licensee explicitly discloses the signed company identity.
    Business reads only a configured offline receipt. Community may check
    PyPI and update its cache; KS4XL_UPDATE_CHECK=off disables that I/O."""
    _pol = _license_policy.active()
    out = {
        "name": _pol.product_id,
        # Self-identification for a connected agent: the brand, the
        # distribution to pip install, where the docs live, and the three
        # siblings. `name` stays the PyPI distribution for compatibility
        # with anything already reading it.
        "product": "KitchenSink4XL",
        "package": _pol.product_id,
        "edition": _pol.edition_id,
        # include_licensee is the client-visible parameter and does not
        # change; redaction is the policy-side default it maps onto.
        "licensing": _pol.describe_license(redacted=not include_licensee),
        "homepage": "https://community.kitchensink4.ai/xl/",
        "family": ["kitchensink4word", "kitchensink4ppt", "kitchensink4web"],
        # The edition distribution's version, which is what a customer
        # installed and what the old single wheel reported. The shared
        # editing core's own version is reported alongside it.
        "version": _pol.distribution_version,
        # What stage of the product this is, and when this build was cut.
        # Business reads both out of its generated build metadata; Community
        # has neither and reports null for both. getattr, because a policy
        # adapter older than these two fields is still a supportable
        # installation and not a reason to fail a status call.
        "channel": getattr(_pol, "channel", None),
        "release_date": getattr(_pol, "release_date", None),
        # WHAT THIS BUILD TRUSTS, and how far it can follow a receipt back to
        # it. chain_version is the certificate schema this build understands,
        # and it is null when the build pins no roots and therefore cannot
        # follow a chain at all. trust_list_version says which retirement list
        # it shipped with, so support can tell at a glance whether an
        # installation has heard that a key was pulled. root_key_ids names the
        # anchors, which is the licensing fact an operator reads aloud when
        # comparing a key out of band. None of the three is a secret: a root's
        # public half is public by construction, and Community reports the
        # same nulls it reports for channel.
        "chain_version": getattr(_pol, "chain_version", None),
        "trust_list_version": getattr(_pol, "trust_list_version", None),
        "root_key_ids": getattr(_pol, "root_key_ids", []),
        "core_version": __version__,
        "surface": _surface_facts(),
        "platform": _os_line(),
        "python": _sys.version.split()[0],
        "config": {
            "sandbox_active": _sandbox.active(),
            "allowed_roots_count": _sandbox.root_count(),
            "verify_com_default": _package.com_verify_default(),
        },
        # When the launch settings turn Excel off,
        # say so, and that verify_com and KS4XL_VERIFY_COM are off with it.
        **({"excel": _runtime.excel_forbidden()} if _runtime.excel_forbidden() else {}),
    }
    # The on-demand update report: at most once every seven days, two-second
    # timeout, off under KS4XL_UPDATE_CHECK=off. Nothing runs at startup, a
    # failed check reports the failure, and this never raises. (The weekly
    # background check starts after a tool call and shares the same cache.)
    out["update_check"] = _update_check_status()
    # WHICH BUNDLE THIS IS. Three product versions can repeat across two
    # bundles, so the product version alone does not name the build a
    # support case is looking at. These three do, and they are null on a
    # Community install, which is itself the answer to "which edition is
    # this".
    build = _installed_build()
    out["bundle_version"] = build["bundle_version"]
    out["release_id"] = build["release_id"]
    out["channel"] = out.get("channel") or build["channel"]
    notice = _upd.update_notice()
    if notice:
        out["update"] = notice
    return out


# ================================================================ DATA PLANE
# The core data plane: lifecycle + discovery, cell/range read/write, the token-
# shaped read (query_range), the grid view + batch layer, and formatting. All
# lite (the daily driver). Each registered wrapper is a thin dispatch over an
# ops/ module; every mutation runs through WorkbookPackage, so the hazard gate,
# backup-before-mutation, and verify-after-write are automatic.


# ------------------------------------------------- lifecycle and discovery


@_op
def create_workbook(path: str, sheets: list[str] | None = None,
                    overwrite: bool = False) -> dict:
    """Create a new .xlsx workbook at path with the given sheet names (default a
    single 'Sheet1'). Sheet names must be unique, at most 31 characters, and
    may not start or end with an apostrophe (Excel refuses to open such a
    file); the parent directory must already exist. An existing file at path
    is left untouched unless overwrite is true, in which case it is FIRST
    rotated into its .ks4xl-backups prev slot and then replaced (restore
    source='prev' brings it back). Returns the file path and the sheets
    created."""
    return _lifecycle.create_workbook(path, sheets=sheets, overwrite=overwrite)


@_op
def copy_workbook(src: str, dst: str, overwrite: bool = False) -> dict:
    """Copy a workbook file byte-for-byte from src to dst, so nothing in the
    original is re-serialized or degraded (charts, shapes, macros, and queries
    all carry over intact). An existing dst is left untouched unless overwrite
    is true, in which case dst is FIRST rotated into its .ks4xl-backups prev
    slot and then replaced (restore source='prev' brings it back). Returns
    the destination path. Use this to branch a working copy before a risky
    batch of edits."""
    return _lifecycle.copy_workbook(src, dst, overwrite=overwrite)


@_op
def get_workbook_metadata(path: str) -> dict:
    """Read a workbook's structure without opening it for edit: every sheet with
    its visibility state, TRUE used range (value-bearing bounds, not the often
    wrong stored dimension), dimensions, and merged-cell count, plus defined
    names, tables, the active sheet, and a round-trip hazard summary (whether
    an openpyxl edit would drop fragile parts). The orient-before-editing
    call. Read-only; works while the file is open in Excel."""
    return _lifecycle.get_workbook_metadata(path)


@_op
def diagnose_workbook(path: str) -> dict:
    """The round-trip hazard scan surfaced as a health readout: which fragile
    parts the workbook holds (slicers, shapes, embedded objects, Power Query,
    VBA, and the rest), whether a file-based openpyxl edit would drop any of
    them, the routing recommendation for a surgical versus a structural edit,
    and a light integrity summary (sheet counts, formula-cell count, keep_vba).
    This is how you check a workbook is safe to edit before mutating it.

    What to do with the verdict: hazards never block reads; a would-lose
    verdict means every mutating tool will refuse unless you route through
    Excel (xl_app) or pass allow_loss:true (an explicit, backed-up
    acceptance of the loss). A clean verdict means file-based edits are
    round-trip safe. Content with no part of its own is covered too: the scan
    reads each worksheet's extLst, so x14 conditional formats (data bars, icon
    sets), sparkline groups and slicer lists come back as a would-lose verdict
    like any other drop-risk hazard. Limit: the extLst walk looks at the
    worksheet's top level, and an extension openpyxl drops from anywhere else
    is caught at save time by openpyxl's own load warning rather than
    here. Read-only."""
    return _lifecycle.diagnose_workbook(path)


@_op
def manage_worksheet(path: str, action: str, sheet: str | None = None,
                     new_name: str | None = None, index: int | None = None,
                     state: str | None = None, allow_loss: bool = False,
                     backup: bool = True,
                     verify_com: bool | None = None) -> dict:
    """Manage the worksheet lifecycle. action is one of: add (new_name, optional
    index), delete (sheet), rename (sheet, new_name), copy (sheet, optional
    new_name), reorder (sheet, index as 0-based target), hide (sheet, state
    'hidden' or 'very_hidden'), unhide (sheet). The workbook always keeps at
    least one VISIBLE sheet, so deleting or hiding the last visible one
    refuses.

    References follow, as in Excel: rename rewrites every reference to the
    sheet (formulas and 3D spans, defined names, conditional-format and
    validation formulas, chart references, pivot sources), and delete turns
    references to the deleted sheet into #REF! (a 3D span whose end it was
    shrinks to the next sheet), both counted in references_rewritten.
    Hyperlink locations are left as they are, as Excel leaves them. A
    hazardous workbook refuses unless allow_loss is true. Auto-backup:
    prev/anchor slots in .ks4xl-backups (backup=false skips rotation); atomic
    verified save, restored on failed verify. Refuses while open in Excel."""
    return _lifecycle.manage_worksheet(
        path, action, sheet=sheet, new_name=new_name, index=index,
        state=state, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# --------------------------------------------------------- cells and ranges


@_op
def read_range(path: str, location: _schemas.Location, values: str = "cached",
               sheet: str | None = None) -> dict:
    """Read a cell or range addressed by a location object (cell, range, r1c1,
    name, table, used_range, region, search, or a grid-view anchor; sheet
    picks the sheet). values: 'cached' the last calculated values, 'formula'
    the formulas, 'both' each value labelled (cached, absent, formula,
    value), 'displayed' the text each cell shows in Excel (0.75 under # ?/?
    shows ' 3/4'). cached and both add shows, {cell: text}, where that text
    differs from the value. A formula with no cached value is 'absent',
    never blank. Read-only."""
    return _cells.read_range(path, location, values=values, sheet=sheet)


@_op
def set_cell(path: str, location: _schemas.Location, value: _schemas.CellValue,
             sheet: str | None = None, allow_loss: bool = False,
             backup: bool = True, verify_com: bool | None = None) -> dict:
    """Write a single cell addressed by a location object. A string beginning
    with '=' is ALWAYS stored as a formula (there is no literal escape),
    normalized so modern functions do not land as #NAME? and flagged to
    recalculate on open; anything else is a literal. A hazardous workbook
    refuses unless allow_loss is true. Auto-backup: prev/anchor slots in
    .ks4xl-backups (backup=false skips rotation); atomic verified save,
    restored on failed verify. Refuses while open in Excel."""
    return _cells.set_cell(path, location, value, sheet=sheet,
                           allow_loss=allow_loss, backup=backup,
                               verify_com=verify_com)


@_op
def write_range(path: str, location: _schemas.Location,
                data: _schemas.CellMatrix, sheet: str | None = None,
                parse_dates: bool = False, echo_cells: bool = False,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Write a 2D block of values and formulas from the location's top-left
    cell. data is a list of equal-length rows (null clears a cell); nothing
    outside the block changes. A date is {"date": "2026-03-24"} (parse_dates
    =true also reads ISO date text); {"text": "=A1"} keeps text that looks
    like a formula. Up to 1,000,000 cells; the answer lists 20 written cells
    (echo_cells=true lists all). A hazardous workbook refuses unless
    allow_loss; auto-backup, verified save; refuses while open in Excel."""
    with _package.echo_cells(echo_cells):
        return _cells.write_range(path, location, data, sheet=sheet,
                                  parse_dates=parse_dates,
                                  allow_loss=allow_loss, backup=backup,
                                  verify_com=verify_com)


@_op
def clear_range(path: str, location: _schemas.Location, what: str = "contents",
                sheet: str | None = None, allow_loss: bool = False,
                backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Clear a cell or range: what='contents' removes values and formulas,
    'formats' resets styles to default, 'all' does both. Neither removes
    merges, conditional formats, validations, comments, or hyperlinks; those
    have their own manage tools. Addressed by a location object. A hazardous
    workbook refuses unless allow_loss is true. Auto-backup: prev/anchor slots
    in .ks4xl-backups (backup=false skips rotation); atomic verified save,
    restored on failed verify. Refuses while open in Excel."""
    return _cells.clear_range(path, location, what=what, sheet=sheet,
                              allow_loss=allow_loss, backup=backup,
                                  verify_com=verify_com)


@_op
def copy_range(path: str, source: _schemas.SourceLocation,
               dest: _schemas.DestLocation, what: str = "all",
               adjust_formulas: bool = True, sheet: str | None = None,
               allow_loss: bool = False, backup: bool = True,
               verify_com: bool | None = None) -> dict:
    """Copy a source rectangle to a destination anchor (location objects, may
    name different sheets). what is 'all', 'values', 'formulas', or 'formats'.
    The destination is overwritten; the source is buffered first, so an
    overlapping paste is safe. Relative refs in copied formulas shift by the
    paste offset like an Excel copy unless adjust_formulas is false; absolute
    ($) anchors stay put. A hazardous workbook refuses unless allow_loss is
    true. Auto-backup to .ks4xl-backups; atomic verified save. Refuses while
    open in Excel."""
    return _cells.copy_range(path, source, dest, what=what,
                             adjust_formulas=adjust_formulas, sheet=sheet,
                             allow_loss=allow_loss, backup=backup,
                                 verify_com=verify_com)


@_op
def move_range(path: str, source: _schemas.SourceLocation,
               dest: _schemas.DestLocation, sheet: str | None = None,
               allow_loss: bool = False, backup: bool = True,
               verify_com: bool | None = None) -> dict:
    """Move a rectangle to a new anchor on the same sheet, rewriting every
    formula, name, conditional format, validation, table ref, and merge that
    pointed into the source so references follow the cells (Excel move
    semantics). A cross-sheet destination refuses (copy_range then clear_range
    instead); cells at the destination are overwritten. A hazardous workbook
    refuses unless allow_loss is true. Auto-backup to .ks4xl-backups; atomic
    verified save. Refuses while open in Excel."""
    return _cells.move_range(path, source, dest, sheet=sheet,
                             allow_loss=allow_loss, backup=backup,
                                 verify_com=verify_com)


@_op
def query_range(path: str, location: _schemas.OptionalLocation = None,
                sheet: str | None = None, header: bool = True,
                columns: _schemas.ColumnList = None,
                where: _schemas.Predicates = None, match: str = "all",
                order_by: _schemas.OrderBy = None,
                aggregate: _schemas.Aggregates = None,
                group_by: _schemas.GroupBy = None, limit: int | None = None,
                offset: int = 0, distinct: bool = False, records: bool = False,
                values: str = "cached") -> dict:
    """Filter, project, sort, paginate, and aggregate a range SERVER-SIDE so an
    agent reads only the rows and columns it needs.

    location defaults to the sheet's true used range. With header=true the
    first row names the columns (else the A1 letters). where is a
    list of {column, op, value} predicates joined by match ('all' or 'any');
    ops: eq, ne, gt, ge, lt, le, contains, startswith, endswith, regex, in,
    not_in, is_blank, not_blank. columns projects a subset; order_by is a list
    of {column, dir} specs (unknown directions refuse); offset and limit page;
    distinct dedupes. aggregate is a list of {column, func} (count,
    count_nonblank, count_distinct, sum, avg, min, max, first, last),
    optionally per group_by (records=true emits objects).

    Semantics: predicates read CACHED and literal values
    (values='displayed': the shown text; uncalculated formulas read as blank;
    recalc for exact results); gt/ge/lt/le compare
    numerically when both sides coerce, else case-folded text; blanks
    never satisfy ordered comparisons; regex is timeout-guarded.
    Aggregates follow Excel: sum/avg/min/max consume NUMERIC cells only
    (text and booleans ignored even when text looks numeric; exclusions
    are reported); count is the RAW row count, unlike Excel COUNT; min/max
    fall back to text when no numbers exist. Filter-hidden rows are read
    and aggregated like any other row, unlike SUBTOTAL; use where to drop
    them. Read-only."""
    return _cells.query_range(
        path, location=location, sheet=sheet, header=header, columns=columns,
        where=where, match=match, order_by=order_by, aggregate=aggregate,
        group_by=group_by, limit=limit, offset=offset, distinct=distinct,
        records=records, values=values)


# ------------------------------------------------------- grid view + batch


@_op
def get_grid_view(path: str, location: _schemas.OptionalLocation = None,
                  sheet: str | None = None, max_rows: int = 50,
                  max_cols: int = 30, values: str = "cached") -> dict:
    """A compact, token-efficient projection of a sheet or range: the true used
    range, a markdown table with A1 addressing (column letters across the top,
    row numbers down the side), formula and merged-cell markers, dimensions, and
    the hazard summary, so an agent can see the grid without a per-cell JSON
    dump.

    location defaults to the sheet's used range. values='cached' shows last
    calculated values with formula cells marked (the florin character U+0192
    marks a formula with no cached value), and shows maps each shown cell
    whose text in Excel differs from its value to that text (read_range
    with values='displayed' gives the text of every cell); 'formula' shows
    the formula strings. The view paginates with max_rows and max_cols (caps 200 and 100)
    and reports truncated flags so the caller knows when to page.
    formula_cells maps addresses to their formula strings; merged ranges
    intersecting the view are listed. The result's anchor is a token for the
    shown rectangle: {"anchor": token} addresses it in any positional tool,
    refusing STALE_ANCHOR if the region changed since this view; cells
    inside stay plain A1. not_shown names what the table leaves out (notes
    and threaded comments on the shown cells, the sheet's charts and
    images) with the op that reads each. Read-only; works while the file is
    open in Excel. Pair it with apply_edits to edit what you see."""
    return _view.get_grid_view(path, location=location, sheet=sheet,
                               max_rows=max_rows, max_cols=max_cols,
                               values=values)


@_op
def apply_edits(path: str, edits: _schemas.BatchEdits, echo_cells: bool = False,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Apply many addressed edits as ONE atomic batch. edits is a list of
    {op, location, ...}: set_value {value}, set_formula {formula}, clear
    {what: contents|formats|all}, write_range {data: 2D array}. location is any
    location object, including a stale-checked get_grid_view anchor.

    Every location is resolved and every op validated BEFORE anything is
    written, so a single bad edit refuses the whole batch and the file stays
    byte-for-byte unchanged. The batch then takes ONE backup (prev/anchor
    slots in .ks4xl-backups), does ONE atomic save, and runs ONE
    verify-after-write, which restores from the backup if the produced file
    fails to read back as intended. Formula edits are normalized and flag
    recalculation; each write_range op honors the 200,000-cell ceiling. A
    hazardous workbook refuses unless allow_loss is true; refuses while the
    file is open in Excel. Returns the count of edits applied and cells
    touched."""
    with _package.echo_cells(echo_cells):
        return _view.apply_edits(path, edits, allow_loss=allow_loss, backup=backup,
                                 verify_com=verify_com)


# --------------------------------------------------------- formatting


@_op
def format_cells(path: str, location: _schemas.Location,
                 number_format: str | None = None, font: dict | None = None,
                 fill: dict | None = None, border: dict | None = None,
                 alignment: dict | None = None, sheet: str | None = None,
                 allow_loss: bool = False, backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Apply formatting to a range, merging onto the existing style so
    unspecified attributes are preserved. number_format is an Excel format
    code; font is {name, size, bold, italic, underline, strike, color}; fill
    is {color} or {pattern, fg, bg}; border is {style, color, sides};
    alignment is {horizontal, vertical, wrap_text, text_rotation, indent};
    colors are hex. Named styles and conditional formats: xl_format.
    Hazard-gated (allow_loss overrides); auto-backup; atomic verified save.
    Refuses while open in Excel."""
    return _format.format_cells(
        path, location, number_format=number_format, font=font, fill=fill,
        border=border, alignment=alignment, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_op
def set_dimensions(path: str, sheet: str | None = None,
                   column_widths: dict | None = None,
                   row_heights: dict | None = None,
                   autofit_columns: _schemas.ColumnNames = None,
                   hide_columns: _schemas.ColumnNames = None,
                   hide_rows: _schemas.RowNumbers = None,
                   group_rows: _schemas.OutlineSpans = None,
                   group_columns: _schemas.OutlineSpans = None,
                   ungroup_rows: _schemas.OutlineSpans = None,
                   ungroup_columns: _schemas.OutlineSpans = None,
                   outline_summary: dict | None = None,
                   allow_loss: bool = False, backup: bool = True,
                   verify_com: bool | None = None) -> dict:
    """Set column widths and row heights, hide rows or columns, group them
    into collapsible outline levels, and service an autofit request.
    column_widths maps column letters or indices to widths in Excel character
    units; row_heights maps row numbers to heights in points; autofit_columns
    sizes columns to content as a best-effort APPROXIMATION (true autofit
    needs Excel via xl_app).

    Grouping is Excel's outline: group_rows and group_columns take spans of
    {start, end, level 1-7, collapsed}, the collapse bracket a budget or
    financial model is built around; ungroup_rows and ungroup_columns clear
    a span's level and unhide it; outline_summary {below, right} says which
    side the total sits on. Read the current outline back from
    get_workbook_metadata, which reports it per sheet when a sheet has one.
    A hazardous workbook refuses unless allow_loss is true. Auto-backup to
    .ks4xl-backups; atomic verified save. Refuses while open in Excel."""
    return _format.set_dimensions(
        path, sheet=sheet, column_widths=column_widths, row_heights=row_heights,
        autofit_columns=autofit_columns, hide_columns=hide_columns,
        hide_rows=hide_rows, group_rows=group_rows,
        group_columns=group_columns, ungroup_rows=ungroup_rows,
        ungroup_columns=ungroup_columns, outline_summary=outline_summary,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


# ============================================ TABLES, RULES, COMMENTS, EXPORT
# Ungated file-tier families: tables and named ranges,
# conditional formatting and data validation, sort and filter, comments and
# hyperlinks, import and export. Each registered wrapper is a thin dispatch over
# an ops/ module; every mutation runs through WorkbookPackage, so the hazard
# gate, backup-before-mutation, and verify-after-write are automatic. Structural
# table edits reuse core.refs so refs stay coherent.


# ------------------------------------------------------------------- tables


@_op
def create_table(path: str, location: _schemas.Location, name: str,
                 header: bool = True, style: str = "TableStyleMedium9",
                 row_stripes: bool = True, col_stripes: bool = False,
                 totals_row: bool = False, totals: dict | None = None,
                 sheet: str | None = None, allow_loss: bool = False,
                 backup: bool = True, verify_com: bool | None = None) -> dict:
    """Turn a range into an Excel table (ListObject) named name. With header
    true the first row supplies the column names (deduplicated); style is a
    built-in style; row_stripes and col_stripes toggle banding. totals maps
    columns to a function (sum, average, count, min, max...). Refuses an
    overlap with an existing table, and a name already taken by a table or
    defined name. Hazardous workbooks refuse unless allow_loss is true.
    Auto-backup to .ks4xl-backups; atomic verified save. Refuses while open
    in Excel."""
    return _tables.create_table(
        path, location, name, header=header, style=style,
        row_stripes=row_stripes, col_stripes=col_stripes,
        totals_row=totals_row, totals=totals, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_op
def get_table(path: str, name: str,
              columns: _schemas.TableColumns = None,
              values: str = "cached", records: bool = False) -> dict:
    """Read a table's data by its name (case-insensitive). columns projects a
    subset; values is cached, formula, both, or displayed (the text Excel
    shows), and cached and both add shows where that text differs from the
    value; records true returns row objects keyed by column name. Returns
    the table ref, the column names, and the data rows without the header or
    totals row; page big tables with query_range and a {table} location.
    Table edits: manage_table. Read-only."""
    return _tables.get_table(path, name, columns=columns, values=values,
                             records=records)


@_op
def manage_table(path: str, name: str, action: str,
                 values: _schemas.TableValues = None,
                 index: int | None = None, column: str | None = None,
                 new_name: str | None = None, new_ref: str | None = None,
                 style: str | None = None, on: bool = True,
                 totals: dict | None = None, row_stripes: bool | None = None,
                 col_stripes: bool | None = None, sheet: str | None = None,
                 allow_loss: bool = False, backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Advanced table lifecycle on the table named name. action is one of:
    add_row (values as a row or list of rows), delete_row (index, 1-based data
    row; REMOVES that row's data), add_column (column name, optional values),
    delete_column (column; REMOVES that column's data), rename (new_name),
    resize (new_ref, keeping the top-left anchor), toggle_totals (on, optional
    per-column totals functions), to_range (drop the table shell, keep the
    data), set_style (style, row_stripes, col_stripes). Row and column inserts
    and deletes rewrite every reference through the core engine so formulas,
    names, and merges stay coherent, and the table ref is reset to its
    intended bounds. A hazardous workbook refuses unless allow_loss is true.
    Auto-backup: prev/anchor slots in .ks4xl-backups (backup=false skips
    rotation); atomic verified save, restored on failed verify; the prev slot
    is the undo for the destructive actions. Refuses while open in Excel."""
    return _tables.manage_table(
        path, name, action, values=values, index=index, column=column,
        new_name=new_name, new_ref=new_ref, style=style, on=on, totals=totals,
        row_stripes=row_stripes, col_stripes=col_stripes, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


# -------------------------------------------------------------- named ranges


@_op
def manage_name(path: str, action: str, name: str | None = None,
                refers_to: str | None = None, scope: str | None = None,
                new_name: str | None = None, allow_loss: bool = False,
                backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Manage defined names (named ranges). action is one of: add (name,
    refers_to as an A1 reference or a formula; a convenience leading '=' is
    stripped, definitions are stored bare), delete (name, optional scope),
    rename (name, new_name), update (name, refers_to), list (every name with
    its scope and refers-to; read-only, no backup). Names live at 'workbook'
    scope or a sheet-title scope, and a sheet-scoped name can shadow a
    workbook-scoped one, so an action on an ambiguous name refuses and returns
    both scopes until you pass scope. The reserved print-area, print-title,
    and filter built-ins (_xlnm.*) are protected from delete and rename. For
    the mutating actions: a hazardous workbook refuses unless allow_loss is
    true; auto-backup to prev/anchor slots in .ks4xl-backups (backup=false
    skips rotation); atomic verified save. Refuses while open in Excel."""
    return _names.manage_name(
        path, action, name=name, refers_to=refers_to, scope=scope,
        new_name=new_name, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# ----------------------------------------------------- conditional formatting


@_op
def manage_conditional_format(path: str, action: str,
                              location: _schemas.OptionalLocation = None,
                              cf_type: str | None = None,
                              params: dict | None = None,
                              sheet: str | None = None,
                              index: int | None = None,
                              allow_loss: bool = False, backup: bool = True,
                              verify_com: bool | None = None) -> dict:
    """Manage conditional-formatting rules. action is add, list (read-only:
    each rule's cf_type, operator, formula and colours, in add's own words),
    or delete. For add, location is the range and cf_type picks the rule with
    its params: cell_is {operator (greater_than, less_than, between, equal,
    not_equal, greater_than_or_equal, less_than_or_equal, not_between),
    formula (a value or [lo, hi]), fill, font_color}; formula {formula,
    fill}; color_scale {colors: 2 or 3 hex stops}; data_bar {color,
    show_value}; icon_set {icon_style, values, show_value, reverse};
    top_bottom {rank, percent, bottom, fill}; unique and duplicate {fill,
    font_color}; above_average {below, equal_average, std_dev 1-3, fill,
    font_color}; time_period {period: today | yesterday | tomorrow |
    last_7_days | this_week | last_week | next_week | this_month |
    last_month | next_month, fill}; blanks and errors {fill, font_color}.
    Excel's camelCase spellings (greaterThan, last7Days) are accepted too.
    For delete,
    location must EXACTLY match the rule's stored range (or one member of it)
    and an optional 0-based index picks one rule of several. Excel
    evaluates the
    rules when it opens the file, so a time_period rule follows the reader's
    clock. For the
    mutating actions: a hazardous workbook refuses unless allow_loss is true;
    auto-backup to prev/anchor slots in .ks4xl-backups; atomic verified save.
    Refuses while open in Excel."""
    return _condformat.manage_conditional_format(
        path, action, location=location, cf_type=cf_type, params=params,
        sheet=sheet, index=index, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# ------------------------------------------------------------ data validation


@_op
def manage_data_validation(path: str, action: str,
                           location: _schemas.OptionalLocation = None,
                           dv_type: str | None = None,
                           values: _schemas.DvValues = None,
                           operator: str | None = None,
                           formula1: _schemas.DvFormula = None,
                           formula2: _schemas.DvFormula = None,
                           allow_blank: bool = True, prompt: str | None = None,
                           error: str | None = None, sheet: str | None = None,
                           allow_loss: bool = False, backup: bool = True,
                           verify_com: bool | None = None) -> dict:
    """Manage data-validation rules. action is add, list (read-only), or
    delete. For add, location is the range and dv_type is one of list, whole,
    decimal, date, time, text_length, or custom. A list takes values (inline
    items, or a range/formula; inline lists over 255 characters draw a
    warning, point long lists at a range; an inline item cannot hold a
    comma); the numeric and date types take operator (between, not_between,
    equal, not_equal, greater_than, less_than, greater_than_or_equal,
    less_than_or_equal) plus formula1 and formula2 bounds; custom takes
    formula1. allow_blank (default true) lets empty entries pass; prompt and
    error set the input and error messages, and an invalid entry is refused
    with Excel's own alert when error is not given. Delete matches the
    rule's stored range exactly. The file stores the rule; Excel enforces it
    on entry. For the
    mutating actions: a hazardous workbook refuses unless allow_loss is true;
    auto-backup to prev/anchor slots in .ks4xl-backups; atomic verified save.
    Refuses while open in Excel."""
    return _datavalidation.manage_data_validation(
        path, action, location=location, dv_type=dv_type, values=values,
        operator=operator, formula1=formula1, formula2=formula2,
        allow_blank=allow_blank, prompt=prompt, error=error, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


# ------------------------------------------------------------ sort and filter


@_op
def sort_range(path: str, location: _schemas.Location, keys: _schemas.SortKeys,
               has_header: bool = True, sheet: str | None = None,
               allow_loss: bool = False, backup: bool = True,
               verify_com: bool | None = None) -> dict:
    """Sort a range or table body by one or more keys, writing the rows back
    reordered. keys is a list of {column, order}: header name, letter, or
    1-based index; asc or desc; later keys break ties. has_header keeps
    the first row put. Moved formulas shift relative refs (Excel
    semantics); keys compare cached values, warning when uncalculated.
    Filter-hidden rows stay pinned and unsorted, as in Excel. Hazardous
    workbooks need allow_loss. Auto-backup (prev is the undo); atomic
    verified save. Refuses while open in Excel."""
    return _sortfilter.sort_range(
        path, location, keys, has_header=has_header, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_op
def set_filter(path: str, location: _schemas.Location,
               criteria: _schemas.FilterCriteria = None, sheet: str | None = None,
               allow_loss: bool = False, backup: bool = True,
               verify_com: bool | None = None) -> dict:
    """Apply an autofilter over a range whose first row is the header, and
    actually hide the non-matching rows: an .xlsx stores filter CRITERIA, not
    hidden state, so criteria (a list of {column, op, value}, ops as in
    query_range, combined as AND) are evaluated here over cached and literal
    values; a row whose tested cell holds an uncalculated formula stays
    visible with a warning. A hazardous workbook refuses unless allow_loss is
    true. Auto-backup to .ks4xl-backups; atomic verified save. Refuses while
    open in Excel."""
    return _sortfilter.set_filter(
        path, location, criteria=criteria, sheet=sheet, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


@_op
def clear_filter(path: str, location: _schemas.OptionalLocation = None,
                 sheet: str | None = None, allow_loss: bool = False,
                 backup: bool = True, verify_com: bool | None = None) -> dict:
    """Remove the autofilter from a sheet and unhide the rows it hid, the
    reverse of set_filter. location or sheet picks the sheet; the sheet's
    active autofilter range is used when location is omitted. A sheet with no
    autofilter refuses (NOT_FOUND) rather than silently no-opping. Returns how
    many rows were unhidden. A hazardous workbook refuses unless allow_loss is
    true. Auto-backup: prev/anchor slots in .ks4xl-backups; atomic verified
    save. Refuses while open in Excel."""
    return _sortfilter.clear_filter(
        path, location=location, sheet=sheet, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


# ----------------------------------------------------- comments and hyperlinks


@_op
def manage_comment(path: str, action: str,
                   location: _schemas.OptionalLocation = None,
                   text: str | None = None, author: str | None = None,
                   sheet: str | None = None, allow_loss: bool = False,
                   backup: bool = True, verify_com: bool | None = None) -> dict:
    """Manage legacy cell comments (the sticky notes). action is add (location,
    text, optional author), edit (location, new text or author), delete
    (location), or list (read-only; every comment with its cell, text, and
    author). This writes legacy notes, which round-trip cleanly through a
    file-based save. Threaded reply-and-resolve comments are a different part
    this tool does not write: the hazard scan flags them and the mutation
    gate refuses rather than dropping them; list reports a threaded comment
    as threaded, with its author and text. Leave author out to sign with
    your own name (the Office user name, else the Windows sign-in name), or
    pass author="assistant" to sign as "AI Assistant via KS4XL", which
    KS4XL_AUTHOR=assistant makes the default; any other author is used as
    given. The answer names the author and its source. An email address is
    never written as an author: the default signs instead and the answer's
    warnings say so.
    For the mutating actions: a hazardous workbook
    refuses unless allow_loss is true; auto-backup to prev/anchor slots in
    .ks4xl-backups; atomic verified save. Refuses while open in Excel."""
    return _annotations.manage_comment(
        path, action, location=location, text=text, author=author,
        sheet=sheet, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


@_op
def manage_hyperlink(path: str, action: str,
                     location: _schemas.OptionalLocation = None,
                     target: str | None = None, display: str | None = None,
                     tooltip: str | None = None, sheet: str | None = None,
                     allow_loss: bool = False, backup: bool = True,
                     verify_com: bool | None = None) -> dict:
    """Manage cell hyperlinks. action is add (location, target as a URL or an
    in-workbook 'Sheet!A1' reference, optional display text and tooltip),
    remove (location), or list (read-only). On add, display replaces the
    cell's value; with no display an empty cell shows the target. remove
    strips only the link: the cell's text and style stay. On list it surfaces
    both real cell hyperlinks and HYPERLINK() formula links so an audit sees
    every kind. For the mutating actions: a hazardous workbook refuses unless
    allow_loss is true; auto-backup to prev/anchor slots in .ks4xl-backups;
    atomic verified save. Refuses while open in Excel."""
    return _annotations.manage_hyperlink(
        path, action, location=location, target=target, display=display,
        tooltip=tooltip, sheet=sheet, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# ------------------------------------------------------------ import / export


@_op
def import_data(path: str, source: str | None = None,
                source_file: str | None = None, fmt: str = "auto",
                location: _schemas.OptionalLocation = None,
                sheet: str | None = None, header: bool = True,
                delimiter: str | None = None, encoding: str = "utf-8",
                formulas: bool = False, column_types: dict | None = None,
                parse_dates: bool = False, echo_cells: bool = False,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Import CSV, TSV or JSON at location (default A1) from source (text) or
    source_file (a path); the format comes from the extension and a CSV's
    delimiter from its lines (a semicolon file reads decimal commas). Codes
    stay text: leading-zero, E-notation and over-15-digit columns are kept
    as text and reported; column_types sets text, number, date or auto per
    column; parse_dates=true reads ISO dates. Text starting =, +, - or @
    stays text unless formulas=true. Auto-backup, verified save; refuses
    while open in Excel."""
    with _package.echo_cells(echo_cells):
        return _dataio.import_data(
            path, source=source, source_file=source_file, fmt=fmt,
            location=location, sheet=sheet, header=header, delimiter=delimiter,
            encoding=encoding, formulas=formulas, column_types=column_types,
            parse_dates=parse_dates, allow_loss=allow_loss,
            backup=backup, verify_com=verify_com)


@_op
def export_range(path: str, location: _schemas.OptionalLocation = None,
                 sheet: str | None = None, fmt: str = "csv",
                 header: bool = True, values: str = "cached",
                 records: bool = False, out_file: str | None = None,
                 overwrite: bool = False, bom: bool = True) -> dict:
    """Export a range, table or sheet (location; default the used range) to
    CSV, TSV or JSON. values is cached, formula, both or displayed (what
    each number format shows, as Excel's CSV does: 00000 keeps 02134).
    CSV/TSV write TRUE/FALSE and plain dates as Excel does; a written file
    starts with Excel's UTF-8 byte-order mark unless bom=false. out_file
    writes a file (never the workbook or .ks4xl-backups; overwrite keeps a
    .bak), else the text returns inline. Whole sheets: export_file.
    Read-only."""
    return _dataio.export_range(
        path, location=location, sheet=sheet, fmt=fmt, header=header,
        values=values, records=records, out_file=out_file,
        overwrite=overwrite, bom=bom)


# ====================================== STRUCTURE, MERGES, FIND AND STYLES
# Ungated file-tier families on the audited engine: the
# structural-edit flagship (modify_grid_structure over core.refs), merges,
# workbook-wide find/replace, the scatter cell pair, and the style layer
# (named styles, format painter, style-bloat audit). Every mutation runs
# through WorkbookPackage.


# ---------------------------------------------------------- grid structure


@_op
def modify_grid_structure(path: str, action: str, at: _schemas.GridPosition,
                          count: int = 1, sheet: str | None = None,
                          allow_loss: bool = False, backup: bool = True,
                          verify_com: bool | None = None) -> dict:
    """Insert or delete rows or columns at a position and REWRITE EVERY
    REFERENCE so the workbook stays coherent: formulas on every sheet
    (cross-sheet refs included), defined names, data validations,
    conditional-format ranges, table refs, and merged ranges all shift with
    the edit. action is insert_rows, delete_rows, insert_cols, or
    delete_cols; at is the 1-based row number or column letter where the
    edit starts (a cell like 'B7' or a location object also works, using
    its top-left); count edits that many at once.

    Whole-column spans like =SUM(B:B) and whole-row spans like $1:$2 shift
    on their own axis; an edit on the other axis leaves them alone (Excel's
    behavior). A reference wholly inside a deleted band becomes #REF! and
    the new-#REF! count is reported, never hidden; an insert that would
    push value-bearing cells off the grid edge refuses. Returns per-kind
    rewrite counts. A hazardous workbook refuses unless allow_loss is true.
    Auto-backup: prev/anchor slots in .ks4xl-backups (backup=false skips
    rotation); atomic verified save, restored on failed verify; the prev
    slot is the undo for a delete. Refuses while open in Excel."""
    return _structure.modify_grid_structure(
        path, action, at, count=count, sheet=sheet, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


# ------------------------------------------------------------------- merges


@_op
def set_merge(path: str, action: str,
              location: _schemas.OptionalLocation = None,
              sheet: str | None = None, confirm_data_loss: bool = False,
              allow_loss: bool = False, backup: bool = True,
              verify_com: bool | None = None) -> dict:
    """Merge or unmerge cell ranges, or list every merge. action is merge
    (location is the multi-cell range), unmerge (location must be the exact
    stored merged range), or list (read-only, one sheet or the whole
    workbook). Excel merge semantics: only the top-left value survives, so
    a merge whose absorbed cells hold values REFUSES until you pass
    confirm_data_loss=true, then reports exactly which values were
    discarded; overlapping an existing merge refuses. Unmerge keeps the
    surviving top-left value and leaves the rest blank. For the mutating
    actions: a hazardous workbook refuses unless allow_loss is true;
    auto-backup to prev/anchor slots in .ks4xl-backups; atomic verified
    save. Refuses while open in Excel."""
    return _cells.set_merge(
        path, action, location=location, sheet=sheet,
        confirm_data_loss=confirm_data_loss, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


# ----------------------------------------------------------- find / replace


@_op
def find_cells(path: str, query: str, look_in: str = "values",
               match: str = "contains", match_case: bool = False,
               sheet: str | None = None,
               location: _schemas.OptionalLocation = None, limit: int = 100,
               offset: int = 0) -> dict:
    """Search cell values and/or formulas across a workbook, sheet, or range and
    return EVERY match with its unambiguous address. match is exact,
    contains, or regex (timeout-guarded, so a pathological pattern refuses
    instead of hanging); look_in is values, formulas, or both; a formula
    cell's searchable value is its last cached one, and a value Excel shows
    as other text carries that text in shows. Results page with limit and
    offset and report the total match count. Read-only; works while the file
    is open in Excel."""
    return _search.find_cells(
        path, query, look_in=look_in, match=match, match_case=match_case,
        sheet=sheet, location=location, limit=limit, offset=offset)


@_op
def replace_cells(path: str, find: str, replace: str, look_in: str = "values",
                  match: str = "contains", match_case: bool = False,
                  sheet: str | None = None,
                  location: _schemas.OptionalLocation = None,
                  dry_run: bool = False, formulas: bool = False,
                  allow_loss: bool = False, backup: bool = True,
                  verify_com: bool | None = None) -> dict:
    """Find-and-replace across a workbook, sheet, or range. match is exact
    (whole cell), contains (literal substring), or regex (timeout-guarded;
    backreferences like \\1 work in replace). look_in 'values' rewrites
    literal cells, 'formulas' rewrites formula text (the cell stays a
    formula, normalized), 'both' does both. dry_run=true previews every
    change without touching the file; a real run validates the whole plan
    first, applies it as ONE batch, and reports cells changed and
    occurrences replaced. A replaced value that parses as a number is
    written as a number; replacement text beginning with =, +, -, or @ is
    written as TEXT to block formula injection unless formulas is true. A
    hazardous workbook refuses unless allow_loss is true. Auto-backup to
    .ks4xl-backups; atomic verified save. Refuses while open in Excel."""
    return _search.replace_cells(
        path, find, replace, look_in=look_in, match=match,
        match_case=match_case, sheet=sheet, location=location,
        dry_run=dry_run, formulas=formulas, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


# ------------------------------------------------------------ scatter cells


@_op
def get_cells(path: str, cells: _schemas.CellAddresses, values: str = "cached",
              sheet: str | None = None) -> dict:
    """Read many individually addressed cells in one call, the scatter complement
    to the rectangular read_range. cells is a list of A1 strings or location
    objects, each resolving to ONE cell (1,000-cell ceiling); values is
    cached, formula, both, or displayed (the text Excel shows), labelled and
    with shows as in read_range, so a formula with no cached value is never
    passed off as blank. Read-only; works while the file is open in Excel."""
    return _cells.get_cells(path, cells, values=values, sheet=sheet)


@_op
def set_cells(path: str, cells: _schemas.CellWrites, sheet: str | None = None,
              echo_cells: bool = False, allow_loss: bool = False,
              backup: bool = True, verify_com: bool | None = None) -> dict:
    """Write many individually addressed cells as ONE atomic batch, the
    scatter complement to write_range. cells is a list of {cell, value}
    items (cell is an A1 string or a location object resolving to one cell;
    1,000-cell ceiling); every address is resolved BEFORE anything is
    written, so one bad item refuses the whole batch untouched. '=' strings
    become formulas, normalized. A hazardous workbook refuses unless
    allow_loss is true. Auto-backup to .ks4xl-backups; atomic verified
    save. Refuses while open in Excel."""
    with _package.echo_cells(echo_cells):
        return _cells.set_cells(path, cells, sheet=sheet, allow_loss=allow_loss,
                                backup=backup, verify_com=verify_com)


# ------------------------------------------------------------- style layer


@_op
def apply_style(path: str, style: str,
                location: _schemas.OptionalLocation = None,
                define: dict | None = None, sheet: str | None = None,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Apply a NAMED cell style to a range, optionally defining it first.
    style is a workbook style or an Excel builtin (Good, Bad, Input, Title,
    Total...); define ({number_format, font, fill, border, alignment})
    registers a new named style under that name (with no location it only
    registers). Named styles keep repeated formatting out of the
    64,000-format registry. A hazardous workbook refuses unless allow_loss
    is true. Auto-backup to .ks4xl-backups; atomic verified save. Refuses
    while open in Excel."""
    return _format.apply_style(
        path, style, location=location, define=define, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_op
def copy_format(path: str, source: _schemas.SourceLocation,
                dest: _schemas.DestLocation, sheet: str | None = None,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """The format painter: copy ONE source cell's complete format (font,
    fill, border, alignment, number format) onto every cell of a
    destination range, leaving values untouched. source must resolve to a
    single cell; dest is any range, cross-sheet allowed; both are location
    objects. A hazardous workbook refuses unless allow_loss is true.
    Auto-backup to .ks4xl-backups; atomic verified save. Refuses while open
    in Excel."""
    return _format.copy_format(path, source, dest, sheet=sheet,
                               allow_loss=allow_loss, backup=backup,
                                   verify_com=verify_com)


@_op
def audit_styles(path: str, top: int = 10) -> dict:
    """Read-only style-bloat audit against Excel's 64,000 distinct-format
    ceiling: counts from the styles.xml registry (cell formats, fonts,
    fills, borders, custom number formats, named styles), the formats
    actually in use per sheet, the `top` heaviest formats by cell count
    with a compact description, and an ok / elevated / critical risk
    verdict with its thresholds. A large registry-vs-in-use gap means
    orphaned entries left by past edits. Nothing is written; works while
    the file is open in Excel."""
    return _format.audit_styles(path, top=top)


# =========================================== OBJECTS, PROTECTION, INSPECTORS
# Ungated file-tier families: objects (images and charts on the
# openpyxl model, with the observed-round-trip gate story), protection, page
# layout and headers/footers, the read-side inspectors (external links, VBA,
# pivots, connections), and the multi-sheet export_file. Sparklines are an x14
# in-sheet extension openpyxl does not model; that tool is deferred to the COM
# tier, not faked here.


# ------------------------------------------------------------------ objects


@_op
def manage_image(path: str, action: str, image_file: str | None = None,
                 location: _schemas.OptionalLocation = None,
                 sheet: str | None = None, width: int | None = None,
                 height: int | None = None, index: int | None = None,
                 out_dir: str | None = None, allow_loss: bool = False,
                 backup: bool = True, verify_com: bool | None = None) -> dict:
    """Manage cell-anchored images. action is insert (image_file, location as
    the cell the image sits on, optional width and height in pixels), list
    (read-only: every image with its sheet, 1-based index, at_cell, size, and
    format), delete (sheet plus index from list, or location as its cell), or extract
    (write every embedded image into out_dir; read-only).

    Round-trip honesty: images survive a file-based save because Pillow is
    installed and the workbook's drawings hold pictures only; the mutating
    actions verify that precondition, downgrade the usual hazard refusal to
    a warning, and verify-after-write still refuses the save if an image
    part is in fact lost. A workbook whose drawings hold real shapes
    (textboxes, controls) refuses unless allow_loss is true, because those
    shapes die on a file-based save. After every mutation the image count
    is read back and a mismatch restores the backup. Auto-backup to
    .ks4xl-backups; atomic verified save. Refuses while open in Excel."""
    return _objects.manage_image(
        path, action, image_file=image_file, location=location, sheet=sheet,
        width=width, height=height, index=index, out_dir=out_dir,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_op
def manage_chart(path: str, action: str, chart_type: str | None = None,
                 data: _schemas.ChartRange = None,
                 categories: _schemas.ChartRange = None,
                 title: str | None = None, x_title: str | None = None,
                 y_title: str | None = None, at_cell: str | None = None,
                 sheet: str | None = None, index: int | None = None,
                 titles_from_data: bool = True, show_axes: bool = True,
                 allow_loss: bool = False, backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Create, list, and delete charts. action create takes chart_type (bar,
    bar_horizontal, line, pie, doughnut, area, scatter), data (a location
    object; its first row supplies series titles unless titles_from_data is
    false, and scatter uses the first column as x values), optional
    categories, title, x_title, y_title, and at_cell, the cell the chart's
    top-left corner sits on (default just right of the data; the old name
    anchor is still accepted). list reports each chart's sheet, 1-based
    index, chart_type, title, at_cell and the ranges each series plots;
    delete takes sheet plus index or title.

    Axes are drawn by default. Pass show_axes false for a bare plot area
    with no category or value axis; pie and doughnut have no axes either
    way.

    Fidelity honesty: chart fidelity is model-mediated. openpyxl
    re-serializes every chart through its own model on save, so a complex
    Excel-authored chart can lose sub-features the model does not know;
    the hazard scan tracks charts as a degrade risk and warns on every
    save. Existing charts are not editable here; delete and recreate is
    the edit route. After create or delete the chart count is read
    back and a mismatch restores the backup. Auto-backup to .ks4xl-backups;
    atomic verified save. Refuses while open in Excel."""
    return _objects.manage_chart(
        path, action, chart_type=chart_type, data=data,
        categories=categories, title=title, x_title=x_title,
        y_title=y_title, anchor=at_cell, sheet=sheet, index=index,
        titles_from_data=titles_from_data, show_axes=show_axes,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


# --------------------------------------------------------------- protection


@_op
def set_protection(path: str, action: str, sheet: str | None = None,
                   password: str | None = None, options: dict | None = None,
                   unlock_ranges: _schemas.RangeList = None, locked: bool = False,
                   structure: bool = True, windows: bool = False,
                   scope: str = "all", allow_loss: bool = False,
                   backup: bool = True,
                   verify_com: bool | None = None) -> dict:
    """Advisory workbook and sheet protection. action is sheet (protect one
    sheet; options maps per-option names like format_cells, insert_rows,
    sort, auto_filter to true = allowed while protected; optional
    password; unlock_ranges leaves those cells editable), workbook (lock
    the structure and optionally windows, optional password), unlock (mark
    unlock_ranges as exceptions that stay editable under protection; locked
    true relocks them), remove (scope sheet, workbook, or all), or status
    (read-only report of every lock and each sheet's unlocked ranges). An
    argument the action does not use is refused. Every other write honours
    the protection: a locked cell, or a change the sheet's options do not
    allow, refuses PROTECTED.

    Honesty: xlsx protection is tamper DISCOURAGEMENT, not security. The
    password is stored as the legacy ECMA-376 hash Excel checks in its UI;
    the file stays a readable zip and any tool can strip the lock. Real
    encryption is a different mechanism and needs the Excel application
    (com_save_with_password, xl_app). Auto-backup to
    .ks4xl-backups; atomic verified save. Refuses while open in Excel."""
    return _protection.set_protection(
        path, action, sheet=sheet, password=password, options=options,
        unlock_ranges=unlock_ranges, locked=locked, structure=structure,
        windows=windows, scope=scope, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# -------------------------------------------------------------- page layout


@_op
def set_page_layout(path: str, sheet: str | None = None,
                    orientation: str | None = None,
                    paper_size: _schemas.PaperSize = None,
                    margins: dict | None = None, scale: int | None = None,
                    fit_to_width: int | None = None,
                    fit_to_height: int | None = None,
                    print_area: str | None = None,
                    print_title_rows: str | None = None,
                    print_title_cols: str | None = None,
                    gridlines: bool | None = None,
                    headings: bool | None = None, allow_loss: bool = False,
                    backup: bool = True, verify_com: bool | None = None) -> dict:
    """Set the print-shaped page settings in one call; every parameter is
    optional and unset ones keep their current values. orientation is
    portrait or landscape; paper_size is a name (letter, legal, tabloid,
    executive, a3, a4, a5, b4, b5, ledger) or an ECMA-376 code; margins
    maps left, right, top, bottom, header, footer to inches; scale (10-400
    percent) and the fit_to_width/fit_to_height page counts are mutually
    exclusive; print_area takes an A1 range or 'clear'; print_title_rows
    ('1:2') and print_title_cols ('A:B') repeat those spans on every
    printed page; gridlines and headings toggle their print flags.
    Auto-backup to .ks4xl-backups; atomic verified save. Refuses while
    open in Excel."""
    return _pagelayout.set_page_layout(
        path, sheet=sheet, orientation=orientation, paper_size=paper_size,
        margins=margins, scale=scale, fit_to_width=fit_to_width,
        fit_to_height=fit_to_height, print_area=print_area,
        print_title_rows=print_title_rows, print_title_cols=print_title_cols,
        gridlines=gridlines, headings=headings, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


@_op
def set_header_footer(path: str, sheet: str | None = None,
                      header: dict | None = None, footer: dict | None = None,
                      apply_to: str = "odd", raw: bool = False,
                      allow_loss: bool = False, backup: bool = True,
                      verify_com: bool | None = None) -> dict:
    """Write the three-section page headers and footers. header and footer
    are {left, center, right} dicts; an empty string clears a section.
    Text is literal by default: & is escaped and the placeholders {page}
    {pages} {date} {time} {file} {sheet} {path} expand to Excel codes;
    raw=true passes & codes through verbatim. apply_to targets odd
    (default), even, or first pages, setting the matching flag. Auto-backup
    to .ks4xl-backups; atomic verified save. Refuses while open in
    Excel."""
    return _pagelayout.set_header_footer(
        path, sheet=sheet, header=header, footer=footer, apply_to=apply_to,
        raw=raw, allow_loss=allow_loss, backup=backup, verify_com=verify_com)


# ---------------------------------------------------- read-side inspectors


@_op
def get_external_links(path: str) -> dict:
    """List external workbook links, read-only: each link's target path from
    its relationship, the sheet names it references, and the cached cell
    values Excel stored at the last refresh (capped per link, with the
    total count and a truncated flag). Cached values are copies of data
    from OTHER workbooks, so treat the output as sensitive. File-based
    edits preserve external links, making this an audit tool; breaking or
    repointing links is not offered at the file tier. Works while the file
    is open in Excel."""
    return _inspectors.get_external_links(path)


@_op
def inspect_vba(path: str) -> dict:
    """Read-only VBA inspection: whether the workbook carries a
    vbaProject.bin, its size and declared content type, and the module
    names, types, and stream sizes extracted from the project's dir stream
    when the container parses as a well-formed compound file. A container
    that does not parse degrades honestly to presence plus size with a
    note saying so. This server never writes or runs VBA; macro authoring
    and execution stay outside the file tier. Works while the file is open
    in Excel."""
    return _inspectors.inspect_vba(path)


@_op
def get_pivot(path: str, sheet: str | None = None) -> dict:
    """Describe existing pivot tables, read-only: each pivot's sheet and
    location, source sheet and range, cache field list, row, column, data,
    and page fields by name, cache record count, who refreshed it last and
    when, and the refresh-on-load flag. Description only, stated honestly:
    creating, refreshing, or deleting a pivot needs the Excel application
    (com_manage_pivot, xl_app), while file-based edits elsewhere in the
    workbook preserve the pivot parts. Works while the file is open in
    Excel."""
    return _inspectors.get_pivot(path, sheet=sheet)


@_op
def get_connections(path: str) -> dict:
    """List data connections and Power Query presence, read-only: each
    connection's name, type, connection string or URL, command, and
    refresh-on-load flag from the connections part, plus DataMashup
    evidence for Power Query and any legacy query-table parts. Connection
    strings can embed server names and credentials, so the output is
    marked sensitive. This server does not refresh a data connection;
    pivot refresh lives in xl_app com_manage_pivot. Works while the file is
    open in Excel."""
    return _inspectors.get_connections(path)


# -------------------------------------------------------------- export_file


@_op
def export_file(path: str, fmt: str = "csv", sheets: _schemas.SheetNames = None,
                out_dir: str | None = None, out_file: str | None = None,
                header: bool = True, values: str = "cached",
                records: bool = False, overwrite: bool = False,
                bom: bool = True) -> dict:
    """Export whole sheets (sheets picks some) to CSV/TSV files in out_dir,
    one per sheet, or to one JSON bundle keyed by sheet (out_file;
    records=true gives row objects); with no target the text returns
    inline. values is cached, formula, both or displayed. CSV/TSV files
    start with a UTF-8 byte-order mark unless bom=false. A target is never
    the workbook or .ks4xl-backups; an existing file needs overwrite=true
    and keeps a .bak. Read-only."""
    return _dataio.export_file(
        path, fmt=fmt, sheets=sheets, out_dir=out_dir, out_file=out_file,
        header=header, values=values, records=records, overwrite=overwrite,
        bom=bom)


# ===================================== FORMULAS, VALIDATE, WORKFLOWS, BACKUPS
# Ungated file-tier tools on the audited engine: the formula pair (write with
# fill, the read-side audit), the multiplex validate checker, workflow
# recipes, backup management over the safesave slots, document properties
# with the calc-settings surface, and sheet-view state. Every mutation runs
# through WorkbookPackage. The COM tier registers below under the env-gated
# com pack.


# ----------------------------------------------------------------- formulas


@_op
def set_formula(path: str, location: _schemas.Location, formula: str,
                sheet: str | None = None, echo_cells: bool = False,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Write a formula to a single cell, or fill a range where each cell gets
    the formula with its relative references shifted by that cell's offset
    (Excel copy semantics; absolute $ anchors stay put). Formulas are
    normalized so modern functions do not land as #NAME?, and the workbook
    is flagged to recalculate on its next open; stored cached results stay
    stale until then (the recalculate op in xl_app populates them).
    Auto-backup to .ks4xl-backups; atomic verified save. Refuses while open
    in Excel."""
    with _package.echo_cells(echo_cells):
        return _formulas.set_formula(path, location, formula, sheet=sheet,
                                     allow_loss=allow_loss, backup=backup,
                                     verify_com=verify_com)


@_op
def audit_formulas(path: str, location: _schemas.OptionalLocation = None,
                   sheet: str | None = None) -> dict:
    """Read-only formula intelligence for a range, one sheet (sheet alone), or
    the whole workbook (no scope): the formula list plus five safety
    reports. external_references flags formulas reaching into other
    workbooks; volatile lists always-recalculating functions (NOW, RAND,
    OFFSET, INDIRECT and kin); missing_cached_values names formula cells
    with NO stored result, which read as blank to every non-Excel consumer
    until a recalculation; error_cells catches #REF!, #NAME?, and the other
    error literals in cached results or formula text; and
    cross_sheet_dependencies summarizes which sheets' formulas depend on
    which. Each list is capped with exact counts and a truncated flag. Use
    it before and after structural edits, and to judge whether cached
    values can be trusted. Read-only; works while the file is open in
    Excel."""
    return _formulas.audit_formulas(path, location=location, sheet=sheet)


# --------------------------------------------------- validation and workflow


@_op
def validate(path: str, checks: list[str] | None = None) -> dict:
    """Run read-only correctness checks and return one report. checks (default
    ['structure', 'references', 'calc_staleness']): structure (package
    opens clean, sheet integrity), references (#REF!/#NAME? and the other
    error cells), names (broken defined names), merges (overlapping or
    orphaned merged ranges), tables (duplicate names, broken refs,
    overlaps), formatting_bloat (the audit_styles counters against the
    64,000-format ceiling), hazards (the round-trip scan as a check),
    external_links (links reported, not repaired), calc_staleness (formulas
    lacking cached values, which read as blank outside Excel). Returns
    {passed, results: {check: {passed, findings}}}; findings keep the
    underlying ops' shapes where those exist, and passed=false means
    findings, not a failed call. Read-only, always; repairs live in the
    editing tools. Works while the file is open in Excel."""
    return _validation.validate(path, checks=checks)


@_op
def get_workflows(task: str | None = None) -> dict:
    """Recommended tool sequences for common multi-step spreadsheet tasks, each
    step naming the tool and op to call and the rationale. Call with no task
    to list the available tasks (recover-workbook, merge-workbooks,
    report-build, data-cleanup, formatting-audit-and-fix,
    safe-edit-of-rich-workbook, migrate-from-incumbent, migrate-from-v1);
    call with task='<name>' for that task's step-by-step recipe and notes.
    migrate-from-v1 carries the full map from every 1.x tool to its 2.0
    tool and op. Pure guidance: reads nothing, changes nothing."""
    return _workflows.get_workflows(task)


# ------------------------------------------------------------------ backups


@_op
def manage_backups(action: str, path: str | None = None,
                   directory: str | None = None, source: str | None = None,
                   scope: str | None = None, dry_run: bool = True,
                   label: str | None = None,
                   dest_dir: str | None = None) -> dict:
    """Manage the automatic backups in the hidden .ks4xl-backups folder next to
    each edited workbook: prev (the state before the latest edit) and anchor
    (session start). action='list': slots, snapshots and orphaned slot
    folders, for path (one workbook) or directory. action='restore':
    overwrite path with source 'prev', 'anchor' or a snapshot label; the
    current content rotates into prev first, so a restore is undoable, the
    copy is validated as a workbook first, and a file open in Excel refuses.
    Deleting a sheet, rows, columns or table rows, clearing cells, a merge
    that discards values, a replace and a protection change also keep a copy
    labelled "pre-<op> <time>", named in their result as snapshot (newest 20
    kept). action='purge': delete backups; scope 'orphans' (folders whose
    workbook is gone) or 'slots' (one workbook's pair); dry_run defaults to
    TRUE. action='snapshot': a permanent DTG-stamped copy,
    YYYYMMDD_HHMM_<name>, optional label and dest_dir; never rotated or
    purged. Damage after the last save (a crash, another program) is not in
    prev; only a snapshot covers it. Lost or corrupt file? get_workflows
    task='recover-workbook'."""
    return _backups.manage_backups(
        action, path=path, directory=directory, source=source, scope=scope,
        dry_run=dry_run, label=label, dest_dir=dest_dir)


# ------------------------------------------------------- properties and view


@_op
def set_workbook_properties(path: str, title: str | None = None,
                            author: str | None = None,
                            subject: str | None = None,
                            keywords: str | None = None,
                            category: str | None = None,
                            comments: str | None = None,
                            calc_mode: str | None = None,
                            full_calc_on_load: bool | None = None,
                            iterative_calc: bool | None = None,
                            max_iterations: int | None = None,
                            max_change: float | None = None,
                            allow_loss: bool = False,
                            backup: bool = True,
                            verify_com: bool | None = None) -> dict:
    """Set core document properties (title, author, subject, keywords,
    category, comments) and calc settings: calc_mode ('auto',
    'auto_no_table', 'manual'), full_calc_on_load, iterative_calc with
    max_iterations and max_change. Only given parameters change; with none
    it reports current values read-only. An email address is never written
    as the author; your own name is set instead. Manual mode skips recalc on
    open, so caches go stale; the result says so. Auto-backup; atomic
    verified save. Refuses while open in Excel."""
    return _properties.set_workbook_properties(
        path, title=title, author=author, subject=subject, keywords=keywords,
        category=category, comments=comments, calc_mode=calc_mode,
        full_calc_on_load=full_calc_on_load, iterative_calc=iterative_calc,
        max_iterations=max_iterations, max_change=max_change,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_op
def set_view(path: str, sheet: str | None = None, freeze: str | None = None,
             split: dict | None = None, gridlines: bool | None = None,
             headings: bool | None = None, zoom: int | None = None,
             selection: str | None = None, tab_color: str | None = None,
             allow_loss: bool = False, backup: bool = True,
             verify_com: bool | None = None) -> dict:
    """Set sheet-view state in one call: freeze panes (freeze='B2' locks the
    rows above and columns left of it; 'clear' removes), split panes
    (split={x, y} positions in points, exclusive with freeze), gridlines
    and headings visibility, zoom (10 to 400 percent), the active selection
    (an A1 cell or range), and the sheet tab color (6-digit hex or
    'clear'). Unset parameters keep their current values; sheet defaults to
    the active sheet. Auto-backup to .ks4xl-backups; atomic verified save.
    Refuses while open in Excel."""
    return _sheetview.set_view(
        path, sheet=sheet, freeze=freeze, split=split, gridlines=gridlines,
        headings=headings, zoom=zoom, selection=selection,
        tab_color=tab_color, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# ============================================================== COM TIER
# The environment-gated com pack: drives a real, private, hidden Excel
# instance for everything the file tier honestly cannot do. Serialized
# through one COM worker thread, alert-suppressed, timeout-bounded, PID-
# journaled; NEVER attaches to the user's Excel session and refuses files
# Excel holds open. VBA stays preserve/inspect only: no macro authoring and
# no macro execution (com_run_macro is deferred from v1).


@_op
def recalculate(path: str, engine: str = "auto",
                timeout_seconds: float | None = None,
                backup: bool = True, allow_network: bool = False) -> dict:
    """Recalculate in a private hidden Excel: engine='com' (or 'auto' where
    Excel exists) runs an explicit CalculateFull, saves, reports how many
    formula cells gained cached values, and names any circular reference.
    engine is 'auto' or 'com'; 'formulas' was removed in 2.0.0 and is refused.
    Auto-backup; refuses while the file is open in Excel.
    allow_network: Let Excel contact the internet for this workbook's
    network functions. Set it only after the user has agreed, for this
    workbook."""
    return _comtier.recalculate(path, engine=engine,
                                timeout_seconds=timeout_seconds,
                                backup=backup, allow_network=allow_network)


@_op
def com_manage_pivot(path: str, action: str, name: str | None = None,
                     source_sheet: str | None = None,
                     source_range: str | None = None,
                     dest_sheet: str | None = None, dest_cell: str = "A3",
                     rows: list[str] | None = None,
                     columns: list[str] | None = None,
                     filters: list[str] | None = None,
                     values: list[dict] | None = None,
                     timeout_seconds: float | None = None,
                     backup: bool = True, allow_network: bool = False) -> dict:
    """Create, refresh, delete, or list REAL pivot tables through Excel
    (never a fake static table). action='create': source_range must include
    a header row; rows/columns/filters name header fields; values is a list
    like [{'field': 'Amount', 'func': 'sum', 'caption'?}] with funcs sum,
    count, average, max, min, product, count_numbers, stdev, var;
    dest_sheet is created if missing (omitted: a new sheet), dest_cell
    defaults to A3, name is optional. action='refresh': one pivot by name,
    or every pivot when name is omitted; source-range changes are picked
    up. action='delete': clears the named pivot's range. action='list':
    read-only inventory with source and location. Runs in a private hidden
    Excel instance, serialized and timeout-bounded (timeout_seconds,
    default 60); mutations auto-backup and verify; Excel's own error text
    is surfaced when it refuses (a wrong field name, an invalid source).
    Refuses while the file is open in your Excel.
    allow_network: Let Excel contact the internet for this workbook's
    network functions. Set it only after the user has agreed, for this
    workbook."""
    return _comtier.com_manage_pivot(
        path, action, name=name, source_sheet=source_sheet,
        source_range=source_range, dest_sheet=dest_sheet,
        dest_cell=dest_cell, rows=rows, columns=columns, filters=filters,
        values=values, timeout_seconds=timeout_seconds, backup=backup,
        allow_network=allow_network)


@_op
def com_export_pdf(path: str, output: str, scope: str = "workbook",
                   sheet: str | None = None, range_a1: str | None = None,
                   overwrite: bool = False,
                   timeout_seconds: float | None = None,
                   allow_network: bool = False) -> dict:
    """Export to PDF via Excel's real renderer. scope='workbook' (every
    sheet), 'sheet' (one sheet, name via sheet), or 'range' (sheet plus
    range_a1 like 'A1:F40'). output must be a .pdf path; an existing file
    refuses unless overwrite:true. Read-only on the workbook (nothing is
    saved); the produced PDF is checked non-empty. Runs in a private
    hidden Excel.
    allow_network: Let Excel contact the internet for this workbook's
    network functions. Set it only after the user has agreed, for this
    workbook."""
    return _comtier.com_export_pdf(
        path, output, scope=scope, sheet=sheet, range_a1=range_a1,
        overwrite=overwrite, timeout_seconds=timeout_seconds,
        allow_network=allow_network)


@_op
def com_render_sheet(path: str, output: str, sheet: str | None = None,
                     range_a1: str | None = None, overwrite: bool = False,
                     timeout_seconds: float | None = None,
                     allow_network: bool = False) -> dict:
    """Render a sheet's used range (or range_a1) to a PNG image exactly as
    Excel displays it: formatting, conditional formats, charts in range,
    and sparklines included. output must be a .png path; existing files
    refuse unless overwrite:true. Read-only on the workbook, in a private
    hidden Excel. A picture that comes back blank is a refusal naming
    com_export_pdf, never an ok.
    allow_network: Let Excel contact the internet for this workbook's
    network functions. Set it only after the user has agreed, for this
    workbook."""
    return _comtier.com_render_sheet(
        path, output, sheet=sheet, range_a1=range_a1, overwrite=overwrite,
        timeout_seconds=timeout_seconds, allow_network=allow_network)


@_op
def com_convert_format(path: str, output: str, format: str | None = None,
                       overwrite: bool = False,
                       timeout_seconds: float | None = None,
                       allow_network: bool = False) -> dict:
    """Convert a workbook between formats via Excel's own SaveAs: xlsx,
    xlsm, xlsb, xls, csv (UTF-8; first worksheet's values only, by the
    format's nature), or ods. format defaults from the output extension;
    the source file is never overwritten. The produced file is checked
    non-empty. Runs in a private hidden Excel; refuses if the source is
    locked.
    allow_network: Let Excel contact the internet for this workbook's
    network functions. Set it only after the user has agreed, for this
    workbook."""
    return _comtier.com_convert_format(
        path, output, format=format, overwrite=overwrite,
        timeout_seconds=timeout_seconds, allow_network=allow_network)


@_op
def com_save_with_password(path: str, password: str,
                           current_password: str | None = None,
                           timeout_seconds: float | None = None,
                           backup: bool = True,
                           allow_network: bool = False) -> dict:
    """Password-protect a workbook with Excel's REAL encryption (the whole
    package is encrypted, unlike the advisory set_protection). password=''
    with current_password removes it; current_password also opens an
    already-encrypted file for re-keying. The password is NOT recoverable,
    and every file-based tool refuses encrypted files, so keep the backup.
    Runs in a private hidden Excel.
    allow_network: Let Excel contact the internet for this workbook's
    network functions. Set it only after the user has agreed, for this
    workbook."""
    return _comtier.com_save_with_password(
        path, password, current_password=current_password,
        timeout_seconds=timeout_seconds, backup=backup,
        allow_network=allow_network)


@_op
def com_autofit(path: str, sheet: str | None = None,
                columns: str | None = None, rows: str | None = None,
                timeout_seconds: float | None = None,
                backup: bool = True, allow_network: bool = False) -> dict:
    """True column and row autofit using Excel's real text metrics (the
    file tier can only estimate widths). Default: every used column on the
    sheet; columns like 'A:D' or 'C', rows like '1:20'. Auto-backup;
    verified save; runs in a private hidden Excel instance, serialized and
    timeout-bounded; refuses while the file is open in Excel.
    allow_network: Let Excel contact the internet for this workbook's
    network functions. Set it only after the user has agreed, for this
    workbook."""
    return _comtier.com_autofit(
        path, sheet=sheet, columns=columns, rows=rows,
        timeout_seconds=timeout_seconds, backup=backup,
        allow_network=allow_network)


@_op
def com_goal_seek(path: str, target_cell: str, target_value: float,
                  changing_cell: str, sheet: str | None = None,
                  save: bool = True, timeout_seconds: float | None = None,
                  backup: bool = True, allow_network: bool = False) -> dict:
    """Excel's Goal Seek: adjust changing_cell until the formula in
    target_cell reaches target_value. Convergence-checked: a non-converging
    seek refuses and nothing is saved. save:false reports the solution
    without persisting it. Auto-backup on save; runs in a private hidden
    Excel; refuses while the file is open in Excel.
    allow_network: Let Excel contact the internet for this workbook's
    network functions. Set it only after the user has agreed, for this
    workbook."""
    return _comtier.com_goal_seek(
        path, target_cell, target_value, changing_cell, sheet=sheet,
        save=save, timeout_seconds=timeout_seconds, backup=backup,
        allow_network=allow_network)


@_op
def com_set_sparkline(path: str, action: str = "create",
                      location: str | None = None, source: str | None = None,
                      type: str = "line", sheet: str | None = None,
                      timeout_seconds: float | None = None,
                      backup: bool = True, allow_network: bool = False) -> dict:
    """Create, clear, or list sparkline groups through Excel.
    action='create': location is the cell/range that shows them (e.g.
    'G2:G10'), source the data range (e.g. 'A2:F10'), type 'line',
    'column', or 'win_loss'. 'clear' removes groups in location; 'list'
    inventories sheet, or the workbook when no sheet is named. Auto-backup
    on mutations; private hidden Excel.
    allow_network: Let Excel contact the internet for this workbook's
    network functions. Set it only after the user has agreed, for this
    workbook."""
    return _comtier.com_set_sparkline(
        path, action=action, location=location, source=source, type=type,
        sheet=sheet, timeout_seconds=timeout_seconds, backup=backup,
        allow_network=allow_network)


@_op
def com_validate_opens_clean(path: str, password: str | None = None,
                             timeout_seconds: float | None = None,
                             allow_network: bool = False) -> dict:
    """The authoritative corruption smoke test: open the file in a private
    hidden Excel and report whether it is accepted WITHOUT a repair prompt
    (a repair demand surfaces as a refusal carrying Excel's own message).
    Read-only; also the verify_com option of file-tier saves. password
    opens encrypted files; without one it reports encrypted, and a wrong
    one refuses at once.
    allow_network: Let Excel contact the internet for this workbook's
    network functions. Set it only after the user has agreed, for this
    workbook."""
    return _comtier.com_validate_opens_clean(
        path, password=password, timeout_seconds=timeout_seconds,
        allow_network=allow_network)


@_op
def com_status() -> dict:
    """Honest COM layer status: whether Excel COM is available on this
    machine, worker and pooled-instance state, journaled Excel PIDs, the
    operation running right now with its elapsed time, queued and completed
    counts, timeouts, and contention waits. Reports busy as busy rather
    than always ready. Never spawns Excel; safe to call anytime, including
    while another COM operation runs."""
    return _comtier.com_status()


# ------------------------------------------------------- the update ops
# The per-product half of core/update_flow.py (the family module): two
# thin ops in xl_app, a tool that is already not read-only, so the
# client's own permission prompt for it is the human check. Both refuse on
# Business, and the surface plan leaves them off the Business tool list.


def _update_op(fn, *args):
    from .core import update_flow
    from .core.errors import SurfaceRefusal

    try:
        return fn(*args)
    except update_flow.UpdateRefused as exc:
        raise SurfaceRefusal(exc.code, exc.message, hint=exc.hint) from None


@_op
def update_now() -> dict:
    """Install the newer version the update notice named; call only when the user asks or
    agrees. Takes no arguments and acts only on the newer stable version
    the server's own check found. Where the app can show a dialog, the
    user confirms there first; a Claude Desktop install opens Desktop's own
    install window. The install runs in the background: the answer comes
    within about 30 seconds, or says it is still running and the next
    answer reports the outcome, with what to do to start using it.
    """
    from .core import update_flow

    return _update_op(update_flow.update_now)


@_op
def update_notices(action: Literal["on", "off", "status"]) -> dict:
    """Turn update notices on or off, or show their status (action). The
    setting is kept in the server's state folder, so it survives updates
    and applies to every app that starts this server. off stops the
    notices and keeps the weekly check; status also reports the last
    notice and the newer version, if one is known.
    """
    from .core import update_flow

    return _update_op(update_flow.update_notices, action)


# ------------------------------------------------------------ the server


def instructions_for(policy) -> str:
    """Product instructions with the licensing components around them: the
    adapter's introduction first, its licensing instructions last (the order
    the released core used)."""
    from .surface import copy as _copy

    text = policy.initialization_text()
    return text["introduction"] + _copy.INSTRUCTIONS + " " + text["instructions"]


def create_server(settings=None):
    """Build the server. Raises PolicySetupError or SettingsError; never
    returns a half-configured server. The tool list is computed once here
    and never changes during the session."""
    from fastmcp import FastMCP

    from .settings import Settings
    from .surface import plan as _surface_plan
    from .surface.tools import make_tool

    policy = _license_policy.active()
    settings = Settings.from_env() if settings is None else settings
    plan = _surface_plan.build(settings)  # validates KS4XL_FORBID names
    _runtime.set_plan(plan)
    # version=: without it the MCP handshake reports fastmcp's own version
    # as serverInfo.version.
    mcp = FastMCP(name=policy.product_id, instructions=instructions_for(policy),
                  version=policy.distribution_version)
    for tool_plan in plan.tools.values():
        mcp.add_tool(make_tool(tool_plan, plan))
    from . import envelope as _envelope

    mcp.add_middleware(_envelope.unknown_tool_middleware(plan))
    # The policy's optional registration hook, offered after every core op
    # is registered; both shipped policies decline it. A tool it did add
    # would be registered here, at start, so the list stays fixed.
    _license_policy.register_additional_tools(mcp)
    _fixed_list(mcp)
    return mcp


def _fixed_list(mcp) -> None:
    """Advertise no list-changed notifications: the tool list is fixed for
    the session. fastmcp 3.4.7 hard-codes tools_changed=True on its
    stdio path; this pins every notification option off
    (tests/unit/test_family_stdio.py checks the capability on the wire)."""
    from mcp.server.lowlevel.server import NotificationOptions

    low = mcp._mcp_server  # noqa: SLF001 (pinned fastmcp; covered by the wire test)
    original = low.create_initialization_options

    def fixed(notification_options=None, **kwargs):
        return original(notification_options=NotificationOptions(), **kwargs)

    low.create_initialization_options = fixed


#: The most one log record may put on stderr.
LOG_LINE_CHARS = 300

#: How many log lines may wait for stderr before new ones are dropped.
LOG_QUEUE_LINES = 256


class _NonBlockingStderr:
    """A logging handler that never makes the server wait on stderr.

    Each record is one short line, but a client that pipes stderr and
    never reads it still fills the 4 KB Windows pipe buffer after about 36
    unknown requests before initialize, and the next write then blocked the
    server's only loop: no answer, not even to initialize (reproduced with
    80 requests). Lines now go through a bounded queue to a
    daemon writer thread. The server's loop only ever enqueues; when the
    queue is full the line is dropped and counted (`dropped`).

    The writer writes with os.write on its own duplicate of file descriptor
    2, never through sys.stderr: a
    thread stuck in sys.stderr.write holds that stream's buffer lock, and at
    interpreter shutdown Python waits for the lock, then writes a fatal
    error to the same full pipe and never exits. A thread stuck in os.write
    holds no Python lock, so the exit does not wait on it."""

    def __init__(self, stream=None, maxsize: int = LOG_QUEUE_LINES):
        import logging
        import queue
        import threading

        class _Handler(logging.Handler):
            def emit(inner, record: logging.LogRecord) -> None:
                try:
                    line = inner.format(record) + "\n"
                except Exception:  # noqa: BLE001 (logging must never raise)
                    return
                try:
                    self.queue.put_nowait(line)
                except queue.Full:
                    self.dropped += 1

            def handleError(inner, record) -> None:  # noqa: N802 (logging API)
                # The default prints a traceback to stderr: the very write
                # this handler exists to avoid.
                self.dropped += 1

        self.queue: "queue.Queue[str]" = queue.Queue(maxsize)
        self.dropped = 0
        self.stream = stream
        self.fd: int | None = None
        if stream is None:
            try:
                self.fd = _os.dup(_sys.stderr.fileno())  # not inheritable
            except Exception:  # noqa: BLE001 (no real stderr: write through the object)
                self.fd = None
        self.handler = _Handler()
        self.thread = threading.Thread(target=self._drain, daemon=True,
                                       name="ks4xl-stderr-writer")
        self.thread.start()

    def _drain(self) -> None:
        while True:
            line = self.queue.get()
            try:
                if self.fd is not None:
                    data = line.encode("utf-8", "replace")
                    while data:
                        data = data[_os.write(self.fd, data):]
                else:
                    stream = self.stream if self.stream is not None else _sys.stderr
                    stream.write(line)
                    stream.flush()
            except Exception:  # noqa: BLE001 (a closed stderr is not the server's fault)
                pass
            finally:
                self.queue.task_done()

    def flush(self, bound: float = 1.0) -> None:
        """Wait up to `bound` seconds for queued lines to reach stderr: the
        queue empty AND the line in hand written. An empty queue alone let
        the process exit while the writer still held the last line, so a
        client that reads stderr lost it. A stderr nobody reads costs this
        one bounded wait."""
        import time as _time

        end = _time.monotonic() + bound
        # unfinished_tasks counts every line put and not yet written
        # (task_done runs after the write), so there is no gap between the
        # writer taking a line and finishing it.
        while self.queue.unfinished_tasks and _time.monotonic() < end:
            _time.sleep(0.01)


#: The installed writer, for the exit flush and the tests.
_STDERR_WRITER: _NonBlockingStderr | None = None


def _hooks_to_logging() -> None:
    """Python's own last-resort writers
    (an exception nobody could raise, an exception that ended a thread, a
    warning) wrote to sys.stderr directly, so on a full pipe they could stop
    the loop exactly as an unread stderr did. They now go through logging,
    one line each, to the same non-blocking writer."""
    import logging
    import threading

    log = logging.getLogger("xlsx_mcp")

    def unraisable(u) -> None:
        log.warning("unraisable %s: %s (%s)", type(u.exc_value).__name__,
                    u.exc_value, u.err_msg or "in a finalizer")

    def thread_ended(args) -> None:
        if args.exc_type is SystemExit:
            return
        name = args.thread.name if args.thread is not None else "a thread"
        log.warning("%s ended on %s: %s", name, args.exc_type.__name__, args.exc_value)

    _sys.unraisablehook = unraisable
    threading.excepthook = thread_ended
    logging.captureWarnings(True)


def _bound_stderr_logging() -> _NonBlockingStderr:
    """One short line on stderr per log record, never a traceback, and
    never a wait on stderr (_NonBlockingStderr).

    The mcp library logs a request that fails validation through the root
    logger (logging.warning), which configures a default handler on first
    use. For one unknown method before initialize that was a 6,379-byte
    pydantic dump. A client that pipes stderr and never reads it (a 4 KB
    pipe buffer on Windows) blocks the server on that write, and the server
    then answers nothing: the server "goes silent". Every record reaching the
    root logger, and fastmcp's own logger, is now one line, clipped, with no
    exception text, handed to a writer thread that may fall behind or stall
    without stopping the server."""
    import logging

    global _STDERR_WRITER

    class _OneLine(logging.Formatter):
        def format(self, record: logging.LogRecord) -> str:
            try:
                text = record.getMessage()
            except Exception:  # noqa: BLE001 (a bad format string is still one line)
                text = str(record.msg)
            lines = text.strip().splitlines() or [""]
            line = f"{record.levelname}:{record.name}:{lines[0]}"
            if len(lines) > 1:
                line += f" (+{len(lines) - 1} lines not shown)"
            if len(line) > LOG_LINE_CHARS:
                line = line[:LOG_LINE_CHARS - 3] + "..."
            return line

    if _STDERR_WRITER is None:
        _STDERR_WRITER = _NonBlockingStderr()
        # Whatever path the process leaves by (a crash, a SystemExit
        # outside main's own finally), queued lines get the same bounded
        # chance to reach a client that reads stderr.
        import atexit

        atexit.register(_STDERR_WRITER.flush, 1.0)
    writer = _STDERR_WRITER
    _hooks_to_logging()
    writer.handler.setFormatter(_OneLine())
    # The root logger (the mcp library's request warnings) and fastmcp's own
    # logger (its rich handlers write to stderr directly, and it does not
    # propagate) both go through the one non-blocking writer, and so do
    # captured warnings (py.warnings).
    for logger in (logging.getLogger(), logging.getLogger("fastmcp")):
        for handler in list(logger.handlers):
            if handler is not writer.handler:
                logger.removeHandler(handler)
        if writer.handler not in logger.handlers:
            logger.addHandler(writer.handler)
    return writer


def _recover_at_start(settings) -> list[str]:
    """Put back every workbook that a save which did not finish left aside
    (a server stopped in the middle of a save). A read-only
    server changes nothing on disk, so it leaves that to the next server
    that may write. Never fails the start."""
    if getattr(settings, "read_only", False):
        return []
    try:
        from .core import safesave as _safesave_mod
        return _safesave_mod.recover_interrupted_saves()
    except Exception:  # noqa: BLE001
        return []


def _legacy_notice(settings) -> None:
    """Retired settings set to an old default are accepted and ignored, with
    ONE line on stderr naming them (never their values), written only after
    every setting has been validated and the surface built."""
    names = tuple(getattr(settings, "legacy_set", ()))
    if not names:
        return
    from .surface import copy as _copy

    _sys.stderr.write("[kitchensink4xl] "
                      + _copy.LEGACY_SETTING_NOTICE.replace("{names}", ", ".join(names)) + "\n")
    _sys.stderr.flush()


def main() -> None:
    from .settings import Settings, SettingsError

    try:
        settings = Settings.from_env()
        mcp = create_server(settings)
    except _license_policy.PolicySetupError as exc:
        _sys.stderr.write(str(exc) + "\n")
        _sys.stderr.flush()
        raise SystemExit(2) from None
    except SettingsError as exc:
        # One readable line, never a traceback: Desktop shows "server failed
        # to start" and buries anything longer. Flushed here, before the exit:
        # the line is the whole answer for a client that reads stderr.
        _sys.stderr.write(f"[kitchensink4xl] {exc} The server did not start.\n")
        _sys.stderr.flush()
        raise SystemExit(2) from None
    _legacy_notice(settings)
    # Put back any workbook a save which did not finish left
    # aside (a server stopped mid-save). A read-only server changes nothing.
    _recover_at_start(settings)
    # No update check here: startup starts no thread and asks PyPI nothing.
    # The weekly check starts in the background after a tool call
    # (core/update_check.maybe_refresh_in_background).
    # The one-time star nudge, last, after startup has already succeeded.
    # It is a single line on stderr on the very first run of an install and
    # nothing at all thereafter; it asks the network nothing.
    _star_nudge.announce_once()
    _bound_stderr_logging()
    # The framework banner has its own PyPI checker, independent of our off
    # switch; keep startup offline in both editions.
    try:
        mcp.run(transport="stdio", show_banner=False)
    finally:
        # The client went away. Calls that never committed are
        # cancelled, and every call still running gets a bounded wait to
        # finish and release its hidden Excel; nothing is ended here.
        from .surface import dispatch as _dispatch

        _dispatch.drain_inflight()
        _stop_update_child()
        if _STDERR_WRITER is not None:
            # Give queued log lines a moment to reach a stderr that is being
            # read; a stderr nobody reads costs this one bounded second.
            _STDERR_WRITER.flush(1.0)


def _stop_update_child() -> None:
    """An update child still running when the client goes away is killed."""
    flow = _sys.modules.get("xlsx_mcp.core.update_flow")
    if flow is not None:
        flow.shutdown()


if __name__ == "__main__":
    main()
