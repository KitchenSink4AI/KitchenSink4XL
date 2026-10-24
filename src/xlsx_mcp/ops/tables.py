"""ops/tables.py: Excel tables (ListObjects).

create_table turns a range into a ListObject with a header row, a table style,
and an optional totals row. get_table reads a table back as records or a matrix
with column addressing and the honest cached-value story. manage_table is the
advanced lifecycle (add/delete rows and columns, rename, resize, toggle totals,
convert to a plain range, restyle).

Structural row/column edits route through WorkbookPackage.modify_structure, so
core.refs rewrites every formula, name, conditional-format, data-validation, and
merge that pointed across the change, and the table ref is then set explicitly
to its intended bounds (the boundary-insert case openpyxl's transposer cannot
disambiguate). Reads load the workbook directly and never mutate.
"""

from __future__ import annotations

import re
from typing import Any

from openpyxl.utils import (
    get_column_letter,
    range_boundaries,
)

from ..core import calc as _calc
from ..core import refs as _refs
from ..core.errors import TargetNotFound, XlMcpError
from ..core.package import WorkbookPackage
from . import gridio


class _ProbeGrid:
    """The three fields gridio.formula_mask reads off a resolved grid, for
    callers (get_table) that address a rectangle without resolving one."""

    def __init__(self, sheet, min_row, min_col, max_row, max_col):
        self.sheet = sheet
        self.min_row = min_row
        self.min_col = min_col
        self.max_row = max_row
        self.max_col = max_col

#: Excel table (ListObject) name grammar: a letter or underscore, then letters,
#: digits, periods, underscores; never a cell-reference shape like "A1".
_NAME_RE = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.]*$")
_CELLREF_RE = re.compile(r"^[A-Za-z]{1,3}\$?\d+$")
#: R1C1-shaped names (R, C, R1C1, R2, C33) are also illegal in Excel; writing
#: one corrupts formula parsing when the workbook opens in R1C1 mode.
_R1C1_RE = re.compile(r"^[Rr](\d+)?([Cc](\d+)?)?$|^[Cc](\d+)?$"
                     r"|^(?:[Rr]?[Cc]|[Rr])[0-9]")  # also a name that BEGINS with one (R1_x, C21_Name)

TABLE_ACTIONS = (
    "add_row", "delete_row", "add_column", "delete_column", "rename",
    "resize", "toggle_totals", "to_range", "set_style",
)

_TOTALS_FUNCS = frozenset({
    "sum", "average", "count", "countNums", "min", "max", "stdDev", "var",
    "none", "custom",
})


#: The SUBTOTAL function number Excel writes in a totals row for each
#: totalsRowFunction (measured: sum -> SUBTOTAL(109,T[Qty]), average -> 101,
#: count -> 103; fix4-scratch/oracle_totals.py). The 1xx forms ignore rows a
#: filter hides, as Excel's own totals row does.
_SUBTOTAL_CODES = {"average": 101, "countNums": 102, "count": 103, "max": 104,
                   "min": 105, "stdDev": 107, "sum": 109, "var": 110}


def _refuse_occupied_totals_row(ws, table_name: str, row: int, min_col: int,
                                max_col: int) -> None:
    """The totals row goes in the row under the table.
    Excel shifts cells there down; this server refuses instead of
    overwriting them (the old behaviour silently made that data the totals
    row)."""
    taken = [gridio.a1(row, c) for c in range(min_col, max_col + 1)
             if ws._cells.get((row, c)) is not None
             and ws._cells[(row, c)].value is not None]
    if taken:
        exc = XlMcpError(
            f"the row under table {table_name} holds data ({', '.join(taken[:5])}), "
            "which a totals row would overwrite. Insert a row below the table first "
            f"(xl_structure modify_grid_structure action='insert_rows' at={row}), then "
            "turn the totals row on.")
        exc.code = "CONFLICT"
        raise exc


def _write_totals_row(pkg, ws, table, row: int, min_col: int) -> list[str]:
    """Put into the totals row's cells what Excel puts
    there: the label as text, and for each column with a function the
    SUBTOTAL formula over the column (T[Qty]), stored through the formula
    path. Before, only the table XML carried totalsRowFunction and the
    cells stayed empty, so Excel showed an empty totals row."""
    from ..core import structrefs as _sr

    written = []
    for i, tc in enumerate(table.tableColumns or []):
        cell = gridio.a1(row, min_col + i)
        fn = tc.totalsRowFunction
        if fn in _SUBTOTAL_CODES:
            pkg.set_cell(ws.title, cell, f"=SUBTOTAL({_SUBTOTAL_CODES[fn]},"
                                         f"{table.displayName}[{_sr.escape(tc.name)}])")
            written.append(cell)
        elif tc.totalsRowLabel:
            pkg.set_cell(ws.title, cell, tc.totalsRowLabel)
            written.append(cell)
    return written


def _validate_name(wb, name: str) -> str:
    if not isinstance(name, str) or not name.strip():
        raise XlMcpError("a table needs a non-empty name")
    name = name.strip()
    if _CELLREF_RE.match(name) or _R1C1_RE.match(name) \
            or not _NAME_RE.match(name):
        raise XlMcpError(
            f"table name {name!r} is not a legal Excel table name (start with "
            "a letter or underscore, no spaces, not an A1 or R1C1 cell "
            "reference)")
    if len(name) > 255:
        raise XlMcpError("table name exceeds 255 characters")
    existing = {t.lower() for w in wb.worksheets for t in getattr(w, "tables", {})}
    if name.lower() in existing:
        raise XlMcpError(f"a table named {name!r} already exists in the workbook")
    # Table names and defined names share one namespace in Excel.
    defined = {n.lower() for n in wb.defined_names}
    for w in wb.worksheets:
        dn = getattr(w, "defined_names", None)
        if dn is not None:
            defined.update(n.lower() for n in dn)
    if name.lower() in defined:
        raise XlMcpError(
            f"{name!r} is already a defined name; table and defined names "
            "share one namespace in Excel")
    return name


def _check_no_table_overlap(ws, min_row, min_col, max_row, max_col) -> None:
    """Excel refuses overlapping ListObjects; writing one anyway produces a
    repair-prompt file, so the overlap refuses here instead."""
    tables = getattr(ws, "tables", {})
    for tname in list(tables):
        t = tables[tname]  # TableList.items() yields ref strings, not Tables
        t_min_c, t_min_r, t_max_c, t_max_r = range_boundaries(t.ref)
        if (min_row <= t_max_r and max_row >= t_min_r
                and min_col <= t_max_c and max_col >= t_min_c):
            raise XlMcpError(
                f"the range overlaps table {tname!r} ({t.ref}) on "
                f"{ws.title!r}; Excel does not allow overlapping tables")


def _find_table(wb, name: str):
    for ws in wb.worksheets:
        tables = getattr(ws, "tables", {})
        if name in tables:
            return ws, tables[name]
        for tname in list(tables):
            # index by name: TableList.items() yields ref STRINGS, not Table
            # objects, so the old items() loop returned a str here
            if tname.lower() == name.lower():
                return ws, tables[tname]
    raise TargetNotFound(
        f"no table named {name!r}; tables: "
        + (", ".join(repr(t) for w in wb.worksheets
                     for t in getattr(w, "tables", {})) or "none"))


def _bounds(ref: str) -> tuple[int, int, int, int]:
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    return min_row, min_col, max_row, max_col


def _ref(min_row: int, min_col: int, max_row: int, max_col: int) -> str:
    return (f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{max_row}")


def _header_names(ws, min_row: int, min_col: int, max_col: int) -> list[str]:
    return [ws.cell(min_row, c).value
            for c in range(min_col, max_col + 1)]


def _sync_columns(ws, table, min_row: int, min_col: int, max_col: int) -> None:
    """Rebuild the table's column list from the current header row, but only if
    the columns were materialized; an empty column list lets openpyxl re-derive
    the columns from the header cells on save. Existing totals functions and
    labels are preserved by column name."""
    if not table.tableColumns:
        return
    from openpyxl.worksheet.table import TableColumn
    prior = {tc.name: tc for tc in table.tableColumns}
    names = _header_names(ws, min_row, min_col, max_col)
    cols = []
    for i, hn in enumerate(names):
        cn = str(hn) if hn is not None else f"Column{i + 1}"
        tc = TableColumn(id=i + 1, name=cn)
        old = prior.get(cn)
        if old is not None:
            tc.totalsRowFunction = old.totalsRowFunction
            tc.totalsRowLabel = old.totalsRowLabel
        cols.append(tc)
    table.tableColumns = cols


# --------------------------------------------------------------- create_table


def create_table(path: str, location: Any, name: str, header: bool = True,
                 style: str = "TableStyleMedium9", row_stripes: bool = True,
                 col_stripes: bool = False, totals_row: bool = False,
                 totals: dict | None = None, sheet: str | None = None,
                 allow_loss: bool = False, backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Turn a range into a table (ListObject). One backup + one verified save."""
    from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

    pkg = WorkbookPackage.open(path)
    wb = pkg.workbook
    name = _validate_name(wb, name)
    grid = pkg.resolve(location, default_sheet=sheet)
    if grid.empty:
        raise XlMcpError("cannot build a table on an empty range")
    ws = wb[grid.sheet]
    min_row, min_col, max_row, max_col = (
        grid.min_row, grid.min_col, grid.max_row, grid.max_col)
    _check_no_table_overlap(ws, min_row, min_col, max_row, max_col)
    ncols = max_col - min_col + 1

    if header:
        raw = _header_names(ws, min_row, min_col, max_col)
        seen: set[str] = set()
        col_names: list[str] = []
        for i, h in enumerate(raw):
            hn = str(h).strip() if h is not None and str(h).strip() else \
                f"Column{i + 1}"
            base, k = hn, 1
            while hn.lower() in seen:
                k += 1
                hn = f"{base}{k}"
            seen.add(hn.lower())
            col_names.append(hn)
        header_row_count = 1
    else:
        col_names = [f"Column{i + 1}" for i in range(ncols)]
        header_row_count = 0

    if totals_row and max_row - min_row + 1 - header_row_count < 1:
        raise XlMcpError("a totals row needs at least one data row above it")

    table = Table(displayName=name, ref=_ref(min_row, min_col, max_row, max_col))
    table.headerRowCount = header_row_count
    need_columns = (not header) or totals_row or bool(totals)
    if need_columns:
        cols = []
        for i, cn in enumerate(col_names):
            tc = TableColumn(id=i + 1, name=cn)
            cols.append(tc)
        table.tableColumns = cols
    if totals_row:
        _refuse_occupied_totals_row(ws, name, max_row + 1, min_col, max_col)
        table.totalsRowCount = 1
        table.ref = _ref(min_row, min_col, max_row + 1, max_col)
        funcs = totals or {}
        for i, tc in enumerate(table.tableColumns):
            fn = funcs.get(tc.name)
            if fn is not None:
                if fn not in _TOTALS_FUNCS:
                    raise XlMcpError(
                        f"totals function {fn!r} must be one of "
                        f"{sorted(_TOTALS_FUNCS)}")
                tc.totalsRowFunction = fn
            elif i == 0:
                tc.totalsRowLabel = "Total"
    table.tableStyleInfo = TableStyleInfo(
        name=style, showRowStripes=row_stripes, showColumnStripes=col_stripes,
        showFirstColumn=False, showLastColumn=False)
    ws.add_table(table)
    totals_cells = (_write_totals_row(pkg, ws, table, max_row + 1, min_col)
                    if totals_row else [])

    pkg._changed["table"] = {
        "created": name, "sheet": ws.title, "ref": table.ref,
        "columns": col_names, "header": header, "totals_row": totals_row}
    if totals_cells:
        pkg._changed["table"]["totals_cells"] = totals_cells
    return pkg.save(allow_loss=allow_loss, backup=backup,
                    verify_com=verify_com)


# ------------------------------------------------------------------ get_table


def get_table(path: str, name: str, columns: list | None = None,
              values: str = "cached", records: bool = False) -> dict:
    """Read a table's data by name. columns projects a subset; values is
    cached | formula | both | displayed ('displayed' is the text each cell
    shows in Excel), and cached and both add `shows` for the cells whose
    shown text differs from their value. Read-only; nothing is written."""
    if values not in gridio.VALUE_MODES:
        raise XlMcpError(f"values must be one of {gridio.VALUE_MODES}")
    formula_wb = gridio.open_wb(path, data_only=False) \
        if values in ("formula", "both") else None
    cached_wb = gridio.open_wb(path, data_only=True) \
        if values in gridio.CACHED_MODES else None
    base = formula_wb if formula_wb is not None else cached_wb
    try:
        ws, table = _find_table(base, name)
        min_row, min_col, max_row, max_col = _bounds(table.ref)
        hrc = table.headerRowCount if table.headerRowCount is not None else 1
        trc = table.totalsRowCount if table.totalsRowCount is not None else 0
        all_names = [str(base[ws.title].cell(min_row, c).value)
                     if hrc else f"Column{c - min_col + 1}"
                     for c in range(min_col, max_col + 1)]
        if columns:
            sel = []
            for c in columns:
                if str(c) not in all_names:
                    raise TargetNotFound(
                        f"table {name!r} has no column {c!r}; columns: "
                        f"{all_names}")
                sel.append(all_names.index(str(c)))
        else:
            sel = list(range(len(all_names)))
        proj_names = [all_names[i] for i in sel]
        data_top = min_row + hrc
        data_bottom = max_row - trc
        rws = base[ws.title]
        cws = cached_wb[ws.title] if cached_wb is not None else None
        fws = formula_wb[ws.title] if formula_wb is not None else None
        # A table body is where a totals or calculated-column formula lives,
        # so 'cached' handing back a null for an uncalculated one is exactly
        # the silent blank the numbers-safety rule forbids. The mask names
        # them; the caller gets the addresses and the remedy.
        mask = None
        probe_failed = False
        if values in ("cached", "displayed"):
            probe_grid = _ProbeGrid(ws.title, data_top, min_col,
                                    data_bottom, max_col)
            raw_mask = gridio.formula_mask(path, probe_grid)
            probe_failed = raw_mask is None
            mask = raw_mask or {}
        absent_cells: list[str] = []
        shown = gridio.DisplayNotes(gridio.is_1904(ws), path) \
            if values == "displayed" else None
        differs = gridio.ShownText(path, gridio.is_1904(ws)) \
            if values in ("cached", "both") else None
        # A formula result read from the cache, which the
        # stale-results note covers (read_range and the other readers
        # already said so; get_table did not).
        cached_formula = False
        text_cells: list[str] = []  # formula mode: text that starts with "="
        out_rows: list[Any] = []
        for r in range(data_top, data_bottom + 1):
            row = []
            for i in sel:
                c = min_col + i
                if values == "formula":
                    v = fws.cell(r, c).value
                    if _calc.looks_like_formula_text(fws.cell(r, c)):
                        text_cells.append(gridio.a1(r, c))
                elif values in ("cached", "displayed"):
                    v = cws.cell(r, c).value
                    if mask and (r, c) in mask:
                        if v is None:
                            absent_cells.append(gridio.a1(r, c))
                        else:
                            cached_formula = True
                    if values == "displayed":
                        v = shown.text(v, cws.cell(r, c).number_format,
                                       gridio.a1(r, c), ws.title, r, c)
                else:
                    fcell = fws.cell(r, c)
                    v = fcell.value if _calc.is_formula_cell(fcell) \
                        else cws.cell(r, c).value
                    if _calc.is_formula_cell(fcell):
                        if cws.cell(r, c).value is None:
                            absent_cells.append(gridio.a1(r, c))
                        else:
                            cached_formula = True
                if differs is not None and v is not None and not (
                        values == "both" and _calc.is_formula_cell(fws.cell(r, c))):
                    differs.add(ws.title, r, c, v, cws.cell(r, c).number_format)
                row.append(gridio.compact_value(v))
            out_rows.append(dict(zip(proj_names, row)) if records else row)
        if shown is not None:
            for (_s, r, c), text in shown.settle().items():
                if c - min_col in sel:
                    j = sel.index(c - min_col)
                    out_rows[r - data_top][proj_names[j] if records else j] = text
        result = {
            "sheet": ws.title, "table": table.displayName or table.name,
            "ref": table.ref, "columns": proj_names,
            "rows": out_rows, "row_count": len(out_rows),
            "has_totals_row": trc > 0, "value_mode": values,
        }
        if differs is not None:
            result.update(differs.fields())
        if absent_cells:
            result["absent_cells"] = absent_cells[:100]
            result["warning"] = gridio.ABSENT_WARNING
        if shown is not None and shown.note():
            result["display_note"] = shown.note()
        result.update(gridio.text_not_formulas(text_cells))
        if values != "formula" and (cached_formula or probe_failed):
            pending = gridio.pending_recalc_note(path)
            if pending:
                result["cached_note"] = pending
        return result
    finally:
        for wb in (formula_wb, cached_wb):
            if wb is not None:
                wb.close()


# --------------------------------------------------------------- manage_table


def manage_table(path: str, name: str, action: str, values: list | None = None,
                 index: int | None = None, column: str | None = None,
                 new_name: str | None = None, new_ref: str | None = None,
                 style: str | None = None, on: bool = True,
                 totals: dict | None = None, row_stripes: bool | None = None,
                 col_stripes: bool | None = None, sheet: str | None = None,
                 allow_loss: bool = False, backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Advanced table lifecycle. One backup + one verified save per call."""
    if action not in TABLE_ACTIONS:
        raise XlMcpError(f"action must be one of {TABLE_ACTIONS}, got {action!r}")
    pkg = WorkbookPackage.open(path)
    wb = pkg.workbook
    ws, table = _find_table(wb, name)
    min_row, min_col, max_row, max_col = _bounds(table.ref)
    hrc = table.headerRowCount if table.headerRowCount is not None else 1
    trc = table.totalsRowCount if table.totalsRowCount is not None else 0
    detail: dict[str, Any] = {"action": action, "table": table.displayName}

    if action == "add_row":
        rows = _as_rows(values)
        n = len(rows)
        insert_at = max_row - trc + 1
        # What covers exactly the table's data (a total, a name, a rule's
        # range, a chart series) grows with the table, as in Excel.
        data_first, data_last = min_row + hrc, max_row - trc
        grow = (((min_row, data_first, data_last, max_row, min_col, max_col),)
                if data_last >= data_first else ())
        pkg.modify_structure(_refs.RefEdit(
            ws.title, _refs.INSERT_ROWS, index=insert_at, count=n, grow=grow))
        ws = pkg.workbook[ws.title]
        table = ws.tables[table.name]
        table.ref = _ref(min_row, min_col, max_row + n, max_col)
        # A calculated column the caller gives no value for takes its column
        # formula, as Excel's own new table row does (ListRows.Add).
        calc = _calculated_formulas(table)
        data_top = min_row + hrc
        filled = []
        for i, row in enumerate(rows):
            r = insert_at + i
            width = max_col - min_col + 1
            for j in range(width):
                val = row[j] if j < len(row) else None
                coord = gridio.a1(r, min_col + j)
                if val is None and j in calc:
                    pkg.set_formula(ws.title, coord, _refs.offset_formula(
                        "=" + calc[j], r - data_top, 0, ws.title))
                    filled.append(coord)
                elif j < len(row):
                    pkg.set_cell(ws.title, coord, val, from_caller=True)
        detail.update(added_rows=n, ref=table.ref)
        if filled:
            detail["calculated_cells"] = filled

    elif action == "delete_row":
        if index is None or int(index) < 1:
            raise XlMcpError("delete_row needs index (1-based data row)")
        data_rows = max_row - trc - (min_row + hrc) + 1
        if int(index) > data_rows:
            raise XlMcpError(
                f"row index {index} out of range; the table has {data_rows} "
                "data row(s)")
        sheet_row = min_row + hrc + int(index) - 1
        pkg.modify_structure(_refs.RefEdit(
            ws.title, _refs.DELETE_ROWS, index=sheet_row, count=1))
        ws = pkg.workbook[ws.title]
        table = ws.tables[table.name]
        table.ref = _ref(min_row, min_col, max_row - 1, max_col)
        detail.update(deleted_row=int(index), ref=table.ref)

    elif action == "add_column":
        if not column:
            raise XlMcpError("add_column needs column (the new header name)")
        new_col = max_col + 1
        if new_col > _refs.MAX_COL:
            raise XlMcpError("the table already reaches the last column")
        if hrc:
            pkg.set_cell(ws.title, gridio.a1(min_row, new_col), column)
        col_vals = _as_flat(values)
        data_top = min_row + hrc
        for i, val in enumerate(col_vals):
            pkg.set_cell(ws.title, gridio.a1(data_top + i, new_col), val,
                         from_caller=True)
        table.ref = _ref(min_row, min_col, max_row, new_col)
        _sync_columns(ws, table, min_row, min_col, new_col)
        detail.update(added_column=column, ref=table.ref)

    elif action == "delete_column":
        if not column:
            raise XlMcpError("delete_column needs column (the header name)")
        names = _header_names(ws, min_row, min_col, max_col)
        idx = next((i for i, h in enumerate(names)
                    if h is not None and str(h).strip() == column), None)
        if idx is None:
            raise TargetNotFound(
                f"table {name!r} has no column {column!r}; columns: {names}")
        if max_col - min_col + 1 == 1:
            raise XlMcpError("cannot delete the only column of a table")
        sheet_col = min_col + idx
        pkg.modify_structure(_refs.RefEdit(
            ws.title, _refs.DELETE_COLS, index=sheet_col, count=1))
        ws = pkg.workbook[ws.title]
        table = ws.tables[table.name]
        table.ref = _ref(min_row, min_col, max_row, max_col - 1)
        _sync_columns(ws, table, min_row, min_col, max_col - 1)
        # What named the deleted column is now #REF!, as Excel stores it.
        dead = _refs.deleted_column_refs(pkg.workbook, table.displayName, str(names[idx]))
        if dead.formulas or dead.rule_formulas:
            pkg._formula_written = True
        detail.update(deleted_column=column, ref=table.ref)
        if dead.formulas or dead.names or dead.rule_formulas:
            n_dead = dead.formulas + dead.names + dead.rule_formulas
            detail.update(references_now_ref_error=n_dead)
            # said in words as well as counted (a reader that skims the counts can miss them), by kind, since
            # only cell formulas are listed by audit_formulas; a formula with a fallback such as IFERROR
            # returns that fallback instead of an error, so the wording says "may"
            kinds = [(dead.formulas, "cell formula", "cell formulas"),
                     (dead.names, "defined name", "defined names"),
                     (dead.rule_formulas, "formatting or validation rule", "formatting or validation rules")]
            parts = [f"{n} {one if n == 1 else many}" for n, one, many in kinds if n]
            which = parts[0] if len(parts) == 1 else ", ".join(parts[:-1]) + " and " + parts[-1]
            tail = " xl_read audit_formulas lists the cell formulas." if dead.formulas else ""
            pkg._notes.append(
                f"{which} that used the deleted column '{column}' of table {table.displayName} now "
                f"{'holds' if n_dead == 1 else 'hold'} #REF! in place of that column, so what depends on "
                f"{'it' if n_dead == 1 else 'them'} may show an error or a fallback value.{tail} Tell the user "
                "what changed before they rely on the workbook")

    elif action == "rename":
        if not new_name:
            raise XlMcpError("rename needs new_name")
        new = _validate_name(wb, new_name)
        old = table.name
        del ws.tables[old]
        table.displayName = new
        table.name = new
        ws.add_table(table)
        # Every structured reference to the table follows
        # the new name, as in Excel (=SUM(Stock_Tbl[Value]) on any sheet, and
        # defined names); text inside a string is left alone, as Excel does.
        rewrites = _refs.rename_table_refs(wb, old, new)
        # formula text changes, results do not: they stay stored
        pkg.note_text_rewrite(
            lambda t, o=old, n=new: _refs.rename_table_in_formula(t, o, n))
        detail.update(renamed_from=old, renamed_to=new,
                      references_rewritten=rewrites.as_dict())

    elif action == "resize":
        if not new_ref:
            raise XlMcpError("resize needs new_ref (an A1 range)")
        nb = _bounds(new_ref.upper())
        if nb[0] != min_row or nb[1] != min_col:
            raise XlMcpError(
                "resize must keep the table's top-left header anchor "
                f"({get_column_letter(min_col)}{min_row}); move it with a "
                "separate edit if needed")
        table.ref = new_ref.upper()
        detail.update(ref=table.ref)

    elif action == "toggle_totals":
        if on and trc == 0:
            if max_row + 1 > _refs.MAX_ROW:
                raise XlMcpError("no room below the table for a totals row")
            _refuse_occupied_totals_row(ws, table.displayName, max_row + 1, min_col,
                                        max_col)
            table.totalsRowCount = 1
            table.ref = _ref(min_row, min_col, max_row + 1, max_col)
            names = _header_names(ws, min_row, min_col, max_col)
            from openpyxl.worksheet.table import TableColumn
            cols = []
            funcs = totals or {}
            for i, hn in enumerate(names):
                cn = str(hn) if hn is not None else f"Column{i + 1}"
                tc = TableColumn(id=i + 1, name=cn)
                fn = funcs.get(cn)
                if fn is not None:
                    if fn not in _TOTALS_FUNCS:
                        raise XlMcpError(
                            f"totals function {fn!r} must be one of "
                            f"{sorted(_TOTALS_FUNCS)}")
                    tc.totalsRowFunction = fn
                elif i == 0:
                    tc.totalsRowLabel = "Total"
                cols.append(tc)
            table.tableColumns = cols
            detail.update(totals_row=True, ref=table.ref,
                          totals_cells=_write_totals_row(pkg, ws, table, max_row + 1,
                                                         min_col))
        elif not on and trc == 1:
            # The totals row's label and SUBTOTAL formulas go with it.
            for c in range(min_col, max_col + 1):
                if ws._cells.get((max_row, c)) is not None:
                    ws.cell(max_row, c).value = None
            pkg.note_region_write(ws.title, max_row, min_col, max_row, max_col)
            table.totalsRowCount = 0
            table.ref = _ref(min_row, min_col, max_row - 1, max_col)
            for tc in table.tableColumns or []:
                tc.totalsRowFunction = None
                tc.totalsRowLabel = None
            detail.update(totals_row=False, ref=table.ref)
        else:
            detail.update(totals_row=bool(trc), note="already in that state")

    elif action == "to_range":
        # Every formula that names the table is rewritten as plain cell
        # references first, as Excel's Convert to Range does (or the whole
        # call is refused, nothing changed, naming what cannot be).
        from ..core import tablerange as _tablerange
        rewritten = _tablerange.convert(wb, ws, table)
        if rewritten["formulas"] or rewritten["names"]:
            pkg._formula_written = True
            detail.update(references_rewritten=rewritten)
        del ws.tables[table.name]
        # Deliberate removal of the table part; declare it so the default-fail
        # inventory verify does not read the vanished xl/tables/ part as
        # silent loss. The cell data stays on the sheet.
        pkg.expect_removal("xl/tables/")
        detail.update(converted_to_range=True)
        # said in words as well as flagged: a weaker model converted a table on its own after a refusal and
        # answered only "Done"
        n_f, n_n = rewritten["formulas"], rewritten["names"]
        uses = [f"{n} {'formula' if n == 1 else 'formulas'}" for n in (n_f,) if n] + \
               [f"{n} defined {'name' if n == 1 else 'names'}" for n in (n_n,) if n]
        moved = (f", and {' and '.join(uses)} that used it now {'uses' if n_f + n_n == 1 else 'use'} "
                 "cell references") if uses else ""
        pkg._notes.append(
            f"table {table.displayName} is now a plain range ({ws.title}!{table.ref}): the cells keep their "
            f"values, the table and its name are gone{moved}. Tell the user that the table was removed, "
            "unless they asked for it")

    else:  # set_style
        from openpyxl.worksheet.table import TableStyleInfo
        cur = table.tableStyleInfo
        table.tableStyleInfo = TableStyleInfo(
            name=style if style is not None else (
                cur.name if cur else "TableStyleMedium9"),
            showRowStripes=row_stripes if row_stripes is not None else (
                cur.showRowStripes if cur else True),
            showColumnStripes=col_stripes if col_stripes is not None else (
                cur.showColumnStripes if cur else False),
            showFirstColumn=cur.showFirstColumn if cur else False,
            showLastColumn=cur.showLastColumn if cur else False)
        detail.update(style=table.tableStyleInfo.name)

    pkg._changed["table"] = detail
    return pkg.save(allow_loss=allow_loss, backup=backup,
                    verify_com=verify_com)


def _calculated_formulas(table) -> dict[int, str]:
    """column index in the table -> its calculated column formula (as
    stored, without '=')."""
    out = {}
    for i, tc in enumerate(table.tableColumns or ()):
        text = getattr(getattr(tc, "calculatedColumnFormula", None),
                       "attr_text", None)
        if isinstance(text, str) and text.strip():
            out[i] = text
    return out


def _as_rows(values) -> list[list]:
    if not values:
        raise XlMcpError("add_row needs values (a row list, or a list of rows)")
    if all(isinstance(v, list) for v in values):
        return values
    return [values]


def _as_flat(values) -> list:
    if not values:
        return []
    if all(isinstance(v, list) for v in values):
        return [v[0] if v else None for v in values]
    return list(values)


__all__ = ["create_table", "get_table", "manage_table", "TABLE_ACTIONS"]
