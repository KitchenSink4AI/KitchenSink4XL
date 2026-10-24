"""ops/gridio.py: shared read-side plumbing for the cell/range/view families.

Reads never mutate, so they do not go through WorkbookPackage; they load the
workbook directly (sandbox-gated, keep_vba for .xlsm and .xltm) and resolve the location
object through core.locate. The one subtlety is the honest calc story:
openpyxl exposes EITHER formula strings (data_only=False) OR the
last cached values (data_only=True), never both from one load. So a read that
wants both loads twice, and every returned value is labelled cached | calculated |
formula | absent | value so no read ever passes off an empty formula cell as a
blank cell.

All the write families still route their mutation through WorkbookPackage; this
module is read-only support plus the value-labelling helpers they share.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any

from openpyxl.utils import get_column_letter

from ..core import calc as _calc
from ..core import hazard as _hazard
from ..core import locate as _locate
from ..core import sharedformulas as _shared
from ..core import xmlsafe as _xmlsafe
from ..core.errors import (
    CeilingExceeded,
    WorkbookCorrupt,
    XlMcpError,
)
from ..core.sandbox import check_path

#: A single read that would materialize more than this many cells refuses and
#: asks the caller to page or slice, so a read can never blow the context window
#: or the process memory by accident (the token-discipline promise).
MAX_READ_CELLS = 200_000
#: Writes and server-side scans return a summary, not the
#: cells, so the read ceiling (which protects the caller's context) does not
#: bind them. These bound the server's own time and memory instead.
MAX_WRITE_CELLS = 1_000_000
MAX_SCAN_CELLS = 5_000_000

VALUE_MODES = ("cached", "formula", "both", "displayed")
#: The modes that read the cached-value load. "displayed" is the cached value
#: as the cell's number format shows it
#: (a ZIP formatted 00000 reads "02134", a date "14-Jan-26").
CACHED_MODES = ("cached", "both", "displayed")


def fill_empty_text_results(wb) -> int:
    """A formula whose cached result is the EMPTY STRING
    (=IF(x="","",...)) is stored by Excel as <c t="str"><f>..</f><v/></c>.
    openpyxl's cached load reads that as None, the same as a formula with no
    cached value at all, so every reader called the result absent and told
    the caller to recalculate, forever, even after Excel's own CalculateFull.
    In a cached (data_only) load the cell keeps the 'str' type, which a
    formula with no cached value never has: those cells get their real
    cached value, "", back. Returns how many were filled."""
    n = 0
    for ws in wb.worksheets:
        for cell in list(getattr(ws, "_cells", {}).values()):
            if cell.value is None and cell.data_type == "str":
                cell._value = ""
                n += 1
    return n


def open_wb(path: str, *, data_only: bool = False):
    """Load a workbook for reading. Sandbox-gated; keep_vba for .xlsm and .xltm so a read
    of a macro workbook does not have to strip anything. Missing file refuses
    with WorkbookNotFound (NOT_FOUND), not a raw error."""
    p = check_path(path, "open workbook")
    # A workbook a save which did not finish left aside is named
    # in the answer; a read never moves it back.
    from ..core import swapwait as _swapwait
    _swapwait.require_workbook(p, f"no such workbook: {p}")
    if os.path.isdir(p):
        raise XlMcpError(
            f"{p} is a directory, not a workbook file; pass the path of an "
            ".xlsx/.xlsm file")
    # A legacy .xls (or an encrypted package) is an OLE container; refuse it
    # here with the honest format-and-remedy message rather than leak
    # openpyxl's raw InvalidFileException, or half-parse the zip fragment
    # modern .xls files embed.
    _hazard.refuse_ole_container(p)
    import zipfile
    keep_vba = _xmlsafe.keeps_vba(p)
    try:
        # A part that declares a document type is refused before
        # openpyxl reads anything (core/xmlsafe.py), whichever part it is in.
        _swapwait.through_swap(p, lambda: _xmlsafe.check_file(p))
        wb = _swapwait.through_swap(p, lambda: _xmlsafe.load_workbook(
            p, data_only=data_only, keep_vba=keep_vba, rich_text=False))
        if data_only:
            fill_empty_text_results(wb)
        return wb
    except zipfile.BadZipFile as exc:
        # Reads used to surface this as a raw "File is not a zip file" tool
        # error with no envelope; mutations already refused properly through
        # the hazard scan.
        raise WorkbookCorrupt(
            f"{p}: not a valid .xlsx/.xlsm package (not a readable zip). "
            f"{exc}") from exc
    except (TypeError, ValueError, IndexError, SyntaxError) as exc:
        # openpyxl's reader raises a bare TypeError/ValueError on markup it
        # cannot load (a <row r="1.0"> carrying a height), an IndexError for
        # a cell pointing past the shared strings, and the XML parser's
        # SyntaxError (ParseError) for a part cut off mid-way; a read
        # refuses plainly, as a damaged workbook, rather than letting the
        # Python error escape as INTERNAL_ERROR or as the caller's fault.
        raise WorkbookCorrupt(_damaged_text(p, exc)) from exc


def _damaged_text(p: str, exc: BaseException) -> str:
    """The refusal for a workbook whose markup the reader cannot load."""
    if isinstance(exc, IndexError):
        what = "a cell or style points at an entry the workbook lacks "
    elif isinstance(exc, SyntaxError):
        what = "a part's XML is damaged or cut off "
    else:
        what = ""
    return (f"{p}: this reader cannot load the workbook; {what}"
            f"({type(exc).__name__}: {exc})")


def resolve(wb, location: Any, *, default_sheet: str | None = None,
            path: str | None = None):
    """Resolve a location object to a ResolvedGrid rectangle.

    path is the anchor-consistency escape hatch: a get_grid_view anchor's
    fingerprint lives in the RAW value space (formula strings, the
    data_only=False load every mutating tool resolves against), so a read
    that resolved against a data_only=True load would recompute the
    fingerprint over cached values and refuse a perfectly fresh anchor.
    When path is given and the location is an anchor against a cached
    load, resolution re-opens a formula-load view just for the anchor
    check."""
    if (path is not None and isinstance(location, dict)
            and "anchor" in location and getattr(wb, "data_only", False)):
        fwb = open_wb(path, data_only=False)
        try:
            return _locate.resolve_location(
                fwb, location, default_sheet=default_sheet)
        finally:
            fwb.close()
    return _locate.resolve_location(wb, location, default_sheet=default_sheet)


def guard_cell_count(grid, *, ceiling: int = MAX_READ_CELLS,
                     advice: str | None = None) -> int:
    n = (grid.max_row - grid.min_row + 1) * (grid.max_col - grid.min_col + 1)
    if n > ceiling:
        raise CeilingExceeded(
            f"the resolved range {grid.a1} on {grid.sheet!r} covers {n:,} "
            f"cells, over the {ceiling:,}-cell ceiling; "
            + (advice or "page it (pass a smaller range, or use xl_read query_range, which "
               "filters and summarizes on the server and returns only what you ask for)"))
    return n


def a1(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


@dataclass
class CellRead:
    """One cell as read back, with the honest label."""
    value: Any
    label: str


def broken_shared_fields(path: str, sheet: str | None = None, *,
                         rect: tuple | None = None, cells=None,
                         qualify: bool = False) -> dict:
    """The flag a read carries when the cells it returns include shared-formula
    members with no usable first cell (core/sharedformulas.py): no formula
    could be recovered, and the value returned is the stored one. `rect` is
    (min_row, min_col, max_row, max_col) on `sheet`; `cells` a set of
    (sheet, row, col) for a scatter read; `qualify` writes "Sheet!B2"."""
    from openpyxl.utils.cell import coordinate_to_tuple
    found = _shared.broken_cells(path)
    if not found:
        return {}
    hits: list[str] = []
    for title, coords in found.items():
        for coord in coords:
            r, c = coordinate_to_tuple(coord)
            if cells is not None:
                if (title, r, c) not in cells:
                    continue
            elif rect is not None:
                if title != sheet or not (rect[0] <= r <= rect[2] and rect[1] <= c <= rect[3]):
                    continue
            elif sheet is not None and title != sheet:
                continue
            hits.append(f"{title}!{coord}" if qualify else coord)
    if not hits:
        return {}
    out = {"cells": hits[:50], "note": _shared.NOTE}
    if len(hits) > 50:
        out["count"] = len(hits)
    return {"broken_shared_formulas": out}


def formula_mask(path: str, grid) -> dict[tuple[int, int], str] | None:
    """Which cells of a rectangle actually hold a formula, and its text.

    A data_only=True load cannot tell an uncalculated formula cell from a
    blank one, which is how the DEFAULT read mode ('cached') used to hand
    back a silent blank where a number belongs. This probe answers that
    question over the rectangle only, streaming (read_only) so the default
    read path does not pay for a second full model build.

    Returns None if the probe could not run; a read must degrade to its old
    unlabelled behavior rather than fail.
    """
    wb = None
    try:
        p = check_path(path, "probe formulas")
        wb = _xmlsafe.load_workbook(
            p, data_only=False, read_only=True,
            keep_vba=_xmlsafe.keeps_vba(p))
        ws = wb[grid.sheet]
        out: dict[tuple[int, int], str] = {}
        for row in ws.iter_rows(min_row=grid.min_row, max_row=grid.max_row,
                                min_col=grid.min_col, max_col=grid.max_col):
            for cell in row:
                text = _calc.formula_text_of(cell)
                if text is not None:
                    out[(cell.row, cell.column)] = text
        return out
    except Exception:  # noqa: BLE001
        return None
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def formula_mask_cells(path: str, sheet: str, coords
                       ) -> dict[tuple[int, int], str] | None:
    """The same honest-label probe as formula_mask, over a SET of scattered
    cells on one sheet, in ONE workbook open.

    get_cells used to reach read_matrix (and therefore formula_mask) once per
    requested cell, which is one full openpyxl parse per cell: 500 cells cost
    501 opens and 5.3 seconds where the equivalent read_range cost 2 opens and
    31 ms. Every one of those parses re-read
    the same unchanged file, so the redundancy was pure: the labels the probe
    produces are identical, only the opens go away.

    Bounded by the requested ROWS rather than by their bounding rectangle, and
    clamped to the sheet's own last row, so a scatter of A1 and A900000 does
    not stream the empty space between them.

    Returns None if the probe could not run, matching formula_mask: a read
    degrades to its old unlabelled behavior rather than failing.
    """
    if not coords:
        return {}
    wb = None
    want = set(coords)
    lo = min(r for r, _c in want)
    hi = max(r for r, _c in want)
    try:
        p = check_path(path, "probe formulas")
        wb = _xmlsafe.load_workbook(
            p, data_only=False, read_only=True,
            keep_vba=_xmlsafe.keeps_vba(p))
        ws = wb[sheet]
        last = ws.max_row
        if isinstance(last, int) and last > 0:
            hi = min(hi, last)
        if hi < lo:
            return {}
        out: dict[tuple[int, int], str] = {}
        for row in ws.iter_rows(min_row=lo, max_row=hi):
            for cell in row:
                # A blank cell streams as an EmptyCell, which has no address;
                # reading one raised here and the whole probe "could not
                # answer", so get_cells' default mode lost its labels (and
                # its stale-results note) on any row with a gap.
                r = getattr(cell, "row", None)
                if r is None:
                    continue
                key = (r, cell.column)
                if key not in want:
                    continue
                text = _calc.formula_text_of(cell)
                if text is not None:
                    out[key] = text
        return out
    except Exception:  # noqa: BLE001
        return None
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


#: read_matrix computes its own formula probe unless the caller passes one.
#: A caller reading MANY rectangles out of one workbook (get_cells) passes a
#: single probe covering all of them; None is a legitimate value (the probe
#: ran and could not answer), which is why the default is a sentinel.
NO_MASK = object()


class DisplayNotes:
    """Displayed text a read cannot give honestly: the cells whose format
    names a locale or calendar Excel's text was not measured for. Each such
    cell reads as its value, and the answer says so (note()), rather than
    showing a guessed text such as English day names for a Korean format."""

    def __init__(self, date1904: bool = False, path: str | None = None) -> None:
        self.cells: dict[str, list[str]] = {}
        self.date1904 = date1904
        self.path = path
        self.ties: dict[tuple[str, int, int], str | None] = {}

    def text(self, value, fmt, addr: str, sheet: str | None = None,
             row: int | None = None, col: int | None = None):
        """The cell's shown text. With its sheet, row and column, a time
        held to the millisecond at exactly half of its last shown unit is
        noted for settle(), which reads the number the file stores."""
        from ..core import numfmt as _numfmt
        out, why = _numfmt.display_checked(value, fmt, date1904=self.date1904)
        if why is None:
            if self.path and sheet is not None and row is not None and \
                    _numfmt.needs_stored_digits(value, fmt, date1904=self.date1904):
                self.ties[(sheet, row, col)] = fmt
            return out
        self.cells.setdefault(why, []).append(addr)
        return value

    def settle(self) -> dict[tuple[str, int, int], str]:
        """{(sheet, row, col): text} for the cells text() noted, from the
        numbers the file stores (one streaming pass per sheet). A cell whose
        stored number cannot be read keeps the text it was given."""
        from ..core import numfmt as _numfmt
        by_sheet: dict[str, dict[tuple[int, int], str | None]] = {}
        for (sheet, row, col), fmt in self.ties.items():
            by_sheet.setdefault(sheet, {})[(row, col)] = fmt
        self.ties = {}
        out: dict[tuple[str, int, int], str] = {}
        for sheet, cells in by_sheet.items():
            stored = stored_numbers(self.path, sheet, set(cells))
            for (row, col), fmt in cells.items():
                raw = stored.get((row, col))
                if raw is None:
                    continue
                text, _why = _numfmt.shown(raw, fmt, date1904=self.date1904)
                if text is not None:
                    out[(sheet, row, col)] = text
        return out

    def settle_into(self, matrix: list[list[Any]], sheet: str, min_row: int,
                    min_col: int) -> None:
        """settle(), written into a rectangular matrix read from sheet."""
        for (s, row, col), text in self.settle().items():
            i, j = row - min_row, col - min_col
            if s == sheet and 0 <= i < len(matrix) and 0 <= j < len(matrix[i]):
                matrix[i][j] = text

    def note(self) -> str | None:
        if not self.cells:
            return None
        parts = []
        for why, addrs in self.cells.items():
            shown = ", ".join(addrs[:10])
            more = f" (+{len(addrs) - 10} more)" if len(addrs) > 10 else ""
            parts.append(f"{shown}{more}: {why}")
        return ("Excel shows these cells with text this reader cannot "
                "produce, so they hold their values instead of a guessed "
                "text: " + "; ".join(parts))


#: A default read lists at most this many cells in `shows`; past that it says
#: more cells differ and that values='displayed' returns them all. Measured
#: (2026-10, the synthetic corpus): most reads list none; a 50-row sheet of
#: dates, prices and percentages lists 196 cells in its grid view, about 750
#: tokens at full length, about 420 at this cap.
SHOWS_LISTED = 100

#: Said when a read lists only the first SHOWS_LISTED cells in `shows`.
SHOWS_MORE_NOTE = (
    "shows lists the first {listed} cells whose text in Excel differs from their value, "
    "and more cells differ. values='displayed' returns the text of every cell")


def value_text(v: Any) -> str:
    """A value as the reads return it, as text (what a reader sees)."""
    return str(compact_value(v))


def _plain(value: Any, fmt: str | None) -> bool:
    """A value whose shown text is its own text without applying the format:
    text under General or @, and a whole number under General that fits
    General's eleven characters."""
    f = (fmt or "General").strip().lower()
    if isinstance(value, str):
        return f in ("general", "@", "")
    if f in ("general", "") and isinstance(value, int) and not isinstance(value, bool):
        return -10**10 < value < 10**11
    return False


class ShownText:
    """The text Excel shows for each read cell where it differs from the
    value the read returns (0.75 under # ?/? shows " 3/4", a date under
    d-mmm-yy "14-Jan-26", 1234.5 under #,##0.00 "1,234.50"), so a default
    read answers what a cell shows without values='displayed'. A cell whose
    text is its value adds nothing. A cell whose text depends on a locale or
    calendar this reader cannot render adds nothing either and is named in
    display_note. A time held to the millisecond at exactly half of the last
    unit its format shows is settled from the number the file stores.

    Keys are bare addresses, or sheet-qualified ones when the read spans
    sheets (qualify=True). A read lists the first SHOWS_LISTED cells and
    then works out no more of them, so a long read of formatted cells costs
    no more than a short one; every=True (find_cells, whose page bounds it)
    works out every cell added."""

    def __init__(self, path: str | None, date1904: bool = False, *,
                 qualify: bool = False, every: bool = False) -> None:
        self.path = path
        self.date1904 = date1904
        self.qualify = qualify
        self.every = every
        self.found: dict[tuple[str, int, int], str | None] = {}
        self.pending: dict[str, dict[tuple[int, int], tuple[Any, str]]] = {}
        self.notes = DisplayNotes(date1904)
        self.differ = 0
        self.more = False

    def add(self, sheet: str, row: int, col: int, value: Any, fmt: str | None) -> None:
        if self.more or value is None or isinstance(value, bool) or _plain(value, fmt):
            return
        from ..core import numfmt as _numfmt
        key = (sheet, row, col)
        if _numfmt.needs_stored_digits(value, fmt, date1904=self.date1904):
            self.found[key] = None  # keeps the cell's place in the order
            self.pending.setdefault(sheet, {})[(row, col)] = (value, fmt)
            return
        text, why = _numfmt.shown(value, fmt, date1904=self.date1904)
        if why is not None:
            self.notes.cells.setdefault(why, []).append(self._name(sheet, row, col))
        elif text is not None and text != value_text(value):
            self.differ += 1
            if not self.every and self.differ > SHOWS_LISTED:
                self.more = True  # listed in full: the rest are not worked out
                return
            self.found[key] = text

    def _name(self, sheet: str, row: int, col: int) -> str:
        addr = a1(row, col)
        return _locate.qualify_a1(sheet, addr) if self.qualify else addr

    def _settle(self) -> None:
        """The pending cells, from the numbers the file stores."""
        from ..core import numfmt as _numfmt
        for sheet, cells in self.pending.items():
            stored = stored_numbers(self.path, sheet, set(cells)) if self.path else {}
            for (row, col), (value, fmt) in cells.items():
                key = (sheet, row, col)
                raw = stored.get((row, col))
                text = None
                if raw is not None:
                    text, _why = _numfmt.shown(raw, fmt, date1904=self.date1904)
                if text is not None and text != value_text(value):
                    self.found[key] = text
                else:
                    self.found.pop(key, None)
        self.pending = {}

    def text_for(self, sheet: str, row: int, col: int) -> str | None:
        self._settle()
        return self.found.get((sheet, row, col))

    def fields(self) -> dict:
        """{"shows": {address: text}} for the cells that differ (the first
        SHOWS_LISTED, then shows_more and its note), and display_note for
        the cells whose text this reader cannot give; {} when neither."""
        self._settle()
        out: dict[str, Any] = {}
        listed = [(k, t) for k, t in self.found.items() if t is not None]
        if listed:
            out["shows"] = {self._name(*k): t for k, t in listed[:SHOWS_LISTED]}
            if self.more or len(listed) > SHOWS_LISTED:
                out["shows_more"] = True
                out["shows_note"] = SHOWS_MORE_NOTE.format(listed=SHOWS_LISTED)
        note = self.notes.note()
        if note:
            out["display_note"] = note
        return out


def stored_numbers(path: str, sheet: str, coords: set[tuple[int, int]]
                   ) -> dict[tuple[int, int], float]:
    """The numbers the file stores for these cells of one sheet, as written.
    A reader hands a cell under a date or time format over as a date, a
    time or a duration held to the millisecond; the shown text of a few of
    them depends on the digits past that (ShownText). One streaming pass
    over the sheet's part, ending after the last row asked for; {} when the
    part cannot be read."""
    import zipfile

    from openpyxl.utils.cell import coordinate_to_tuple

    from ..core import surgical as _surgical

    if not coords:
        return {}
    last = max(r for r, _c in coords)
    out: dict[tuple[int, int], float] = {}
    try:
        with zipfile.ZipFile(check_path(path, "open workbook")) as zf:
            pkg = _surgical.Package(zf)
            part = pkg.sheet_parts().get(sheet)
            real = pkg.real(part) if part else None
            if real is None:
                return {}
            with zf.open(real) as fh:
                for _event, elem in _xmlsafe.iterparse(fh, part=real):
                    tag = elem.tag
                    if tag.endswith("}row") or tag == "row":
                        try:
                            done = int(elem.get("r") or 0) >= last
                        except ValueError:
                            done = False
                        elem.clear()
                        if done:
                            break
                        continue
                    if not (tag.endswith("}c") or tag == "c"):
                        continue
                    ref = elem.get("r")
                    if not ref or elem.get("t") not in (None, "n"):
                        continue
                    key = coordinate_to_tuple(ref)
                    if key not in coords:
                        continue
                    for child in elem:
                        if child.tag.endswith("}v") or child.tag == "v":
                            try:
                                out[key] = float(child.text)
                            except (TypeError, ValueError):
                                pass
    except Exception:  # noqa: BLE001 (the cells then say nothing)
        return out
    return out


def read_matrix(grid, *, mode: str, formula_wb=None, cached_wb=None,
                path: str | None = None, mask: Any = NO_MASK,
                display_notes: DisplayNotes | None = None,
                shown: ShownText | None = None, settle: bool = True
                ) -> tuple[list[list[Any]], list[list[str]], bool]:
    """Return (values, labels, has_formula) for a resolved rectangle.

    mode: cached -> last cached values, labelled cached | absent | value;
          formula -> formula strings / literals;
          both -> values are cached, labels tell formula cells apart and a
                  parallel formula read is folded into the label set.
    formula_wb is a data_only=False load; cached_wb is a data_only=True load.
    The caller supplies whichever loads the mode needs. In 'cached' mode,
    passing `path` buys the honest label (a streaming formula probe over the
    rectangle); without it the mode falls back to the old blind labelling.
    `mask` overrides that probe with one the caller already ran (see
    formula_mask_cells); None is a real value meaning "the probe could not
    answer", so the default is the NO_MASK sentinel, not None. `shown`
    collects the text Excel shows for the cached and both modes' values.
    """
    if mode not in VALUE_MODES:
        raise XlMcpError(
            f"values must be one of {VALUE_MODES}, got {mode!r}")
    values: list[list[Any]] = []
    labels: list[list[str]] = []
    has_formula = False
    fws = formula_wb[grid.sheet] if formula_wb is not None else None
    cws = cached_wb[grid.sheet] if cached_wb is not None else None
    if mask is NO_MASK:
        mask = formula_mask(path, grid) \
            if (mode in ("cached", "displayed") and path is not None) else None
    for r in range(grid.min_row, grid.max_row + 1):
        vrow: list[Any] = []
        lrow: list[str] = []
        for c in range(grid.min_col, grid.max_col + 1):
            formula = None
            if fws is not None:
                formula = _calc.formula_text_of(fws.cell(r, c))
            elif mask is not None:
                formula = mask.get((r, c))
            cached = cws.cell(r, c).value if cws is not None else None
            if formula is not None:
                has_formula = True
            if mode == "formula":
                # The value returned here IS the formula text, so the label is
                # LABEL_FORMULA -- the cache was never consulted and nothing is
                # stale. Labelling it 'absent' (what label_cell(formula, None)
                # returns) fired the absent-cache warning on a read that had
                # asked for formulas in the first place, and left the
                # LABEL_FORMULA constant dead.
                #
                # The text comes from formula_text_of, not cell.value: an
                # ARRAY formula's value is an openpyxl ArrayFormula OBJECT,
                # and returning it handed the transport a repr instead of a
                # formula. Every LAMBDA Excel itself authors is stored
                # t="array", so reading back a real Excel workbook was the
                # common case, not the corner. denormalize_formula then does
                # the job it was written for and never wired to: the reader
                # sees =LAMBDA(x,[y],...), which is what Excel's own formula
                # bar shows, instead of the stored
                # _xlfn.LAMBDA(_xlpm.x,_xlop.y,...).
                if formula is not None:
                    out = _calc.denormalize_formula(formula)
                    label = _calc.LABEL_FORMULA
                else:
                    out = fws.cell(r, c).value
                    label = _calc.LABEL_VALUE
            elif mode in ("cached", "displayed"):
                # The probe (when available) is what lets this mode say
                # 'absent' instead of handing back a silent blank; without
                # it nothing is over-claimed.
                out = cached
                label = _calc.label_cell(formula, cached) if mask is not None \
                    else _calc.LABEL_VALUE
                if mode == "displayed" and cws is not None:
                    fmt = cws.cell(r, c).number_format
                    if display_notes is not None:
                        out = display_notes.text(cached, fmt, a1(r, c),
                                                 grid.sheet, r, c)
                    else:
                        from ..core import numfmt as _numfmt
                        out = _numfmt.display(cached, fmt,
                                              date1904=is_1904(cws))
            else:  # both
                out = cached if formula is not None else (
                    cached if cached is not None else (
                        fws.cell(r, c).value if fws is not None else None))
                label = _calc.label_cell(formula, cached)
            if shown is not None and out is not None and mode in ("cached", "both"):
                src = cws if cws is not None else fws
                shown.add(grid.sheet, r, c, out, src.cell(r, c).number_format)
            vrow.append(out)
            lrow.append(label)
        values.append(vrow)
        labels.append(lrow)
    if display_notes is not None and settle:
        display_notes.settle_into(values, grid.sheet, grid.min_row, grid.min_col)
    return values, labels, has_formula


#: The one sentence every read surface says when it hands back a formula cell
#: that no engine has computed yet.
ABSENT_WARNING = (
    "some cells hold formulas with NO cached value (label 'absent') and read "
    "as blank to every non-Excel consumer; run recalculate (xl_app) or open "
    "in Excel to populate them before trusting these numbers")


#: Said beside cached results when the file is flagged to recalculate on
#: open: a program changed inputs it could not recalculate (this product
#: sets the flag whenever its own edit may have left a result out of date).
PENDING_NOTE = (
    "this workbook is flagged to recalculate when Excel next opens it, so "
    "some cached results may be out of date; run recalculate (xl_app) for "
    "current values")


def pending_recalc_note(path: str | None) -> str | None:
    """PENDING_NOTE when the package's calcPr asks for a full recalculation
    on open, else None. Reads only the workbook part."""
    if not path:
        return None
    try:
        from ..core import pkginfo as _pkginfo
        from ..core.sandbox import check_path as _check_path
        if _pkginfo.calc_state(_check_path(path, "open workbook"))[
                "recalculates_on_open"]:
            return PENDING_NOTE
    except Exception:  # noqa: BLE001
        return None
    return None


def pending_note_for(path: str | None, grid=None) -> str | None:
    """PENDING_NOTE for a read that returned cached formula results from a
    workbook flagged to recalculate on open (query_range,
    export_range, get_grid_view and find_cells used those results with no
    note, while read_range and get_cells said so). With a grid, the note is
    given only when the rectangle holds a formula, or when the probe cannot
    tell."""
    note = pending_recalc_note(path)
    if not note or grid is None:
        return note
    mask = formula_mask(path, grid)
    if mask is None or mask:
        return note
    return None


#: The label an ordinary cell carries: a literal value, nothing to report.
#: It is what an address NOT listed in a sparse label set carries, which is
#: why the wire shape names it in `labels_default` instead of assuming it.
LABEL_DEFAULT = _calc.LABEL_VALUE


def sparse_labels(labels, *, min_row: int = 1, min_col: int = 1,
                  sheet: str | None = None) -> dict[str, list[str]]:
    """Group a dense label matrix by label, listing only the addresses whose
    label is not the default.

    The labelling itself is the honest-calc contract and is untouched; this
    is its ENCODING. A dense parallel matrix spends one entry per cell to say
    "value" about cells that are 87% ordinary: on a 200x8 read it was 13,478
    of the payload's 25,757 characters, 52%, of which 1,400 of 1,608 entries
    were the literal string "value".

    Grouped-by-label is lossless against the dense form (anything unlisted
    carries LABEL_DEFAULT) and is never larger: an address costs fewer
    characters than a quoted label plus its comma, so the worst case, where
    no cell is ordinary, still comes out ahead.

    `sheet` qualifies every address ('Data!B7') for callers whose cells can
    span sheets; omitted, the addresses are bare.
    """
    out: dict[str, list[str]] = {}
    for r, row in enumerate(labels):
        for c, label in enumerate(row):
            if label == LABEL_DEFAULT:
                continue
            addr = a1(min_row + r, min_col + c)
            if sheet is not None:
                addr = _locate.qualify_a1(sheet, addr)
            out.setdefault(label, []).append(addr)
    return out


def absent_note(labels) -> str | None:
    """The absent-cache warning if any label in a matrix says 'absent'."""
    for row in labels:
        if _calc.LABEL_ABSENT in row:
            return ABSENT_WARNING
    return None


#: What a formula-mode read says about text cells that start with "=".
TEXT_NOT_FORMULA_NOTE = (
    "these cells hold text that starts with '=', not a formula. To write one "
    "back as text, pass {\"text\": \"=...\"} or start it with an apostrophe "
    "('=...); a plain '=...' string, or set_formula, would turn the text "
    "into a live formula")
_TEXT_LISTED = 100


def text_not_formulas(addresses: list[str]) -> dict:
    """The keys a formula-mode read adds for text cells that start with "="
    (a formula and such text read the same there): the addresses, at most
    100, and the note. {} when there are none."""
    if not addresses:
        return {}
    note = TEXT_NOT_FORMULA_NOTE
    if len(addresses) > _TEXT_LISTED:
        note += f" ({len(addresses) - _TEXT_LISTED:,} more are not listed)"
    return {"text_not_formulas": addresses[:_TEXT_LISTED], "text_note": note}


def is_1904(ws) -> bool:
    """Whether the worksheet's workbook counts dates from 1904."""
    from openpyxl.utils.datetime import MAC_EPOCH
    return getattr(getattr(ws, "parent", None), "epoch", None) == MAC_EPOCH


def elapsed_text(v) -> str:
    """A duration as Excel's elapsed time shows it ([h]:mm:ss): 1.75 days is
    42:00:00, not Python's "1 day, 18:00:00"."""
    total = v.total_seconds()
    sign = "-" if total < 0 else ""
    whole, frac = divmod(abs(total), 1)
    h, rest = divmod(int(whole), 3600)
    m, s = divmod(rest, 60)
    text = f"{sign}{h}:{m:02d}:{s:02d}"
    if frac >= 5e-7:
        text += f"{frac:.6f}"[1:].rstrip("0")
    return text


def compact_value(v: Any) -> Any:
    """Render a value for a token-efficient projection: dates/datetimes to ISO,
    durations as Excel's elapsed time, everything else passed through (JSON
    handles str/num/bool/None)."""
    import datetime as _dt
    if isinstance(v, _dt.timedelta):
        return elapsed_text(v)
    if isinstance(v, _dt.datetime) and v.time() == _dt.time():
        # A date with no time part reads as the date
        # ("2026-01-15", not "2026-01-15T00:00:00"), as Excel shows it.
        return v.date().isoformat()
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time)):
        return v.isoformat()
    return v


__all__ = [
    "MAX_READ_CELLS", "MAX_WRITE_CELLS", "MAX_SCAN_CELLS", "VALUE_MODES", "CACHED_MODES", "open_wb", "resolve", "guard_cell_count",
    "a1", "CellRead", "read_matrix", "compact_value", "formula_mask",
    "formula_mask_cells", "NO_MASK", "absent_note", "ABSENT_WARNING",
    "sparse_labels", "LABEL_DEFAULT",
]
