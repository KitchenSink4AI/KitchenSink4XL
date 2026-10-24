"""core/refs.py: reference rewriting on structural edits.

One of the riskiest modules in the product. When rows or columns are inserted
or deleted, or a range is moved,
every formula, defined-name definition, data-validation range, conditional-
format range, chart source ref, table ref, and merged range that points at or
across the change MUST be rewritten so the workbook stays coherent. No file-
based library does this; it is
central infrastructure, not per-tool logic.

The module has two layers:

  1. A REFERENCE TRANSPOSER (pure, heavily unit-tested): parse A1 references in
     a formula string, including absolute ``$`` anchors, sheet-qualified and
     3D refs, and ranges; shift each in-scope reference on an insert / delete /
     move; leave out-of-scope references untouched; a reference whose target is
     wholly deleted becomes ``#REF!`` (Excel semantics).

  2. A WORKBOOK REWRITER: apply one structural edit across an open openpyxl
     workbook, rewriting formula cells, defined names, per-sheet conditional-
     format and data-validation ranges, table refs, and merged ranges, and
     report a count per kind.

Every rule is checked against Excel's own result on the same file. A
structural edit also rewrites
conditional-format RULE formulas, data-validation formulas, chart series,
category and title references, the print area and print titles, the
autofilter and pivot cache sources; a 3D reference is never shifted by an edit
on one sheet of its span (Excel leaves it); renaming a sheet rewrites every
qualifier that names it, quoting the new name where Excel does
(rename_sheet_refs); deleting a sheet turns references to it into #REF! and
shrinks a 3D span whose end it was (delete_sheet_refs); renaming a table
rewrites every structured reference to it (rename_table_refs). Hyperlink
locations are left alone in every case, as Excel leaves them.

Whole-column (``A:A``, ``$B:$D``) and whole-row (``1:1``, ``$1:$2``) spans ARE
transposed on their own axis (this used to be a gap: ``=SUM(B:B)`` went
stale on every column insert/delete, and print-title names like ``$1:$2`` went
stale on row inserts); an edit on the other axis leaves them untouched, which
is Excel's own behavior. Known bounded gaps, stated honestly (no warranty
language): 3D references (``Sheet1:Sheet3!A1``) are transposed only when the
edited sheet is a named endpoint of the span; structured table references
(``Table[Col]``) inside formulas are not coordinate-shifted (they follow the
table, which the rewriter resizes separately); partial-overlap MOVE of a range
and MOVE over a whole-row/column span are left untouched. The transposer never
guesses: an ambiguous construct is passed through unchanged rather than
corrupted.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field, replace as _replace

from openpyxl.utils import column_index_from_string, get_column_letter

from . import calc as _calc

# --------------------------------------------------------------- edit model

INSERT_ROWS = "insert_rows"
DELETE_ROWS = "delete_rows"
INSERT_COLS = "insert_cols"
DELETE_COLS = "delete_cols"
MOVE = "move"

_ROW_KINDS = (INSERT_ROWS, DELETE_ROWS)
_COL_KINDS = (INSERT_COLS, DELETE_COLS)
_INSERT_KINDS = (INSERT_ROWS, INSERT_COLS)
_DELETE_KINDS = (DELETE_ROWS, DELETE_COLS)

REF_ERROR = "#REF!"

MAX_ROW = 1_048_576
MAX_COL = 16_384


@dataclass(frozen=True)
class RefEdit:
    """One structural edit, the unit the transposer and rewriter both consume.

    sheet: the sheet the structural change happens on. Only references whose
        effective sheet resolves to this sheet are shifted; a formula living on
        another sheet shifts only its refs that are sheet-qualified to `sheet`.
    kind: insert_rows | delete_rows | insert_cols | delete_cols | move.
    index: 1-based row or column where the insert/delete starts (row/col kinds).
    count: number of rows or columns inserted/deleted (row/col kinds).
    src / dst: MOVE only. src is (min_row, min_col, max_row, max_col); dst is
        the (row, col) top-left the block moves to. References that fall wholly
        inside src are translated by the src->dst delta.
    grow: INSERT_ROWS only. Each table the insert grows, as (top_row,
        first_row, last_row, bottom_row, min_col, max_col) before the edit:
        the header row, the first and last data rows, the totals row (or the
        last data row when there is none) and the table's columns. A
        reference that reaches over all of first_row..last_row, stays between
        top_row and bottom_row and lies within min_col..max_col is extended
        by the inserted rows instead of being shifted or left alone, as Excel
        does when a table grows (a total written as SUM(D2:D13) over a
        table's column becomes SUM(D2:D14) when a row is added at the end, at
        the totals row or at the first data row).
    """

    sheet: str
    kind: str
    index: int = 0
    count: int = 0
    src: tuple[int, int, int, int] | None = None
    dst: tuple[int, int] | None = None
    grow: tuple[tuple[int, int, int, int, int, int], ...] = ()

    def __post_init__(self) -> None:
        if self.kind not in (INSERT_ROWS, DELETE_ROWS, INSERT_COLS,
                             DELETE_COLS, MOVE):
            raise ValueError(f"unknown structural edit kind {self.kind!r}")
        if self.kind == MOVE:
            if self.src is None or self.dst is None:
                raise ValueError("move edit requires src and dst")
        else:
            if self.index < 1:
                raise ValueError("index is 1-based, got " + str(self.index))
            if self.count < 1:
                raise ValueError("count must be >= 1, got " + str(self.count))
        if self.grow and self.kind != INSERT_ROWS:
            raise ValueError("grow applies only to insert_rows")


# ---------------------------------------------------- line-shift primitives


def _shift_single(pos: int, edit: RefEdit) -> int | None:
    """Shift one coordinate on the edited axis. Returns the new coordinate, or
    None when the coordinate sat in a deleted band (a wholly-deleted target)."""
    if edit.kind in _INSERT_KINDS:
        return pos + edit.count if pos >= edit.index else pos
    # delete
    a = edit.index
    b = edit.index + edit.count - 1
    if pos < a:
        return pos
    if pos > b:
        return pos - edit.count
    return None  # inside the deleted band


def _shift_span(lo: int, hi: int, edit: RefEdit) -> tuple[int, int] | None:
    """Shift a [lo, hi] span on the edited axis. Returns the new (lo, hi), or
    None when the whole span was deleted. On a partial delete the span is
    clamped to the surviving rows/cols (Excel's shrink behavior)."""
    if edit.kind in _INSERT_KINDS:
        new_lo = lo + edit.count if lo >= edit.index else lo
        new_hi = hi + edit.count if hi >= edit.index else hi
        return new_lo, new_hi
    a = edit.index
    b = edit.index + edit.count - 1
    if lo >= a and hi <= b:
        return None  # entire span removed
    new_lo = lo if lo < a else (lo - edit.count if lo > b else a)
    new_hi = hi if hi < a else (hi - edit.count if hi > b else a - 1)
    return new_lo, new_hi


# ------------------------------------------------------ single-ref transform


@dataclass(frozen=True)
class _Endpoint:
    col_abs: bool
    col: int
    row_abs: bool
    row: int


def _parse_cell(token: str) -> _Endpoint | None:
    m = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", token)
    if not m:
        return None
    col = column_index_from_string(m.group(2).upper())
    row = int(m.group(4))
    if col > MAX_COL or row > MAX_ROW:
        return None
    return _Endpoint(bool(m.group(1)), col, bool(m.group(3)), row)


def _render_cell(ep: _Endpoint) -> str:
    return "{}{}{}{}".format(
        "$" if ep.col_abs else "", get_column_letter(ep.col),
        "$" if ep.row_abs else "", ep.row,
    )


def _in_scope(ref_sheet: str | None, formula_sheet: str | None,
              edit_sheet: str) -> bool:
    """A reference is in scope when its effective sheet is the edited sheet.
    An unqualified ref inherits the formula's home sheet."""
    effective = ref_sheet if ref_sheet is not None else formula_sheet
    if effective is None:
        # No home sheet context (a bare ref transform with no formula_sheet):
        # treat as in scope so standalone range strings on the edited sheet
        # still transpose.
        return True
    return _norm_sheet(effective) == _norm_sheet(edit_sheet)


def _norm_sheet(name: str) -> str:
    n = name.strip()
    if n.startswith("'") and n.endswith("'") and len(n) >= 2:
        n = n[1:-1].replace("''", "'")
    return n.lower()


def _move_point(row: int, col: int, edit: RefEdit) -> tuple[int, int] | None:
    """Translate a point if it falls inside the MOVE source rectangle."""
    smin_r, smin_c, smax_r, smax_c = edit.src  # type: ignore[misc]
    dr = edit.dst[0] - smin_r  # type: ignore[index]
    dc = edit.dst[1] - smin_c  # type: ignore[index]
    if smin_r <= row <= smax_r and smin_c <= col <= smax_c:
        return row + dr, col + dc
    return None


def _covers_table_body(e1: _Endpoint, e2: _Endpoint, edit: RefEdit) -> bool:
    """True when the range e1:e2 reaches over the whole data body of a table
    this insert grows, without leaving the table's rows or columns. Measured
    in Excel: such a range grows with the table; a range that covers only part
    of the data (C5:C13 of C4:C13, or C4:C12) does not, and neither does one
    that runs past the table's last row or into a column outside it."""
    r_lo, r_hi = sorted((e1.row, e2.row))
    c_lo, c_hi = sorted((e1.col, e2.col))
    return any(top <= r_lo <= first and last <= r_hi <= bottom
               and c1 <= c_lo and c_hi <= c2
               for top, first, last, bottom, c1, c2 in edit.grow)


def _transform_endpoints(
    e1: _Endpoint, e2: _Endpoint | None, edit: RefEdit
) -> tuple[_Endpoint, _Endpoint | None] | None:
    """Apply an edit to one or two endpoints. Returns the new endpoints, or
    None to signal the whole reference collapses to #REF!."""
    if edit.kind == MOVE:
        if e2 is None:
            moved = _move_point(e1.row, e1.col, edit)
            if moved is None:
                return e1, None
            return _Endpoint(e1.col_abs, moved[1], e1.row_abs, moved[0]), None
        # range: translate only when the whole rectangle is inside src.
        m1 = _move_point(e1.row, e1.col, edit)
        m2 = _move_point(e2.row, e2.col, edit)
        if m1 is None or m2 is None:
            return e1, e2  # partial overlap or outside: leave untouched
        return (
            _Endpoint(e1.col_abs, m1[1], e1.row_abs, m1[0]),
            _Endpoint(e2.col_abs, m2[1], e2.row_abs, m2[0]),
        )

    axis_row = edit.kind in _ROW_KINDS
    if e2 is None:
        pos = e1.row if axis_row else e1.col
        shifted = _shift_single(pos, edit)
        if shifted is None:
            return None
        if axis_row:
            return _Endpoint(e1.col_abs, e1.col, e1.row_abs, shifted), None
        return _Endpoint(e1.col_abs, shifted, e1.row_abs, e1.row), None

    lo = (e1.row if axis_row else e1.col)
    hi = (e2.row if axis_row else e2.col)
    span = _shift_span(min(lo, hi), max(lo, hi), edit)
    if span is None:
        return None
    if edit.grow and axis_row and edit.kind == INSERT_ROWS and             _covers_table_body(e1, e2, edit):
        span = (min(lo, hi), max(lo, hi) + edit.count)
    new_lo, new_hi = span
    # A range reaching the sheet's last row or column is clamped
    # at the edge, as Excel does (A1:A1048576 after an insert is still
    # A1:A1048576); only a range pushed wholly off the grid is #REF!. The
    # unclamped end raised "Max value is 1048576" from openpyxl.
    limit = MAX_ROW if axis_row else MAX_COL
    if new_lo > limit:
        return None
    new_hi = min(new_hi, limit)
    # preserve endpoint order (lo endpoint keeps e1's abs flags)
    if lo <= hi:
        p1, p2 = new_lo, new_hi
    else:
        p1, p2 = new_hi, new_lo
    if axis_row:
        return (
            _Endpoint(e1.col_abs, e1.col, e1.row_abs, p1),
            _Endpoint(e2.col_abs, e2.col, e2.row_abs, p2),
        )
    return (
        _Endpoint(e1.col_abs, p1, e1.row_abs, e1.row),
        _Endpoint(e2.col_abs, p2, e2.row_abs, e2.row),
    )


# ------------------------------------------------------ formula tokenizer

# String literals: double-quoted, with "" as the escape. Kept for callers that
# blank strings out of a formula; the rewriters below use the lexer instead,
# because a double quote inside a quoted sheet name ('North "Q"'!A1) is part
# of the name, not the start of a string.
_STRING_RE = re.compile(r'"(?:[^"]|"")*"')

# Characters that end a name or an unquoted sheet name: Excel's operators,
# separators, brackets and quotes, the whitespace between tokens, and the two
# spaces Excel quotes a sheet name for (no-break and ideographic). Every other
# character, in any script, can stand in a defined name or in a sheet name
# Excel writes without quotes. Measured with Excel's own save: Données!A1,
# 売上!A1, ยอดขาย!A1, Cost€!A1 and ab·c!A1 are stored unquoted, and the
# names αA1, A1α, 売上A1, a?b and a\b are accepted as defined names.
_ENDS_TEXT = " \t\r\n!\"#$%&'()*+,-/:;<=>@[]^`{|}~ 　"
_CONTROL = r"\x00-\x1f\x7f"
_ENDS = re.escape(_ENDS_TEXT) + _CONTROL
_NAME_CH = "[^" + _ENDS + "]"
# An unquoted sheet name: a name character that is not a digit or a period,
# then any name characters (Excel quotes '2024'!A1 and '.x'!A1).
_UNQUOTED_SHEET = "[^" + _ENDS + "0-9.]" + _NAME_CH + "*"
_QUOTED_SHEET = r"'(?:[^']|'')*'"

# A sheet qualifier: 'Quoted Name'! or BareName! optionally 3D (Sheet1:Sheet3!).
_SHEET = (
    "(?:" + _QUOTED_SHEET + "|" + _UNQUOTED_SHEET + ")"
    "(?::(?:" + _QUOTED_SHEET + "|" + _UNQUOTED_SHEET + "))?!"
)
# Row numbers are ASCII digits only: Excel reads Q１ (a full-width digit) as a
# name, never as cell Q1, and Python's \d would match it.
_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]+"
# Whole-column (B:B, $A:$C) and whole-row (1:1, $2:$5) spans. Both endpoints
# must be the same shape, so "A1:B" or "A:B2" never half-match.
_COLSPAN = r"(?P<cs1>\$?[A-Za-z]{1,3}):(?P<cs2>\$?[A-Za-z]{1,3})"
_ROWSPAN = r"(?P<rs1>\$?[0-9]+):(?P<rs2>\$?[0-9]+)"
# What may stand just before a reference: anything that ends a name, except a
# quote or a ! (the reference would be the tail of a qualifier), a $ (it
# belongs to the reference) or a '[' (a structured reference's column). A
# reference after ']' (another workbook's prefix, [1]Data!A1) or '#' (an
# error, #REF!A1) still MATCHES, whole, so that no part of it can be read on
# its own; the rewriters then leave it as it is (_Hit.before).
_BEFORE_OK = re.escape(_ENDS_TEXT.replace("'", "").replace("!", "").replace("$", "")
                       .replace("[", "")) + _CONTROL
# What stands before a reference that belongs to another workbook or to an
# error: an edit in this workbook never changes it.
_NOT_OURS = ("]", "#")
# What may stand just after one: anything that ends a name, except '(' (a
# function call such as LOG10(), '[' (a structured reference), ! or a quote.
_AFTER_OK = re.escape(_ENDS_TEXT.replace("(", "").replace("[", "").replace("!", "")
                      .replace("'", "")) + _CONTROL
# After a table's name: the same, except that its brackets may follow.
_AFTER_TABLE_OK = re.escape(_ENDS_TEXT.replace("(", "").replace("!", "").replace("'", "")) + _CONTROL
# A reference: optional sheet prefix, then a cell (with optional ':cell'
# range), a whole-column span, or a whole-row span. Guards: not preceded by a
# name character in any script (so it is not a name's tail: αA1, ÀB12) and not
# followed by one (so it is not a name's head: A1α, Q1Data). The cell
# alternative is tried first, so 'A1:B2' is a cell range, never a truncated
# span.
_REF_RE = re.compile(
    "(?<![^" + _BEFORE_OK + "])"
    r"(?P<sheet>" + _SHEET + r")?"
    r"(?:"
    r"(?P<c1>" + _CELL + r")(?::(?P<c2>" + _CELL + r"))?"
    r"|" + _COLSPAN +
    r"|" + _ROWSPAN +
    r")"
    "(?![^" + _AFTER_OK + "])"
)

# The lexer. Formula text is read once, left to right, the way Excel reads it,
# into the runs where no reference can stand:
#   "..."  a string literal ("" escapes a quote).
#   [...]  a structured reference's specifier (Sales[Q1], [@Q1],
#          Sales[[#This Row],[Q1]]; brackets nest and ' escapes the next
#          character). Column names, never cells, whatever they look like
#          (a table's columns are often named Q1 or FY2024).
#   [n]    another workbook's prefix: an unnamed bracket that runs straight on
#          into a sheet name, a quote or a ! ([1]Data!A1, [1]!Total). Excel
#          stores every reference to another workbook this way, and an edit in
#          this workbook never moves one.
# A quoted sheet name ('North "Q"'!A1) is walked over whole, so a double
# quote, a bracket or a ! inside it is read as part of the name, and nothing
# inside it is ever read as a reference ('A1 B2 C'!Total names a sheet).
_STR, _BRK, _EXT, _QUO, _BAD = "str", "brk", "ext", "quo", "bad"
_ENDS_CH = re.compile("[" + _ENDS + "]")


def _bracket_close(text: str, i: int) -> int | None:
    """Index of the ']' closing the '[' at i, honouring ' escapes and nesting
    (Excel's structured-reference rules); None when unbalanced."""
    depth, k, n = 0, i, len(text)
    while k < n:
        c = text[k]
        if c == "'":
            k += 2
            continue
        if c == "[":
            depth += 1
        elif c == "]":
            depth -= 1
            if depth == 0:
                return k
        k += 1
    return None


def _is_name_char(ch: str) -> bool:
    return bool(ch) and not _ENDS_CH.fullmatch(ch)


def _lex(body: str) -> list[tuple[int, int, str]]:
    """The runs of formula text no reference can stand in, as (start, end,
    kind), plus the quoted sheet names (kind _QUO). Anything unterminated
    (a string, a quoted name or a bracket that never closes) protects the
    rest of the text (kind _BAD): an unreadable formula is never rewritten."""
    out: list[tuple[int, int, str]] = []
    i, n = 0, len(body)
    while i < n:
        ch = body[i]
        if ch in "\"'":
            j = i + 1
            while j < n:
                if body[j] == ch:
                    if j + 1 < n and body[j + 1] == ch:
                        j += 2
                        continue
                    break
                j += 1
            if j >= n:
                out.append((i, n, _BAD))
                break
            out.append((i, j + 1, _STR if ch == '"' else _QUO))
            i = j + 1
            continue
        if ch == "[":
            close = _bracket_close(body, i)
            if close is None:
                out.append((i, n, _BAD))
                break
            before = body[i - 1] if i > 0 else ""
            after = body[close + 1] if close + 1 < n else ""
            unnamed = not _is_name_char(before)
            ext = unnamed and (_is_name_char(after) or after in ("'", "!")
                               or body[i + 1:close].isdigit())
            out.append((i, close + 1, _EXT if ext else _BRK))
            i = close + 1
            continue
        i += 1
    return out


class _Read:
    """The lexer's reading of one formula: the view the patterns run over
    (every run no reference can stand in blanked to \\x01, the same length as
    the text, so a match in the view is a match at the same place in the
    text) and where each quoted sheet name starts."""

    __slots__ = ("text", "view", "quoted_starts")

    def __init__(self, text: str, *, other_books_move: bool = False):
        chars = list(text)
        self.quoted_starts: set[int] = set()
        for start, end, kind in _lex(text):
            if kind == _BAD:
                lo, hi = start, end
            elif kind == _EXT and other_books_move:
                # Copying a formula moves its relative references to other
                # workbooks too: the prefix reads as spaces, so the reference
                # after it is read and moved as usual.
                for k in range(start, end):
                    chars[k] = " "
                continue
            else:
                # A string keeps its quotes, a bracket its brackets, a quoted
                # name its quotes, so what stands beside them reads as before.
                # Another workbook's prefix keeps its ']', so the reference
                # after it is never read and never moved.
                lo, hi = start + 1, end - 1
                if kind == _QUO:
                    self.quoted_starts.add(start)
            for k in range(lo, hi):
                chars[k] = "\x01"
        self.text = text
        self.view = "".join(chars)

    def hits(self, pattern: re.Pattern, quote_group: str | None):
        """pattern's matches in the view that are real: a quoted name in
        quote_group must open at a quoted sheet name the lexer found (never
        at a closing quote)."""
        for m in pattern.finditer(self.view):
            hit = _Hit(m, self.text, self.view)
            q = hit.group(quote_group) if quote_group else None
            if q and q.startswith("'") and m.start(quote_group) not in self.quoted_starts:
                continue
            yield hit

    def sub(self, pattern: re.Pattern, quote_group: str | None, fn) -> str:
        """pattern.sub over the view, written into the text: fn(hit) gives
        the replacement for one match."""
        out: list[str] = []
        pos = 0
        for hit in self.hits(pattern, quote_group):
            out.append(self.text[pos:hit.start()])
            out.append(fn(hit))
            pos = hit.end()
        out.append(self.text[pos:])
        return "".join(out)


class _Hit:
    """One match, its groups read from the formula text itself (the view
    blanks the inside of quoted sheet names). before is the character the
    view shows just before it ('' at the start)."""

    __slots__ = ("_m", "_text", "before")

    def __init__(self, m: re.Match, text: str, view: str):
        self._m, self._text = m, text
        self.before = view[m.start() - 1] if m.start() > 0 else ""

    def group(self, name: int | str = 0) -> str | None:
        s, e = self._m.span(name)
        return None if s < 0 else self._text[s:e]

    def start(self, name: int | str = 0) -> int:
        return self._m.start(name)

    def end(self, name: int | str = 0) -> int:
        return self._m.end(name)


def find_references(text: str) -> list[_Hit]:
    """Every A1 reference in formula text that this workbook's structural
    edits move, with the groups of _REF_RE (sheet, c1, c2, cs1, cs2, rs1,
    rs2): none inside a string, a quoted sheet name, a structured
    reference's brackets, another workbook's reference or an error."""
    return [h for h in _Read(text).hits(_REF_RE, "sheet") if h.before not in _NOT_OURS]


def _sheet_scope(prefix: str | None) -> tuple[str | None, bool]:
    """From a matched sheet prefix (with trailing '!'), return (sheet_name,
    is_3d). sheet_name is the bare (unquoted) name; for a 3D span it is the
    first endpoint."""
    if not prefix:
        return None, False
    body = prefix[:-1]  # drop '!'
    is_3d = False
    # split on a top-level ':' not inside quotes
    if ":" in body and not (body.startswith("'") and body.count("'") >= 2
                            and ":" in body[1:body.rfind("'")]):
        pass
    first = body
    if ":" in body:
        # only a real 3D span if the ':' is outside quotes
        parts = _split_3d(body)
        if len(parts) == 2:
            is_3d = True
            first = parts[0]
    if first.startswith("'") and first.endswith("'"):
        first = first[1:-1].replace("''", "'")
    return first, is_3d


def _split_3d(body: str) -> list[str]:
    out: list[str] = []
    cur = []
    inq = False
    i = 0
    while i < len(body):
        ch = body[i]
        if ch == "'":
            inq = not inq
            cur.append(ch)
        elif ch == ":" and not inq:
            out.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
        i += 1
    out.append("".join(cur))
    return out


# ------------------------------------------------- whole-row / whole-col spans


def _parse_span_tok(tok: str, *, is_col: bool) -> tuple[bool, int] | None:
    """One endpoint of a B:B / 1:1 span -> (is_absolute, index), or None when
    the token is out of grid (left untouched by the caller)."""
    absolute = tok.startswith("$")
    body = tok[1:] if absolute else tok
    if is_col:
        idx = column_index_from_string(body.upper())
        if idx > MAX_COL:
            return None
    else:
        idx = int(body)
        if idx > MAX_ROW:
            return None
    return absolute, idx


def _render_span_tok(absolute: bool, idx: int, *, is_col: bool) -> str:
    body = get_column_letter(idx) if is_col else str(idx)
    return ("$" + body) if absolute else body


def _transpose_span(t1: str, t2: str, edit: RefEdit, prefix: str,
                    whole: str, *, is_col: bool) -> str:
    """Shift a whole-column or whole-row span for a structural edit on its own
    axis. An edit on the other axis, or a MOVE, leaves it untouched (Excel's
    behavior: inserting rows never changes B:B)."""
    if edit.kind == MOVE:
        return whole
    axis_matches = (edit.kind in _COL_KINDS) if is_col \
        else (edit.kind in _ROW_KINDS)
    if not axis_matches:
        return whole
    p1 = _parse_span_tok(t1, is_col=is_col)
    p2 = _parse_span_tok(t2, is_col=is_col)
    if p1 is None or p2 is None:
        return whole
    (abs1, i1), (abs2, i2) = p1, p2
    span = _shift_span(min(i1, i2), max(i1, i2), edit)
    if span is None:
        return prefix + REF_ERROR
    new_lo, new_hi = span
    limit = MAX_COL if is_col else MAX_ROW
    if new_lo > limit:
        return prefix + REF_ERROR
    new_hi = min(new_hi, limit)
    if i1 <= i2:
        n1, n2 = new_lo, new_hi
    else:
        n1, n2 = new_hi, new_lo
    return (prefix + _render_span_tok(abs1, n1, is_col=is_col)
            + ":" + _render_span_tok(abs2, n2, is_col=is_col))


def _name_axis_abs(ep: _Endpoint, edit: RefEdit) -> bool:
    """Whether a name's endpoint is a fixed position on the edit's axis."""
    return ep.row_abs if edit.kind in _ROW_KINDS else ep.col_abs


def _transform_name_endpoints(
    e1: _Endpoint, e2: _Endpoint | None, edit: RefEdit
) -> tuple[_Endpoint, _Endpoint | None] | None:
    """A defined name's reference under a structural edit.

    Excel stores a relative reference in a defined name relative to cell A1
    and resolves it from the cell that uses the name ("the cell above" is
    Data!A1048576), so it is an offset, not a position: an insert or a
    delete anywhere never changes it (measured: Excel's own row and column
    inserts and deletes leave Data!A1048576 and Data!XFD1 as they are). Only
    the absolute ($) parts name positions on the sheet, and only they move.
    Shifting the relative parts as positions wrote Data!A1048577, a row
    past the grid, and turned "the cell above" into "two cells above"."""
    if edit.kind == MOVE:
        if e1.row_abs and e1.col_abs and (
                e2 is None or (e2.row_abs and e2.col_abs)):
            return _transform_endpoints(e1, e2, edit)
        return e1, e2
    a1 = _name_axis_abs(e1, edit)
    if e2 is None:
        return _transform_endpoints(e1, None, edit) if a1 else (e1, None)
    a2 = _name_axis_abs(e2, edit)
    if a1 and a2:
        return _transform_endpoints(e1, e2, edit)
    if not a1 and not a2:
        return e1, e2
    # one fixed end and one offset end: the fixed end moves on its own
    out = []
    for ep, fixed in ((e1, a1), (e2, a2)):
        if fixed:
            moved = _transform_endpoints(ep, None, edit)
            if moved is None:
                return None
            ep = moved[0]
        out.append(ep)
    return out[0], out[1]


def _transpose_match(m: re.Match, edit: RefEdit,
                     formula_sheet: str | None, in_name: bool = False) -> str:
    whole = m.group(0)
    if getattr(m, "before", "") in _NOT_OURS:
        # Another workbook's reference ([1]Data!A1) or an error (#REF!A1):
        # an edit in this workbook never moves it (Excel leaves it).
        return whole
    prefix = m.group("sheet")
    c1 = m.group("c1")
    c2 = m.group("c2")
    sheet_name, is_3d = _sheet_scope(prefix)

    if is_3d:
        # A 3D reference (North:West!F4) is left as it is.
        # Excel's own result for rows inserted or deleted on ONE sheet of the
        # span, measured with a private hidden Excel: =SUM(North:West!F4)
        # stays F4 while =North!F4
        # becomes F5. The span names every sheet in it; shifting it for one
        # sheet would be wrong for the others.
        return whole
    else:
        if not _in_scope(sheet_name, formula_sheet, edit.sheet):
            return whole

    pre = prefix or ""
    if m.group("cs1"):
        t1, t2 = m.group("cs1"), m.group("cs2")
        if in_name and not (t1.startswith("$") and t2.startswith("$")):
            return whole  # a relative span in a name is an offset
        return _transpose_span(t1, t2, edit, pre, whole, is_col=True)
    if m.group("rs1"):
        t1, t2 = m.group("rs1"), m.group("rs2")
        if in_name and not (t1.startswith("$") and t2.startswith("$")):
            return whole
        return _transpose_span(t1, t2, edit, pre, whole, is_col=False)

    ep1 = _parse_cell(c1)
    ep2 = _parse_cell(c2) if c2 else None
    if ep1 is None or (c2 and ep2 is None):
        return whole  # not a coordinate pair we transform; leave alone

    if in_name:
        result = _transform_name_endpoints(ep1, ep2, edit)
    else:
        result = _transform_endpoints(ep1, ep2, edit)
    pre = prefix or ""
    if result is None:
        return pre + REF_ERROR
    new1, new2 = result
    if new2 is None:
        return pre + _render_cell(new1)
    return pre + _render_cell(new1) + ":" + _render_cell(new2)


def transpose_formula(formula: str, edit: RefEdit,
                      formula_sheet: str | None = None, *,
                      in_name: bool = False) -> str:
    """Rewrite the A1 references in a formula string for a structural edit.

    formula: the formula, with or without a leading '='.
    edit: the structural change (RefEdit).
    formula_sheet: the sheet the formula lives on, used to resolve unqualified
        references. When None, unqualified refs are treated as in scope.
    in_name: the text is a defined name's definition, whose relative
        references are offsets from the cell that uses the name: only its
        absolute ($) parts move (_transform_name_endpoints).

    Returns the rewritten formula (leading '=' preserved). References outside
    the edited sheet are untouched; a wholly-deleted target becomes #REF!.
    String literals, structured references (Sales[Q1] names a column, not
    cell Q1) and references to other workbooks ([1]Data!A1) are never
    modified, as Excel leaves them.
    """
    if not formula:
        return formula
    lead = ""
    body = formula
    if body.startswith("="):
        lead, body = "=", body[1:]
    return lead + _Read(body).sub(
        _REF_RE, "sheet",
        lambda m: _transpose_match(m, edit, formula_sheet, in_name))


def _offset_span(t1: str, t2: str, prefix: str, whole: str, delta: int,
                 *, is_col: bool) -> str:
    """Copy/paste shift for a whole-column or whole-row span: relative
    endpoints move by the offset on their own axis, absolute ($) endpoints
    stay put, off-grid results become #REF! (Excel copy semantics)."""
    if delta == 0:
        return whole
    p1 = _parse_span_tok(t1, is_col=is_col)
    p2 = _parse_span_tok(t2, is_col=is_col)
    if p1 is None or p2 is None:
        return whole
    limit = MAX_COL if is_col else MAX_ROW
    out = []
    for absolute, idx in (p1, p2):
        new = idx if absolute else idx + delta
        if new < 1 or new > limit:
            return prefix + REF_ERROR
        out.append(_render_span_tok(absolute, new, is_col=is_col))
    return prefix + out[0] + ":" + out[1]


def _offset_match(m: re.Match, dr: int, dc: int) -> str:
    if getattr(m, "before", "") == "#":
        return m.group(0)  # an error such as #REF!A1 is not a reference
    prefix = m.group("sheet") or ""
    c1 = m.group("c1")
    c2 = m.group("c2")
    if m.group("cs1"):
        return _offset_span(m.group("cs1"), m.group("cs2"), prefix,
                            m.group(0), dc, is_col=True)
    if m.group("rs1"):
        return _offset_span(m.group("rs1"), m.group("rs2"), prefix,
                            m.group(0), dr, is_col=False)

    def shift(tok: str) -> str | None:
        ep = _parse_cell(tok)
        if ep is None:
            return tok
        row = ep.row if ep.row_abs else ep.row + dr
        col = ep.col if ep.col_abs else ep.col + dc
        if row < 1 or col < 1 or row > MAX_ROW or col > MAX_COL:
            return None
        return _render_cell(_Endpoint(ep.col_abs, col, ep.row_abs, row))

    s1 = shift(c1)
    if s1 is None:
        return prefix + REF_ERROR
    if c2 is None:
        return prefix + s1
    s2 = shift(c2)
    if s2 is None:
        return prefix + REF_ERROR
    return prefix + s1 + ":" + s2


def offset_formula(formula: str, dr: int, dc: int,
                   formula_sheet: str | None = None) -> str:
    """Shift the RELATIVE (non-$) references in a formula by (dr, dc), Excel's
    copy/paste semantics: a copied formula's relative refs move with the paste
    offset, absolute ($) anchors stay put, and a reference that would move off
    the grid becomes #REF!. Sheet-qualified refs keep their qualifier; string
    literals and structured references are never touched; a relative
    reference to another workbook moves like any other ([1]Data!A1 copied a
    row down is [1]Data!A2, as in Excel). Used by copy_range when replicating
    formulas."""
    if not formula or (dr == 0 and dc == 0):
        return formula
    lead = ""
    body = formula
    if body.startswith("="):
        lead, body = "=", body[1:]
    return lead + _Read(body, other_books_move=True).sub(
        _REF_RE, "sheet", lambda m: _offset_match(m, dr, dc))


def transpose_ref(ref: str, edit: RefEdit,
                  ref_sheet: str | None = None) -> str | None:
    """Rewrite a bare reference string (no leading '='), e.g. a defined-name
    target, a conditional-format sqref member, or a table ref: 'A1:C10',
    '$B$2', 'Sheet1!$A$1:$D$20'. Returns the new ref, or None when the whole
    reference is deleted (the caller decides how to render #REF! for its part).

    ref_sheet is the sheet the ref belongs to when the ref itself carries no
    qualifier (a per-sheet CF/DV range). Unlike a formula, a bare ref with no
    qualifier and no ref_sheet is treated as in scope."""
    if not ref:
        return ref
    m = _REF_RE.fullmatch(ref.strip())
    if m is None:
        # Not a single cell/range token (multi-area, whole col/row, name):
        # run the general formula-style pass so multi-token strings still get
        # their in-scope refs shifted; return as-is if nothing matched.
        return transpose_formula(ref, edit, ref_sheet)
    new = _transpose_match(m, edit, ref_sheet)
    # A wholly-deleted reference renders as "#REF!" or "Sheet!#REF!"; signal
    # the collapse to the caller with None so it can drop the whole ref.
    if new == REF_ERROR or new.endswith("!" + REF_ERROR):
        return None
    return new


# ------------------------------------------------------ workbook rewriter


@dataclass
class RewriteReport:
    """What the workbook rewrite touched, for the response envelope and tests."""

    formulas: int = 0
    ref_errors: int = 0
    names: int = 0
    conditional_formats: int = 0
    data_validations: int = 0
    tables: int = 0
    merges: int = 0
    dropped_merges: int = 0
    details: list[str] = field(default_factory=list)
    # The sites Excel also rewrites.
    rule_formulas: int = 0
    chart_refs: int = 0
    print_settings: int = 0
    pivot_sources: int = 0
    # What sits at a position on the edited sheet and moves with its cells.
    positions: int = 0
    #: Places that cover exactly a growing table's data but are not extended
    #: (the answer names each as a warning; nothing about them is silent).
    not_extended: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        out = {
            "formulas": self.formulas,
            "ref_errors": self.ref_errors,
            "names": self.names,
            "conditional_formats": self.conditional_formats,
            "data_validations": self.data_validations,
            "tables": self.tables,
            "merges": self.merges,
            "dropped_merges": self.dropped_merges,
            "details": self.details,
        }
        for key in ("rule_formulas", "chart_refs", "print_settings", "pivot_sources",
                    "positions"):
            if getattr(self, key):
                out[key] = getattr(self, key)
        return out


# ------------------------------------------------------------ text sites
# Every place a reference is written as TEXT that Excel rewrites with the
# workbook: formula cells (and array formulas), defined names, conditional
# format rule formulas, data-validation formulas and chart references. The
# structural edit, the sheet rename, the sheet delete and the table rename
# all walk the same sites (one rewriter).


def _chart_ref_objects(chart):
    """Every reference object (NumRef, StrRef, MultiLevelStrRef: an `f`
    string) inside one openpyxl chart model: series values, categories,
    names, x/y values, bubble sizes and titles. Yields (ref, path), path
    being the (owner, attribute) pairs from the chart down to it."""
    seen: set[int] = set()
    stack = [(chart, ())]
    while stack:
        obj, path = stack.pop()
        if id(obj) in seen:
            continue
        seen.add(id(obj))
        if type(obj).__name__ in ("NumRef", "StrRef", "MultiLevelStrRef") \
                and isinstance(getattr(obj, "f", None), str):
            yield obj, path
            continue
        items = vars(obj).items() if hasattr(obj, "__dict__") else ()
        for name, v in items:
            if isinstance(v, (list, tuple)):
                stack.extend((x, path + ((obj, name),)) for x in v if hasattr(x, "__dict__"))
            elif hasattr(v, "__dict__") and type(v).__module__.startswith("openpyxl.chart"):
                stack.append((v, path + ((obj, name),)))


def _all_charts(wb):
    for ws in list(wb.worksheets) + list(getattr(wb, "chartsheets", [])):
        for chart in getattr(ws, "_charts", []):
            yield ws, chart


def _rule_and_chart_sites(wb, fn, report: "RewriteReport", on_site=None) -> None:
    """fn(text, home_sheet) applied to conditional-format rule formulas,
    data-validation formulas and chart references. on_site(kind, where), when
    given, is called just before fn at each site: kind is "conditional
    format", "data validation" or "chart", where the cells the rule covers
    (or the chart's reference)."""
    note = on_site or (lambda kind, where: None)
    for ws in wb.worksheets:
        home = ws.title
        rules_map = getattr(getattr(ws, "conditional_formatting", None), "_cf_rules", None)
        for cf, rules in (rules_map or {}).items():
            for rule in rules:
                formulas = list(getattr(rule, "formula", None) or [])
                if not formulas:
                    continue
                note("conditional format", f"{qualifier([home])}{_sqref_text(getattr(cf, 'sqref', None))}")
                new = [fn(f, home) if isinstance(f, str) else f for f in formulas]
                if new != formulas:
                    rule.formula = new
                    report.rule_formulas += 1
        dvs = getattr(ws, "data_validations", None)
        for dv in list(getattr(dvs, "dataValidation", None) or []):
            for attr in ("formula1", "formula2"):
                val = getattr(dv, attr, None)
                if isinstance(val, str) and val:
                    note("data validation", f"{qualifier([home])}{_sqref_text(getattr(dv, 'sqref', None))}")
                    new = fn(val, home)
                    if new != val:
                        setattr(dv, attr, new)
                        report.rule_formulas += 1
    for ws, chart in _all_charts(wb):
        for ref, path in list(_chart_ref_objects(chart)):
            note("chart", ref.f)
            new = fn(ref.f, ws.title)
            if new == ref.f:
                continue
            report.chart_refs += 1
            collapsed = REF_ERROR in new and REF_ERROR not in ref.f
            owner = next(((o, a) for o, a in reversed(path) if a == "cat"), None)
            if collapsed and owner is not None:
                # Excel drops a series' category reference whose cells were
                # deleted (=SERIES(Budget!$C$1,,Budget!$C$2:$C$4,1)).
                setattr(owner[0], "cat", None)
            else:
                ref.f = new


def _sqref_text(sqref) -> str:
    """The cells a rule covers as text (B2:B4, or several areas separated
    by commas, as Excel writes them), '?' when the rule does not say."""
    text = str(sqref).strip() if sqref is not None else ""
    return ",".join(text.split()) or "?"


def rewrite_text_sites(wb, fn, report: "RewriteReport") -> None:
    """Apply fn(text, home_sheet) -> new text at every text site above,
    counting what changed. fn is given whole formula text and must leave
    string literals alone (every rewriter in this module does)."""
    from openpyxl.worksheet.formula import ArrayFormula

    for ws in wb.worksheets:
        home = ws.title
        for cell in list(getattr(ws, "_cells", {}).values()):
            val = cell.value
            if isinstance(val, str) and _calc.is_formula_cell(cell):
                new = fn(val, home)
                if new != val:
                    cell.value = new
                    report.formulas += 1
                    if REF_ERROR in new and REF_ERROR not in val:
                        report.ref_errors += 1
            elif isinstance(val, ArrayFormula) and val.text:
                new = fn(val.text, home)
                if new != val.text:
                    cell.value = ArrayFormula(val.ref, new)
                    report.formulas += 1
                    if REF_ERROR in new and REF_ERROR not in val.text:
                        report.ref_errors += 1
    scopes = [(None, wb.defined_names)]
    for ws in wb.worksheets:
        dn = getattr(ws, "defined_names", None)
        if dn is not None:
            scopes.append((ws.title, dn))
    for home, dndict in scopes:
        for _name, defn in list(dndict.items()):
            val = getattr(defn, "value", None) or getattr(defn, "attr_text", None)
            if not isinstance(val, str) or not val:
                continue
            new = fn(val, home)
            if new != val:
                try:
                    defn.value = new
                except Exception:
                    defn.attr_text = new
                report.names += 1
    _rule_and_chart_sites(wb, fn, report)
    # A table's own copies of its calculated-column and totals formulas, so a
    # sheet rename, a sheet delete or a table rename reaches the formula the
    # table gives its next row.
    _rewrite_table_formula_text(wb, None, report, fn=fn)


# ------------------------------------------------------- sheet qualifiers

# A sheet qualifier before '!': 'Quoted Name'!, Bare!, or a 3D span
# (Bare:Bare! or 'A:B C'!), the bare form in any script (Données!, 売上!).
# Not preceded by a name character, a quote or a !. One after ']' (another
# workbook's, [1]Jan:Mar!A1) or '#' (#REF! is an error, not sheet REF)
# matches whole and is left alone, so its second half is never read alone.
_QUAL_RE = re.compile(
    "(?<![^" + _BEFORE_OK + "])"
    "(?P<q>" + _QUOTED_SHEET + "|" + _UNQUOTED_SHEET + "(?::" + _UNQUOTED_SHEET + ")?)!")

# A name Excel could read as a cell (A1, SUM2, XFD1) or as an R1C1 reference
# (R, C1, RC, R1C, and R1C1x, which Excel quotes too).
_LOOKS_LIKE_A1 = re.compile(r"[A-Za-z]{1,3}[0-9]+")
_LOOKS_LIKE_R1C1 = re.compile(r"[Rr][0-9]*[Cc]?[0-9]*|[Cc][0-9]*|[Rr][0-9]+[Cc][0-9]+.*",
                              re.DOTALL)


def _needs_quotes(name: str) -> bool:
    """Whether a sheet name is quoted in formula text. Excel leaves a name
    bare when it is made of letters, marks and digits in any script, with _
    and periods, and starts with a letter or _ (Données!A1, 売上!A1, Q_1!A1);
    it quotes one that holds a space or other punctuation, starts with a
    digit, a mark or a period, could be read as a reference, or is TRUE or
    FALSE. Excel reads a quoted name wherever it reads a bare one, so a name
    this cannot place is quoted."""
    if not name:
        return True
    first = unicodedata.category(name[0])
    if not (name[0] == "_" or first.startswith("L")):
        return True
    for ch in name:
        if ch not in "_." and unicodedata.category(ch)[0] not in "LMN":
            return True
    return (bool(_LOOKS_LIKE_A1.fullmatch(name)) or bool(_LOOKS_LIKE_R1C1.fullmatch(name))
            or name.upper() in ("TRUE", "FALSE"))


def qualifier(names: list[str]) -> str:
    """'Sheet'! text for one sheet or a 3D span, quoted the way Excel quotes
    it: the whole span inside one pair of quotes when any name needs them
    (Excel wrote =SUM('North:West Coast'!F4))."""
    if any(_needs_quotes(n) for n in names):
        return "'" + ":".join(n.replace("'", "''") for n in names) + "'!"
    return ":".join(names) + "!"


def _qual_names(q: str) -> list[str] | None:
    if q.startswith("'"):
        body = q[1:-1].replace("''", "'")
        if "[" in body:
            return None  # another workbook ('[1]Data'!A1); no sheet name holds a bracket
        return body.split(":")
    return q.split(":")


def _rewrite_qualifiers(text: str, fn) -> str:
    """fn(hit) for every sheet qualifier in formula text, read by the lexer:
    never inside a string, a quoted name's own text, a structured
    reference's brackets or another workbook's reference."""
    return _Read(text).sub(_QUAL_RE, "q",
                           lambda m: m.group(0) if m.before in _NOT_OURS else fn(m))


def rename_sheet_in_formula(text: str, old: str, new: str) -> str:
    """Every qualifier naming sheet `old` (alone or as a 3D endpoint) now
    names `new`, quoted where Excel quotes it."""
    low = old.lower()

    def sub(m) -> str:
        names = _qual_names(m.group("q"))
        if not names or not any(n.lower() == low for n in names):
            return m.group(0)
        return qualifier([new if n.lower() == low else n for n in names])

    return _rewrite_qualifiers(text, sub)


# A qualifier with what it qualifies: a cell, a range, a whole-column or
# whole-row span, or a sheet-level name (North!LocalRate).
_REF_TOKEN = (r"\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?"
              r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[0-9]+:\$?[0-9]+")
_QUALIFIED_RE = re.compile(
    _QUAL_RE.pattern + "(?P<tok>(?:" + _REF_TOKEN + ")(?![^" + _AFTER_OK + "])|" + _NAME_CH + "*)")


def delete_sheet_in_formula(text: str, deleted: str, order: list[str]) -> str:
    """Excel's result when sheet `deleted` goes (order: the sheet names
    before the delete), as Excel STORES it: a reference to it alone becomes
    #REF!, cell part and all (=North!F5 -> =#REF!, =SUM(North!F2:F5) ->
    =SUM(#REF!)); a 3D span whose END it was shrinks to the next sheet
    inside the span, and to a one-sheet reference when one sheet is left; a
    span it sat inside is unchanged. (Excel shows =#REF!F5 while the
    workbook is open, but saves #REF!, and it opens a file holding #REF!F5
    only after a repair.)"""
    low = deleted.lower()
    pos = {n.lower(): i for i, n in enumerate(order)}

    def sub(m) -> str:
        if m.before in _NOT_OURS:
            return m.group(0)
        names = _qual_names(m.group("q"))
        if not names or not any(n.lower() == low for n in names):
            return m.group(0)
        if len(names) == 1:
            return REF_ERROR
        a, b = (pos.get(n.lower()) for n in names)
        if a is None or b is None:
            return m.group(0)
        if names[0].lower() != low and names[-1].lower() != low:
            return m.group(0)
        span = [n for n in order[min(a, b):max(a, b) + 1] if n.lower() != low]
        if not span:
            return REF_ERROR
        return qualifier([span[0]] if len(span) == 1 else [span[0], span[-1]]) + m.group("tok")

    return _Read(text).sub(_QUALIFIED_RE, "q", sub)


def _pivot_caches(wb):
    seen: set[int] = set()
    for ws in wb.worksheets:
        for pt in getattr(ws, "_pivots", []) or []:
            cache = getattr(pt, "cache", None)
            if cache is not None and id(cache) not in seen:
                seen.add(id(cache))
                yield cache


def _worksheet_source(cache):
    return getattr(getattr(cache, "cacheSource", None), "worksheetSource", None)


def rename_sheet_refs(wb, old: str, new: str) -> "RewriteReport":
    """A sheet was renamed from old to new (the caller already set the
    title): every reference to it, everywhere Excel rewrites one, follows.
    Pivot cache sources name their sheet in an attribute."""
    report = RewriteReport()
    rewrite_text_sites(wb, lambda t, _h: rename_sheet_in_formula(t, old, new), report)
    for cache in _pivot_caches(wb):
        src = _worksheet_source(cache)
        if src is not None and isinstance(src.sheet, str) and src.sheet.lower() == old.lower():
            src.sheet = new
            report.pivot_sources += 1
    return report


def delete_sheet_refs(wb, deleted: str, order: list[str]) -> "RewriteReport":
    """Sheet `deleted` was removed; order is the sheet names before it was.
    References to it take Excel's result (delete_sheet_in_formula)."""
    report = RewriteReport()
    rewrite_text_sites(wb, lambda t, _h: delete_sheet_in_formula(t, deleted, order), report)
    return report


def rename_table_refs(wb, old: str, new: str) -> "RewriteReport":
    """A table was renamed: every structured reference that names it
    follows (Excel: =SUM(Stock_Tbl[Value]) -> =SUM(Inventory[Value]),
    and the defined name =Stock_Tbl[Value] too; text inside a string, such
    as INDIRECT("Stock_Tbl[Value]"), is left as it is, as Excel leaves it)."""
    report = RewriteReport()
    rewrite_text_sites(wb, lambda t, _h: rename_table_in_formula(t, old, new),
                       report)
    return report


def rename_table_in_formula(text: str, old: str, new: str) -> str:
    """Every structured reference naming table `old` now names `new`."""
    # The table's name wherever it stands as a name: before its brackets
    # (Stock_Tbl[Value]) or alone (=ROWS(Stock_Tbl)); never a longer name in
    # any script, a function (Name(...)), a sheet qualifier (Name!A1), or text
    # inside a string literal, a quoted sheet name or a structured
    # reference's brackets (the lexer blanks those).
    if old.lower() not in text.lower():
        return text
    name_re = re.compile("(?<![^" + _BEFORE_OK + "])" + re.escape(old)
                         + "(?![^" + _AFTER_TABLE_OK + "])", re.IGNORECASE)
    return _Read(text).sub(name_re, None, lambda _m: new)


def _rewrite_sqref_member(member: str, edit: RefEdit,
                          sheet_name: str) -> str | None:
    return transpose_ref(member, edit, sheet_name)


def rewrite_workbook(wb, edit: RefEdit) -> RewriteReport:
    """Apply one structural edit's reference rewrites across an open openpyxl
    workbook. This does NOT move cell values (openpyxl's insert/delete does
    that, or the caller); it repairs every REFERENCE so the workbook stays
    coherent: formula cells on every sheet, workbook and sheet defined names,
    per-sheet conditional-format and data-validation ranges, table refs, and
    merged ranges on the edited sheet; and conditional-format
    rule formulas, data-validation formulas, chart references, the print
    area and print titles, the autofilter and pivot cache sources.
    """
    report = RewriteReport()
    # The positions, tables, merges and the print and filter settings follow
    # the cells; only the places that READ the table's cells (formulas, names,
    # rule ranges, chart series) are extended with a growing table.
    plain = _replace(edit, grow=()) if edit.grow else edit
    _rewrite_sites_beyond_cells_and_names(wb, edit, report)
    _rewrite_table_formula_text(wb, edit, report)
    if edit.kind != MOVE:
        _rewrite_positions(wb, plain, report)

    # 1. Formula cells across every sheet (cross-sheet refs to the edited
    #    sheet must be repaired too). Array (CSE / dynamic) formulas are stored
    #    as ArrayFormula objects, not strings; their text and anchor ref are
    #    both rewritten (they were once silently skipped).
    from openpyxl.worksheet.formula import ArrayFormula
    for ws in wb.worksheets:
        home = ws.title
        cells = getattr(ws, "_cells", None)
        iterator = cells.values() if cells is not None else (
            c for row in ws.iter_rows() for c in row)
        for cell in list(iterator):
            val = cell.value
            # data_type, not the leading '=': a TEXT cell holding "=A1+1"
            # (import_data's neutralized injection text) is DATA. Rewriting
            # its references would edit the user's text, and the write-back
            # would re-type it as a live formula.
            if isinstance(val, str) and _calc.is_formula_cell(cell):
                new = transpose_formula(val, edit, home)
                if new != val:
                    cell.value = new
                    report.formulas += 1
                    if REF_ERROR in new and REF_ERROR not in val:
                        report.ref_errors += 1
            elif isinstance(val, ArrayFormula):
                text = val.text or ""
                new_text = transpose_formula(text, edit, home) if text else text
                ref = val.ref
                new_ref = ref
                if isinstance(ref, str) and ref:
                    shifted = transpose_ref(ref, edit, home)
                    if shifted is not None:
                        new_ref = shifted
                    # a wholly-deleted anchor keeps its old ref (conservative;
                    # openpyxl's own delete already removed the cells)
                if new_text != text or new_ref != ref:
                    cell.value = ArrayFormula(new_ref, new_text)
                    report.formulas += 1
                    if REF_ERROR in new_text and REF_ERROR not in text:
                        report.ref_errors += 1

    # 2. Defined names (workbook + sheet scope).
    scopes = [("workbook", wb.defined_names)]
    for ws in wb.worksheets:
        dn = getattr(ws, "defined_names", None)
        if dn is not None:
            scopes.append((ws.title, dn))
    for _scope, dndict in scopes:
        for name, defn in list(dndict.items()):
            val = getattr(defn, "value", None) or getattr(defn, "attr_text", None)
            if not isinstance(val, str) or not val:
                continue
            # A name's relative references are offsets from the cell that
            # uses it, so only its absolute parts move (as in Excel).
            new = transpose_formula(val, edit, None, in_name=True)
            if new != val:
                try:
                    defn.value = new
                except Exception:
                    defn.attr_text = new
                report.names += 1

    # 3. Conditional formatting + 4. Data validation + 5. Tables + 6. Merges,
    #    per sheet.
    for ws in wb.worksheets:
        sheet_name = ws.title

        # conditional formatting. A ConditionalFormatting object hashes on its
        # sqref, so mutating sqref in place corrupts the internal _cf_rules
        # dict; rebuild the dict instead.
        cf_container = getattr(ws, "conditional_formatting", None)
        rules_map = getattr(cf_container, "_cf_rules", None)
        if isinstance(rules_map, dict) and rules_map:
            from openpyxl.worksheet.cell_range import MultiCellRange
            old_items = list(rules_map.items())
            rules_map.clear()
            for cf_obj, rules in old_items:
                new_members = []
                changed = False
                for member in [str(x) for x in cf_obj.sqref.ranges]:
                    nm = _rewrite_sqref_member(member, edit, sheet_name)
                    if nm is None:
                        changed = True
                        continue
                    if nm != member:
                        changed = True
                    new_members.append(nm)
                if not new_members:
                    report.conditional_formats += 1  # whole range deleted
                    continue
                cf_obj.sqref = MultiCellRange(" ".join(new_members))
                rules_map[cf_obj] = rules
                if changed:
                    report.conditional_formats += 1

        # data validation
        dvs = getattr(ws, "data_validations", None)
        if dvs is not None:
            for dv in list(dvs.dataValidation):
                new_members = []
                changed = False
                for member in [str(x) for x in dv.sqref.ranges]:
                    nm = _rewrite_sqref_member(member, edit, sheet_name)
                    if nm is None:
                        changed = True
                        continue
                    if nm != member:
                        changed = True
                    new_members.append(nm)
                if changed:
                    from openpyxl.worksheet.cell_range import MultiCellRange
                    dv.sqref = MultiCellRange(" ".join(new_members))
                    report.data_validations += 1

        # tables (ws.tables is a name->ref mapping; the Table object is
        # fetched by name)
        for tname in list(getattr(ws, "tables", {})):
            table = ws.tables[tname]
            ref = table.ref
            nm = transpose_ref(ref, plain, sheet_name)
            if nm is None:
                report.details.append(
                    f"table {tname!r} on {sheet_name!r} lost its range to a "
                    "delete; left unchanged for the caller to resolve")
                continue
            if nm != ref:
                table.ref = nm
                report.tables += 1

        # merged cells
        merged = getattr(ws, "merged_cells", None)
        if merged is not None:
            existing = [str(r) for r in list(merged.ranges)]
            for member in existing:
                nm = transpose_ref(member, plain, sheet_name)
                if nm == member:
                    continue
                # The cells have already moved. unmerge_cells would delete
                # every cell of the OLD range but the first, and what sits
                # there now is the moved top-left cell (or the rows that
                # moved up): take the range out of the list and leave the
                # cells alone.
                try:
                    from openpyxl.worksheet.cell_range import CellRange
                    merged.remove(CellRange(member))
                except Exception:
                    pass
                if nm is None:
                    report.dropped_merges += 1
                    continue
                try:
                    ws.merge_cells(nm)
                    report.merges += 1
                except Exception:
                    report.dropped_merges += 1

    return report


def deleted_column_refs(wb, table: str, column: str) -> "RewriteReport":
    """After a table column is deleted, every structured reference that named
    it becomes #REF!, as Excel stores it (measured: SUM(Sales[FY24]) and
    SUM(Sales[[FY23]:[FY24]]) are stored as SUM(#REF!), Sales[[#Headers],[FY24]]
    as #REF!, in cells and in defined names alike). Left as text, the dead
    reference would bind again to any later column of the same name, which
    Excel's #REF! never does. Only a reference that names the column is
    touched: one whose range merely spans it (Sales[[FY22]:[FY24]] after FY23
    goes) keeps its two endpoints, which still exist."""
    from . import structrefs as _sr

    low_t, low_c = table.lower(), column.strip().lower()

    def fn(text: str, _home) -> str:
        if "[" not in text:
            return text
        out, pos = [], 0
        for ref in _sr.find_refs(text):
            if ref.table is None or ref.table.lower() != low_t:
                continue
            if not any(c.strip().lower() == low_c for c in ref.columns):
                continue
            out.append(text[pos:ref.start])
            out.append(REF_ERROR)
            pos = ref.end
        if not out:
            return text
        out.append(text[pos:])
        return "".join(out)

    report = RewriteReport()
    rewrite_text_sites(wb, fn, report)
    return report


def _rewrite_table_formula_text(wb, edit: RefEdit | None, report: RewriteReport,
                                fn=None) -> None:
    """A table keeps its own copy of a calculated column's formula and of a
    custom totals formula in the table part. The cells move with the edit; the
    copies must too, or the next row Excel adds to the table takes the old
    formula (=[@Q1]*$C$4 kept $C$4 after an insert above it)."""
    for ws in wb.worksheets:
        for tname in list(getattr(ws, "tables", {})):
            table = ws.tables[tname]
            for tc in table.tableColumns or []:
                for attr in ("calculatedColumnFormula", "totalsRowFormula"):
                    tf = getattr(tc, attr, None)
                    text = getattr(tf, "attr_text", None)
                    if not isinstance(text, str) or not text:
                        continue
                    new = fn(text, ws.title) if fn is not None else \
                        transpose_formula(text, edit, ws.title)
                    if new != text:
                        tf.attr_text = new
                        report.rule_formulas += 1


def _moved_index(pos: int, edit: RefEdit) -> int | None:
    """Where row or column `pos` (1-based, on the edit's axis) lands after
    the edit; None when the edit deletes it."""
    lo, n = edit.index, edit.count
    if edit.kind in (INSERT_ROWS, INSERT_COLS):
        return pos + n if pos >= lo else pos
    if pos > lo + n - 1:
        return pos - n
    return None if pos >= lo else pos


def _shift_marker(marker, edit: RefEdit, rows: bool) -> int:
    """Move one drawing-anchor corner (0-based row or column plus an offset
    inside it) with its cell. A corner in a deleted row or column goes to the
    start of the one that takes its place, as Excel draws it. Returns how far
    the corner moved."""
    attr, off, limit = ("row", "rowOff", MAX_ROW) if rows else ("col", "colOff", MAX_COL)
    old = getattr(marker, attr)
    new = _moved_index(old + 1, edit)
    if new is None:
        new = edit.index
        setattr(marker, off, 0)
    new = min(new, limit) - 1
    setattr(marker, attr, new)
    return new - old


def _rewrite_positions(wb, edit: RefEdit, report: RewriteReport) -> None:
    """What sits AT a position on the edited sheet and moves with its cells
    when rows or columns are inserted or deleted (Excel moves each of them;
    openpyxl moves only the cells): what-if data tables (their range and
    input cells), pivot tables, charts and pictures (their anchors, by their
    move-and-size setting), manual page breaks and scenario input cells.
    The edits that would cut a pivot table or a data table are refused
    before this runs (split_conflict), as Excel refuses them."""
    from openpyxl.worksheet.formula import DataTableFormula

    edited = _norm_sheet(edit.sheet)
    rows = edit.kind in (INSERT_ROWS, DELETE_ROWS)
    for ws in wb.worksheets:
        if _norm_sheet(ws.title) != edited:
            continue
        home = ws.title
        # Data tables: the top-left cell holds the table's formula, with its
        # range (ref) and its input cells (r1, r2) as cell addresses.
        for cell in list(getattr(ws, "_cells", {}).values()):
            dt = cell._value
            if not isinstance(dt, DataTableFormula):
                continue
            changed = False
            new = transpose_ref(dt.ref, edit, home) if dt.ref else dt.ref
            if new is not None and new != dt.ref:
                dt.ref, changed = new, True
            for key, gone in (("r1", "del1"), ("r2", "del2")):
                cur = getattr(dt, key)
                if not cur:
                    continue
                new = transpose_ref(cur, edit, home)
                if new is None:
                    # The input cell was deleted: the part says so (del1/del2)
                    # and keeps the address, as the file format records it.
                    setattr(dt, gone, True)
                    changed = True
                elif new != cur:
                    setattr(dt, key, new)
                    changed = True
            report.positions += changed
        # Pivot tables: where the table stands on the sheet.
        for pt in getattr(ws, "_pivots", []) or []:
            loc = getattr(pt, "location", None)
            ref = getattr(loc, "ref", None)
            if not isinstance(ref, str) or not ref:
                continue
            new = transpose_ref(ref, edit, home)
            if new is None:
                report.details.append(
                    f"pivot table {getattr(pt, 'name', '?')!r} on {home!r} stood in the "
                    "deleted cells; refresh or rebuild it in Excel")
            elif new != ref:
                loc.ref = new.replace("$", "")
                report.positions += 1
        # Charts and pictures: their anchor corners follow the cells under
        # them. Move-and-size (twoCell, the default) moves each corner with
        # its own cell; move-only (oneCell) moves the whole object with its
        # top-left corner; absolute stays where it is.
        for obj in list(getattr(ws, "_charts", []) or []) + list(getattr(ws, "_images", []) or []):
            anchor = getattr(obj, "anchor", None)
            if isinstance(anchor, str):
                new = transpose_ref(anchor, edit, home)
                if new is not None and new != anchor:
                    obj.anchor = new.replace("$", "")
                    report.positions += 1
                continue
            start = getattr(anchor, "_from", None)
            if start is None:
                continue  # an absolute anchor: placed by position, not by cell
            end = getattr(anchor, "to", None)
            how = getattr(anchor, "editAs", None) or ("twoCell" if end is not None else "oneCell")
            if how == "absolute":
                continue
            moved = _shift_marker(start, edit, rows)
            if end is not None:
                if how == "oneCell":
                    attr = "row" if rows else "col"
                    limit = (MAX_ROW if rows else MAX_COL) - 1
                    setattr(end, attr, min(getattr(end, attr) + moved, limit))
                else:
                    moved = _shift_marker(end, edit, rows) or moved
            report.positions += bool(moved)
        # Manual page breaks: a break belongs to the row (or column) that
        # starts the page, and goes with it; deleting that row deletes it.
        brks = getattr(ws, "row_breaks" if rows else "col_breaks", None)
        if brks is not None and getattr(brks, "brk", None):
            kept = []
            for b in list(brks.brk):
                new = _moved_index(b.id + 1, edit)
                if new is None:
                    report.positions += 1
                    continue
                if new - 1 != b.id:
                    b.id = new - 1
                    report.positions += 1
                kept.append(b)
            brks.brk = kept
        # What-if scenarios: the cells each scenario changes.
        scen = getattr(ws, "scenarios", None)
        for sc in getattr(scen, "scenario", None) or []:
            for ic in getattr(sc, "inputCells", None) or []:
                if not ic.r:
                    continue
                new = transpose_ref(ic.r, edit, home)
                if new is None:
                    ic.deleted = True
                    report.positions += 1
                elif new != ic.r:
                    ic.r = new.replace("$", "")
                    report.positions += 1


def split_conflict(ws, kind: str, index: int, count: int) -> str | None:
    """Excel refuses an insert or delete that would cut a pivot table or a
    what-if data table in two (measured: Rows(n).Insert() through either
    raises). The same edit here would leave a pivot or a data table that no
    longer matches its cells, so it is refused with what stands in the way.
    None when nothing is cut."""
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.worksheet.formula import DataTableFormula

    rows = kind in (INSERT_ROWS, DELETE_ROWS)
    insert = kind in (INSERT_ROWS, INSERT_COLS)
    lo, hi = index, index + count - 1

    def cut(ref: str, frame: int = 0) -> bool:
        try:
            c0, r0, c1, r1 = range_boundaries(ref.replace("$", ""))
        except Exception:
            return False
        a, b = (r0, r1) if rows else (c0, c1)
        if a is not None:
            a = max(1, a - frame)
        if a is None or b is None:
            return False
        if insert:
            return a < lo <= b
        # a delete that takes the whole object, or misses it, is not a cut
        return not (hi < a or lo > b or (lo <= a and b <= hi))

    for pt in getattr(ws, "_pivots", []) or []:
        ref = getattr(getattr(pt, "location", None), "ref", None)
        if isinstance(ref, str) and cut(ref):
            return (f"this edit would cut pivot table {getattr(pt, 'name', '?')!r} "
                    f"({ws.title}!{ref}) in two. Excel refuses the same edit (\"We can't "
                    "make this change for the selected cells because it will affect a "
                    "PivotTable\"). Insert or delete above or below it (or to its left or "
                    "right), or move the pivot table first")
    for cell in getattr(ws, "_cells", {}).values():
        dt = cell._value
        # A data table's ref holds its results only; the table itself also
        # takes the row above them and the column to their left (the input
        # values and the formulas), so an edit between those and the results
        # cuts it too.
        if isinstance(dt, DataTableFormula) and dt.ref and cut(dt.ref, frame=1):
            return (f"this edit would cut the what-if data table whose results are "
                    f"{ws.title}!{dt.ref} in two. Excel refuses the same edit (\"Can't "
                    "change part of a data table\"). Insert or delete outside it, or "
                    "remove the data table first")
    if rows:
        # Excel refuses a sheet-row insert or delete inside the rows of two
        # or more tables side by side (measured with its own tables: "Insert
        # method of Range class failed"); through one table, above or below
        # them, it takes the edit.
        through = []
        for tname in list(getattr(ws, "tables", {}) or {}):
            t = ws.tables[tname]
            ref = getattr(t, "ref", None)
            if not isinstance(ref, str):
                continue
            if not insert:
                # a delete cuts a table when it takes some of its data rows
                try:
                    c0, r0, c1, r1 = range_boundaries(ref.replace("$", ""))
                except Exception:
                    continue
                hdr = t.headerRowCount if t.headerRowCount is not None else 1
                ref = f"{_col(c0)}{r0 + hdr}:{_col(c1)}{r1}"
                if r0 + hdr > r1:
                    continue
            if cut(ref):
                through.append(f"{tname} ({ws.title}!{t.ref})")
        if len(through) > 1:
            verb = "insert" if insert else "delete"
            return (f"this edit would {verb} sheet rows inside {len(through)} tables "
                    f"side by side: {', '.join(through)}. Excel refuses the same edit "
                    "(\"Insert method of Range class failed\"). Edit one table's rows at "
                    "a time (manage_table add_row adds a row to one table), or insert "
                    "or delete above or below the tables")
    else:
        return _table_column_conflict(ws, insert, lo, hi)
    return None


def _quoted_list(names: list[str]) -> str:
    q = [f"'{n}'" for n in names]
    return q[0] if len(q) == 1 else ", ".join(q[:-1]) + " and " + q[-1]


def _table_column_conflict(ws, insert: bool, lo: int, hi: int) -> str | None:
    """A sheet-column insert or delete through a table's columns. Excel
    deletes or adds TABLE columns there (formulas that name a deleted one
    become #REF!, and a table loses itself with its last column). Done to
    the sheet alone, the table kept a list of columns that no longer matched
    its cells: its save then read a shifted header as a rename, or, with no
    formula naming the column, wrote a table Excel repairs on open. Refused
    with the table op that makes the edit; None when no table is cut."""
    from openpyxl.utils.cell import range_boundaries

    found = []
    for tname in list(getattr(ws, "tables", {}) or {}):
        t = ws.tables[tname]
        ref = getattr(t, "ref", None)
        if not isinstance(ref, str):
            continue
        try:
            c0, _r0, c1, _r1 = range_boundaries(ref.replace("$", ""))
        except Exception:
            continue
        if c0 is None or c1 is None:
            continue
        found.append((c0, c1, t))
    for c0, c1, t in sorted(found, key=lambda x: x[0]):
        names = [str(tc.name) for tc in (t.tableColumns or [])]
        where = f"table {t.displayName} ({ws.title}!{t.ref})"
        if insert:
            if not c0 < lo <= c1:
                continue
            left = names[lo - 1 - c0] if 0 <= lo - 1 - c0 < len(names) else _col(lo - 1)
            right = names[lo - c0] if 0 <= lo - c0 < len(names) else _col(lo)
            return (f"this edit would insert a sheet column inside {where}, between its columns "
                    f"'{left}' and '{right}'. This op does not add a table column, so "
                    "nothing was changed. manage_table with action='add_column' adds a "
                    "column at the table's right edge; to insert a sheet column, pick one "
                    "left or right of the table. Converting the table to a range "
                    "(manage_table action='to_range') removes the table, so do that only "
                    "if the user asks for it")
        if hi < c0 or lo > c1:
            continue
        if lo <= c0 and c1 <= hi:
            return (f"this edit would delete every column of {where}. This op does not delete a "
                    "table, so nothing was changed. To remove the table and its data, "
                    "turn it into plain cells first with manage_table and "
                    "action='to_range', then delete the columns")
        cut = [names[c - c0] if c - c0 < len(names) else _col(c)
               for c in range(max(lo, c0), min(hi, c1) + 1)]
        what = (f"column {_quoted_list(cut)}" if len(cut) == 1
                else f"columns {_quoted_list(cut)}")
        how = ("action='delete_column' and column='{}'".format(cut[0]) if len(cut) == 1 else
               "action='delete_column' once for each ("
               + ", then ".join(f"column='{n}'" for n in cut) + ")")
        return (f"this edit would delete {what} of {where}. This op does not delete a table "
                "column, so nothing was changed. To delete it as Excel does (formulas "
                f"that name it become #REF!), use manage_table with {how}; to delete "
                "only sheet columns, pick columns outside the table")
    return None


def _col(n: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(n)


def _note_not_extended(report: "RewriteReport", edit: RefEdit, ref: str,
                       sheet: str, what: str) -> None:
    """Name a place that covers exactly a growing table's data and was left at
    its size, because this edit does not extend it: `what` reads as the place."""
    if not edit.grow or not isinstance(ref, str):
        return
    grown = transpose_ref(ref, edit, sheet)
    if grown is not None and grown != transpose_ref(ref, _replace(edit, grow=()), sheet):
        report.not_extended.append(f"{what} {sheet}!{ref.replace('$', '')} covers the "
                                   "table's data and was not extended")


def _rewrite_sites_beyond_cells_and_names(wb, edit: RefEdit,
                                          report: RewriteReport) -> None:
    """The sites beyond cells and names (rule formulas, charts, print
    settings; the autofilter and pivot sources with them). Cells and
    names stay in rewrite_workbook's own passes."""
    from openpyxl.worksheet.print_settings import ColRange, RowRange

    plain = _replace(edit, grow=()) if edit.grow else edit
    _rule_and_chart_sites(wb, lambda text, home: transpose_formula(text, edit, home), report)

    edited = _norm_sheet(edit.sheet)
    for ws in wb.worksheets:
        if _norm_sheet(ws.title) != edited:
            continue
        # The print area (_xlnm.Print_Area), member by member.
        pa = getattr(ws, "_print_area", None)
        members = [str(r) for r in getattr(pa, "ranges", None) or []]
        if members:
            kept = []
            for m in members:
                nm = transpose_ref(m, plain, ws.title)
                _note_not_extended(report, edit, m, ws.title, "the print area")
                if nm is not None:
                    kept.append(nm)
            if kept != members:
                ws.print_area = kept or None
                report.print_settings += 1
        # The print titles (_xlnm.Print_Titles): whole rows and columns.
        for attr, cls, getter in (("_print_rows", RowRange, "print_title_rows"),
                                  ("_print_cols", ColRange, "print_title_cols")):
            cur = getattr(ws, getter)
            if not cur:
                continue
            nm = transpose_ref(cur, plain, ws.title)
            if nm != cur:
                setattr(ws, attr, cls(nm.replace("$", "")) if nm else None)
                report.print_settings += 1
        # The autofilter range (_xlnm._FilterDatabase).
        af = getattr(ws, "auto_filter", None)
        if af is not None and af.ref:
            nm = transpose_ref(af.ref, plain, ws.title)
            if nm != af.ref:
                af.ref = nm
                report.print_settings += 1
    # Pivot cache sources on the edited sheet.
    for cache in _pivot_caches(wb):
        src = _worksheet_source(cache)
        if src is None or not isinstance(src.ref, str) or not src.sheet \
                or _norm_sheet(src.sheet) != edited:
            continue
        nm = transpose_ref(src.ref, plain, src.sheet)
        _note_not_extended(report, edit, src.ref, src.sheet, "a pivot table's source range")
        if nm is None:
            report.details.append(
                f"a pivot table's source range {src.sheet}!{src.ref} was deleted; left "
                "as it was, so refresh it in Excel")
        elif nm != src.ref:
            src.ref = nm
            report.pivot_sources += 1


__all__ = [
    "rename_sheet_refs", "delete_sheet_refs", "rename_table_refs",
    "rename_table_in_formula", "rewrite_text_sites",
    "rename_sheet_in_formula", "delete_sheet_in_formula", "qualifier",
    "RefEdit", "RewriteReport", "rewrite_workbook",
    "transpose_formula", "transpose_ref", "offset_formula",
    "INSERT_ROWS", "DELETE_ROWS", "INSERT_COLS", "DELETE_COLS", "MOVE",
    "REF_ERROR",
]
