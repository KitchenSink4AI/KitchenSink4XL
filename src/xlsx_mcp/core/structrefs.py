"""Structured references (Excel table references) in formula text.

Two rules live here.

1.   Excel's TYPING shorthand for a structured reference is not its FILE
     form. A formula written as typed (=T[@v]*2) is opened by Excel as
     =#REF!*2 and the formula is lost, while the save reported verified.
     Excel itself, given the same text through Range.Formula, stores the file
     form (measured with a private hidden Excel):

        typed                      stored by Excel
        T[@v]  T[@[v]]             T[[#This Row],[v]]
        T[@]                       T[#This Row]
        [@v] [v] (inside table T)  T[[#This Row],[v]]
        t[V]                       T[v] (table and column names in their case)
        T[ v ]                     refused: no column named " v "
        T[x'#]  T[a'[b]            kept (the ' escape is the file form)
        T[[#Headers],[v]]          kept; T[ [v]:[x'#] ] kept with its spaces

     to_file_form() applies exactly those rules before a formula is stored,
     and refuses what Excel refuses (an unknown table or column, an
     unqualified reference outside every table).

2.   The table rename check looked for the old column name in brackets
     verbatim, so a reference Excel writes escaped (T[x'#] for a column named
     x#) was not seen and a rename left that formula pointing at nothing.
     columns_named() reads references through the same parser, so escaped
     and unescaped, qualified and unqualified forms are all found.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .errors import XlMcpError

#: Characters a column name escapes with ' inside a structured reference.
ESCAPED = "[]#'"

#: The special item specifiers, keyed by their folded spelling.
SPECIALS = {"#all": "#All", "#data": "#Data", "#headers": "#Headers",
            "#totals": "#Totals", "#this row": "#This Row", "#thisrow": "#This Row"}

_NAME_CHARS = "_.\\"


def escape(name: str) -> str:
    return "".join("'" + ch if ch in ESCAPED else ch for ch in name)


def unescape(text: str) -> str:
    out, i = [], 0
    while i < len(text):
        if text[i] == "'" and i + 1 < len(text):
            out.append(text[i + 1])
            i += 2
            continue
        out.append(text[i])
        i += 1
    return "".join(out)


@dataclass
class Ref:
    """One structured reference in a formula: formula[start:end] is the whole
    reference, table is the name as written (None when unqualified), body
    the raw text between its outer brackets."""

    start: int
    end: int
    table: str | None
    body: str
    columns: list[str] = field(default_factory=list)  # unescaped column names it names


def _match(text: str, i: int) -> int | None:
    """Index of the ']' closing the '[' at i, honouring ' escapes and one or
    more levels of nesting; None when unbalanced."""
    depth, k = 0, i
    while k < len(text):
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


def _items(body: str) -> list[str] | None:
    """The bracketed items of a complex specifier ("[#This Row],[v]" or
    "[v]:[w]"), raw; None when the body is not made only of bracketed items
    separated by commas, colons and spaces."""
    items, k = [], 0
    while k < len(body):
        c = body[k]
        if c in " ,:":
            k += 1
            continue
        if c != "[":
            return None
        close = _match(body, k)
        if close is None:
            return None
        items.append(body[k + 1:close])
        k = close + 1
    return items


def _columns_of(body: str) -> list[str]:
    b = body.strip()
    if b.startswith("@"):
        rest = b[1:].strip()
        if not rest:
            return []
        if rest.startswith("["):
            items = _items(rest) or []
            return [unescape(x) for x in items if not x.strip().startswith("#")]
        return [unescape(rest)]
    if b.startswith("["):
        items = _items(b) or []
        return [unescape(x) for x in items if not x.strip().startswith("#")]
    if not b or b.startswith("#"):
        return []
    return [unescape(body)]


_PARAM = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")


def _optional_param(formula: str, start: int, close: int, body: str) -> bool:
    """[y] directly in LAMBDA's own argument list, between separators and
    never last (the last argument is the body): Excel's optional-parameter
    syntax, which the formula normalizer stores as _xlop.y."""
    before = formula[:start].rstrip()[-1:]
    after = formula[close + 1:].lstrip()[:1]
    return before in ("(", ",") and after == "," and bool(_PARAM.fullmatch(body.strip()))


def find_refs(formula: str) -> list[Ref]:
    """Every structured reference in formula text, in order. String
    literals and quoted sheet names are skipped; an external workbook
    prefix ([1]Sheet1!A1, [Book.xlsx]Sheet1!A1) and a table in another
    workbook (Book.xlsx!T[v]) are not references this workbook resolves."""
    out: list[Ref] = []
    calls: list[str] = []  # the function each open parenthesis belongs to
    i, n = 0, len(formula)
    while i < n:
        ch = formula[i]
        if ch == "(":
            j = i
            while j > 0 and (formula[j - 1].isalnum() or formula[j - 1] in "._"):
                j -= 1
            calls.append(formula[j:i].upper())
            i += 1
            continue
        if ch == ")":
            if calls:
                calls.pop()
            i += 1
            continue
        if ch in "\"'":
            q = ch
            i += 1
            while i < n:
                if formula[i] == q:
                    if i + 1 < n and formula[i + 1] == q:
                        i += 2
                        continue
                    i += 1
                    break
                i += 1
            continue
        if ch != "[":
            i += 1
            continue
        close = _match(formula, i)
        if close is None:
            break  # unbalanced: leave the text to Excel's own parser
        j = i
        while j > 0 and (formula[j - 1].isalnum() or formula[j - 1] in _NAME_CHARS):
            j -= 1
        name = formula[j:i]
        before = formula[j - 1] if j > 0 else ""
        after = formula[close + 1] if close + 1 < n else ""
        if before == "!" or (name and not (name[0].isalpha() or name[0] in "_\\")):
            i = close + 1  # a table in another workbook, or not a table name
            continue
        if not name and (after.isalpha() or (after != "" and after in "_'")
                         or formula[i + 1:close].isdigit()):
            i = close + 1  # an external workbook prefix
            continue
        body = formula[i + 1:close]
        if not name and calls and calls[-1].endswith("LAMBDA") and _optional_param(
                formula, j, close, body):
            i = close + 1  # LAMBDA(x,[y],...): an optional parameter, not a column
            continue
        out.append(Ref(j, close + 1, name or None, body, _columns_of(body)))
        i = close + 1
    return out


# ------------------------------------------------------------------ tables


@dataclass
class TableInfo:
    name: str
    sheet: str
    ref: str
    columns: list[str]

    def column(self, wanted: str) -> str | None:
        """The table's own spelling of a column, matched without regard to
        case (Excel's rule), or None."""
        low = wanted.lower()
        for c in self.columns:
            if c.lower() == low:
                return c
        return None

    def contains(self, sheet: str, row: int, col: int) -> bool:
        from openpyxl.utils.cell import range_boundaries

        if sheet != self.sheet:
            return False
        c0, r0, c1, r1 = range_boundaries(self.ref)
        return r0 <= row <= r1 and c0 <= col <= c1


def tables_of(wb) -> dict[str, TableInfo]:
    """Every table of the workbook, keyed by its folded name."""
    out: dict[str, TableInfo] = {}
    for ws in wb.worksheets:
        for t in ws.tables.values():
            cols = [str(c.name) for c in (t.tableColumns or ())]
            out[t.displayName.lower()] = TableInfo(t.displayName, ws.title, t.ref, cols)
    return out


def _refuse(message: str, detail: dict | None = None) -> None:
    exc = XlMcpError(message)
    if detail:
        exc.detail = detail
    raise exc


def _column_or_refuse(info: TableInfo, wanted: str, written: str) -> str:
    hit = info.column(wanted)
    if hit is None:
        stripped = info.column(wanted.strip())
        hint = (f" Write {info.name}[{escape(stripped)}] (no spaces inside the brackets)."
                if stripped is not None else "")
        _refuse(f"the formula names {written}, but table {info.name} has no column "
                f"{wanted!r}; Excel refuses this formula. Columns: "
                f"{', '.join(info.columns)}.{hint}",
                {"table": info.name, "columns": info.columns})
    return hit


def to_file_form(formula: str, tables: dict[str, TableInfo], sheet: str | None,
                 row: int | None = None, col: int | None = None) -> str:
    """The formula with every structured reference in the form Excel stores
    (module docstring). Raises XlMcpError (BAD_PARAMS) for what Excel would
    refuse. A formula with no structured reference comes back unchanged."""
    if "[" not in formula:
        return formula
    refs = find_refs(formula)
    if not refs:
        return formula
    out = formula
    for ref in reversed(refs):
        written = formula[ref.start:ref.end]
        if ref.table is not None:
            info = tables.get(ref.table.lower())
            if info is None:
                _refuse(f"the formula names {written}, but this workbook has no table named "
                        f"{ref.table!r}; Excel would show #REF!. Tables: "
                        f"{', '.join(t.name for t in tables.values()) or 'none'}. Create the "
                        "table first (xl_structure create_table).",
                        {"tables": [t.name for t in tables.values()]})
        else:
            info = next((t for t in tables.values() if sheet is not None and row is not None
                         and col is not None and t.contains(sheet, row, col)), None)
            if info is None:
                _refuse(f"the formula uses {written} with no table name, which works only "
                        "in a cell inside a table; name the table, for example "
                        f"T{written} for a table named T.")
        body = ref.body.strip()
        if body.startswith("@"):
            rest = body[1:].strip()
            if not rest:
                new_body = "#This Row"
            elif rest.startswith("["):
                items = _items(rest)
                if items is None:
                    _refuse(f"cannot read the structured reference {written}")
                parts = []
                for raw in items:
                    parts.append("[" + escape(_column_or_refuse(info, unescape(raw), written)) + "]")
                sep = ":" if ":" in rest else ","
                new_body = "[#This Row]," + sep.join(parts)
            else:
                new_body = "[#This Row],[" + escape(_column_or_refuse(info, unescape(rest),
                                                                       written)) + "]"
        elif body.startswith("["):
            items = _items(body)
            if items is None:
                _refuse(f"cannot read the structured reference {written}")
            for raw in items:
                item = raw.strip()
                if item.startswith("#"):
                    if item.lower() not in SPECIALS:
                        _refuse(f"{written} uses {item}, which is not one of #All, #Data, "
                                "#Headers, #Totals or #This Row")
                    continue
                _column_or_refuse(info, unescape(raw), written)
            new_body = ref.body  # Excel keeps a complex specifier as written
        elif not ref.body:
            new_body = ""
        elif ref.body.strip().startswith("#"):
            special = SPECIALS.get(ref.body.strip().lower())
            if special is None:
                _refuse(f"{written} uses {ref.body.strip()}, which is not one of #All, "
                        "#Data, #Headers, #Totals or #This Row")
            new_body = special
        else:
            column = _column_or_refuse(info, unescape(ref.body), written)
            if ref.table is None:
                # Excel stores an unqualified column reference typed inside
                # the table as the current row of that column.
                new_body = "[#This Row],[" + escape(column) + "]"
            else:
                new_body = escape(column)
        out = out[:ref.start] + info.name + "[" + new_body + "]" + out[ref.end:]
    return out


def columns_named(formula: str, table: str, column: str, *, sheet: str | None = None,
                  row: int | None = None, col: int | None = None,
                  tables: dict[str, TableInfo] | None = None) -> bool:
    """Whether the formula names `column` of `table`: qualified
    references to that table, and unqualified ones written in a cell inside
    it. Names compare without regard to case, escaped or not."""
    if "[" not in formula:
        return False
    low_t, low_c = table.lower(), column.lower()
    info = (tables or {}).get(low_t)
    for ref in find_refs(formula):
        if ref.table is not None:
            if ref.table.lower() != low_t:
                continue
        elif info is None or row is None or col is None or sheet is None \
                or not info.contains(sheet, row, col):
            continue
        if any(c.lower() == low_c or c.strip().lower() == low_c for c in ref.columns):
            return True
    return False


__all__ = ["Ref", "TableInfo", "find_refs", "tables_of", "to_file_form", "columns_named",
           "escape", "unescape"]
