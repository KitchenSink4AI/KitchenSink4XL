"""ops/names.py: defined names / named ranges.

manage_name is the defined-name lifecycle: add, delete, rename, update the
refers-to, and list. Names live at workbook scope or sheet scope, and the two
can collide (a sheet-scoped name shadows a workbook-scoped one), so every action
that targets an existing name disambiguates by an optional scope. The reserved
print-area and filter names (the _xlnm.* built-ins) are protected from delete
and rename so a cleanup never breaks a workbook's print or filter setup.

get_workbook_metadata already lists names for discovery; this tool is the write
path. Mutations route through WorkbookPackage for the backup and verify.

A definition is stored the way Excel stores it, with the same prefixes as a
cell formula (core.calc.normalize_formula): _xlfn. on newer functions,
_xlpm. on LET and LAMBDA names, _xlop. on an optional LAMBDA parameter.
MEASURED in Excel: a name stored as LAMBDA(a,a+1) reads as a call to an
unknown add-in function, so every cell that calls it is #NAME?; stored as
_xlfn.LAMBDA(_xlpm.a,_xlpm.a+1), as Excel itself writes it, it calculates.
Reads show a definition as written, as Excel's Name Manager does.
"""

from __future__ import annotations

import re as _re
from typing import Any

from ..core import calc as _calc
from ..core import limits as _limits
from ..core import locate as _locate
from ..core.errors import AmbiguousTarget, TargetNotFound, XlMcpError
from ..core.package import WorkbookPackage

NAME_ACTIONS = ("add", "delete", "rename", "update", "list")

#: Excel's built-in names (print area, print titles, autofilter range). Guarded
#: against delete/rename so a name cleanup cannot break print or filter setup.
_RESERVED_PREFIXES = ("_xlnm.", "_xlnm._")


def _is_reserved(name: str) -> bool:
    return name.startswith("_xlnm")


def _scopes(wb):
    yield "workbook", wb.defined_names
    for ws in wb.worksheets:
        dn = getattr(ws, "defined_names", None)
        if dn is not None:
            yield ws.title, dn


def _find(wb, name: str, scope: str | None):
    hits = []
    for scope_name, dnd in _scopes(wb):
        if name in dnd:
            hits.append((scope_name, dnd, dnd[name]))
    if scope is not None:
        hits = [h for h in hits if h[0].lower() == scope.lower()]
    if not hits:
        allnames = sorted({n for _s, dnd in _scopes(wb) for n in dnd})
        raise TargetNotFound(
            f"no defined name {name!r}"
            + (f" in scope {scope!r}" if scope else "")
            + "; names: " + (", ".join(repr(n) for n in allnames[:25])
                             or "none defined"))
    if len(hits) > 1:
        exc = AmbiguousTarget(
            f"defined name {name!r} exists in {len(hits)} scopes; pass scope "
            "(a sheet title or 'workbook') to disambiguate")
        exc.matches = [{"scope": s, "value": getattr(d, "value", None)}
                       for s, _dnd, d in hits]
        raise exc
    return hits[0]


#: Excel's grammar for a defined name: first character a LETTER (in any
#: script), underscore, or backslash; then letters, digits, periods,
#: underscores; no spaces; at most 255 characters; and never something Excel
#: would read as a cell reference (A1 style or R1C1 style), nor the bare
#: letters R and C.
#:
#: "Letter" is not "A-Za-z". The class was written ASCII-only under a
#: re.UNICODE flag that could not widen a literal range, so Excel-legal
#: non-ASCII names were refused: the Korean name 환율 was turned away
#: with a message claiming Excel requires a letter, which is
#: exactly what 환율 is. [^\W\d] under UNICODE is "a word character
#: that is not a digit", i.e. any script's letter plus the underscore.
_NAME_RE = _re.compile(r"(?:[^\W\d]|\\)[\w.\\]*", _re.UNICODE)
#: Which characters Excel reads as DIGITS when it decides whether a name looks
#: like a cell reference. Measured in Excel itself (Names.Add on every Unicode
#: character that carries a digit value): the ASCII digits, and the digits of
#: these scripts, each block counted from its zero: Arabic-Indic, Extended
#: Arabic-Indic, Devanagari, Bengali, Gurmukhi, Gujarati, Oriya, Tamil,
#: Telugu, Kannada, Malayalam, Sinhala, Thai, Lao, Tibetan, Myanmar, Khmer,
#: Mongolian, Tai Tham (two blocks), Javanese, Myanmar Tai Laing, Meetei Mayek,
#: plus U+19DA. Not digits to Excel: fullwidth digits, mathematical digits,
#: circled or subscript digits, NKo and the other scripts not listed. The
#: superscripts 1, 2 and 3 count as digits after a leading R or C but not
#: inside an A1-style reference.
_DIGIT_BLOCKS = (0x30, 0x660, 0x6F0, 0x966, 0x9E6, 0xA66, 0xAE6, 0xB66,
                 0xBE6, 0xC66, 0xCE6, 0xD66, 0xDE6, 0xE50, 0xED0, 0xF20,
                 0x1040, 0x17E0, 0x1810, 0x1A80, 0x1A90, 0xA9D0, 0xA9F0,
                 0xABF0)
_DIGIT_VALUE = {chr(start + i): i for start in _DIGIT_BLOCKS for i in range(10)}
_DIGIT_VALUE["\u19da"] = 1
_SUPERSCRIPT_VALUE = {"\u00b9": 1, "\u00b2": 2, "\u00b3": 3}
_MAX_ROW, _MAX_COL = 1048576, 16384


def _digits(name: str, i: int, *, superscripts: bool) -> tuple[int | None, int]:
    """(value, end) of the run of Excel-digits starting at name[i]; value is
    None when there is none."""
    j, value = i, 0
    while j < len(name):
        d = _DIGIT_VALUE.get(name[j])
        if d is None and superscripts:
            d = _SUPERSCRIPT_VALUE.get(name[j])
        if d is None:
            break
        value = value * 10 + d
        j += 1
    return (value if j > i else None), j


def _a1_reference(token: str) -> bool:
    """One to three ASCII letters then a row number, with the column within
    XFD and the row within 1 to 1,048,576 (A1, XFD1, TAX2023)."""
    letters = 0
    while letters < len(token) and token[letters].isascii()             and token[letters].isalpha():
        letters += 1
    if not 1 <= letters <= 3:
        return False
    row, end = _digits(token, letters, superscripts=False)
    if row is None or end != len(token) or not 1 <= row <= _MAX_ROW:
        return False
    col = 0
    for ch in token[:letters].upper():
        col = col * 26 + ord(ch) - 64
    return col <= _MAX_COL


def _r1c1_reference(token: str) -> bool:
    """R<row>, C<col>, RC<col>, R<row>C or R<row>C<col> with every number
    present a real row or column, as a whole token."""
    low, i, found = token.lower(), 0, False
    if low.startswith("r"):
        i = 1
        num, end = _digits(token, i, superscripts=True)
        if num is not None:
            if not 1 <= num <= _MAX_ROW:
                return False
            found, i = True, end
    if i < len(low) and low[i] == "c":
        num, end = _digits(token, i + 1, superscripts=True)
        if num is not None:
            if not 1 <= num <= _MAX_COL:
                return False
            found, i = True, end
        else:
            i += 1
    return found and i == len(low)


def _reads_as_a_reference(name: str) -> bool:
    """True when Excel will not take `name` as a defined name because it
    reads as (or begins like) a cell reference. The rule is Excel's own,
    measured with Names.Add and by opening files that carry the name:
    - A1 style: one to three letters then a row number is refused only when
      the column is within XFD and the row within 1 to 1,048,576 (A1, XFD1,
      TAX2023); A0, XFE1, ZZZ1 and A1048577 are fine.
    - R1C1 style: a name that BEGINS with R, C or RC followed by a number is
      refused when that number is a real row (1 to 1,048,576) or column (1 to
      16,384) and the next character is not a period: R1_x, R2D2, C21_Name,
      Rc21_Name, R1C1x, R1C0_x (the R1 is a real row) are refused; R0_x,
      C0_x, R1048577_x, C16385_x, RC0_x, R0C1_x and R999999999999_x are
      accepted, and so are R1.x and C1.x. A number is a run of Excel-digits
      (see above), ASCII or not: R and then the Arabic-Indic digit two is
      refused like R2. After a superscript digit a period does not save
      the name.
    - The bare R, C and RC are refused.
    - Two references joined by a period (R1.R1, A1.A1, R1.D5) are refused."""
    low = name.lower()
    if low in ("r", "c", "rc"):
        return True
    for lead, limit in (("rc", _MAX_COL), ("r", _MAX_ROW), ("c", _MAX_COL)):
        if low.startswith(lead):
            num, end = _digits(name, len(lead), superscripts=True)
            if num is not None and 1 <= num <= limit:
                supers = any(ch in _SUPERSCRIPT_VALUE
                             for ch in name[len(lead):end])
                if name[end:end + 1] != "." or supers:
                    return True
    parts = name.split(".")
    refs = [_a1_reference(p) or _r1c1_reference(p) for p in parts]
    return any(a and b for a, b in zip(refs, refs[1:])) or (
        len(parts) == 1 and refs[0])


def _validate_name(name: str) -> str:
    """Refuse a defined name Excel's own grammar rejects.

    This is a corruption guard, not a style check: adding the name '1bad'
    through this tool saved and verified clean, and
    then EXCEL REFUSED TO OPEN THE FILE (repair prompt). Nothing downstream
    can catch that, because the part inventory and the XML are both intact;
    only the grammar is wrong."""
    if not isinstance(name, str) or not name.strip():
        raise XlMcpError("name must be a non-empty string")
    if len(name) > 255:
        raise XlMcpError(
            f"defined name {name[:40]!r}... is {len(name)} characters; "
            "Excel's limit is 255")
    if name.startswith("_xlnm"):
        return name  # the built-ins, handled by the reserved-name guards
    if not _NAME_RE.fullmatch(name):
        raise XlMcpError(
            f"{name!r} is not a valid defined name: Excel requires the first "
            "character to be a letter (any script), underscore, or backslash "
            "and the rest to be letters, digits, periods, or underscores (no "
            "spaces, no operators). Excel refuses to open a workbook carrying "
            "an invalid name, so this is refused here.")
    if _reads_as_a_reference(name):
        raise XlMcpError(
            f"{name!r} reads as a cell reference, which Excel does not allow "
            "as a defined name; pick a name that cannot be parsed as an "
            "address")
    return name


def _validate_refers_to(rt: str) -> str:
    """Cheap syntax sanity on a definition body.

    Not a formula parser: it catches the shapes that make Excel demand a
    repair on open (unbalanced parentheses or quotes, a definition that is
    only operators, a stray leading operator). '(((' passed before, and the
    workbook it produced would not open."""
    depth = 0
    in_quote = False
    for ch in rt:
        if ch == '"':
            in_quote = not in_quote
            continue
        if in_quote:
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth < 0:
                break
    if in_quote or depth != 0:
        raise XlMcpError(
            f"refers_to {rt[:60]!r} has unbalanced parentheses or quotes; "
            "Excel refuses to open a workbook whose defined name does not "
            "parse")
    if rt.count("'") % 2:
        raise XlMcpError(
            f"refers_to {rt[:60]!r} has an unbalanced sheet-name quote (')")
    if not _re.search(r"[A-Za-z0-9_$]", rt):
        raise XlMcpError(
            f"refers_to {rt[:60]!r} holds no reference, number, or name; "
            "give an A1 reference like Sheet1!$A$1:$B$9 or a formula")
    if rt[-1] in "+-*/^&,=<>(":
        raise XlMcpError(
            f"refers_to {rt[:60]!r} ends on an operator; the definition is "
            "incomplete")
    return rt


def _normalize_refers_to(refers_to: str, what: str = "the definition") -> str:
    """Defined-name definitions are stored WITHOUT a leading '=' (workbook.xml
    holds the bare expression). A definition saved with '=' breaks openpyxl's
    destinations parsing, which makes the name unresolvable by the locate
    layer, so a convenience '=' from the caller is stripped.

    The result is the definition as Excel stores it: newer functions and
    LET or LAMBDA names carry Excel's prefixes, exactly as a cell formula
    does. A definition that cannot be given them (a space between a
    function and its parenthesis, an unbalanced LET or LAMBDA) is refused
    rather than stored as something Excel cannot calculate."""
    rt = str(refers_to).strip()
    if rt.startswith("="):
        rt = rt[1:].strip()
    if not rt:
        raise XlMcpError("refers_to must be a non-empty A1 reference or formula")
    rt = _validate_refers_to(rt)
    # Excel's 8,192-character limit, on the definition as it is stored.
    _limits.check_formula_length(rt, what=what)
    stored, _prefixed = _calc.normalize_formula(rt)
    return stored


def _shown(stored):
    """A stored definition as written, without Excel's storage prefixes."""
    return _calc.denormalize_formula(stored) if isinstance(stored, str) else stored


def manage_name(path: str, action: str, name: str | None = None,
                refers_to: str | None = None, scope: str | None = None,
                new_name: str | None = None,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Add / delete / rename / update / list defined names. Backup + verify."""
    if action not in NAME_ACTIONS:
        raise XlMcpError(f"action must be one of {NAME_ACTIONS}, got {action!r}")

    if action == "list":
        wb = None
        from . import gridio
        wb = gridio.open_wb(path, data_only=False)
        try:
            out = []
            for scope_name, name_obj, defn in _locate._all_defined_names(wb):
                out.append({
                    "name": name_obj, "scope": scope_name,
                    "refers_to": _shown(getattr(defn, "value", None)),
                    "reserved": _is_reserved(name_obj)})
            # Excel's built-in names (print area, print titles, the filter
            # range): openpyxl folds them into sheet settings on load, so
            # they are read from the file itself.
            from ..core import pkginfo as _pkginfo
            have = {(o["scope"], o["name"]) for o in out}
            try:
                builtins = _pkginfo.builtin_names(path)
            except Exception:  # noqa: BLE001 - the list still answers
                builtins = []
            for scope_name, name_obj, value, hidden in builtins:
                if (scope_name, name_obj) in have:
                    continue
                entry = {"name": name_obj, "scope": scope_name,
                         "refers_to": _shown(value), "reserved": True}
                if hidden:
                    entry["hidden"] = True
                out.append(entry)
            return {"names": out, "count": len(out)}
        finally:
            wb.close()

    from openpyxl.workbook.defined_name import DefinedName
    pkg = WorkbookPackage.open(path)
    wb = pkg.workbook
    detail: dict[str, Any] = {"action": action}

    if action == "add":
        if not name or not refers_to:
            raise XlMcpError("add needs name and refers_to")
        _validate_name(name)
        rt = _normalize_refers_to(refers_to, what=f"the definition of {name}")
        if scope in (None, "workbook"):
            target = wb.defined_names
            scope_name = "workbook"
        else:
            if scope not in wb.sheetnames:
                raise TargetNotFound(f"no sheet named {scope!r} for the scope")
            target = wb[scope].defined_names
            scope_name = scope
        if name in target:
            raise XlMcpError(
                f"a name {name!r} already exists in scope {scope_name!r}")
        target.add(DefinedName(name=name, attr_text=rt))
        detail.update(name=name, scope=scope_name, refers_to=_shown(rt))

    elif action == "delete":
        if not name:
            raise XlMcpError("delete needs name")
        if _is_reserved(name):
            raise XlMcpError(
                f"{name!r} is a reserved built-in name (print area / titles / "
                "filter); it is protected from deletion")
        scope_name, dnd, _defn = _find(wb, name, scope)
        del dnd[name]
        detail.update(name=name, scope=scope_name)

    elif action == "rename":
        if not name or not new_name:
            raise XlMcpError("rename needs name and new_name")
        _validate_name(new_name)
        if _is_reserved(name):
            raise XlMcpError(f"{name!r} is a reserved built-in name; protected")
        scope_name, dnd, defn = _find(wb, name, scope)
        if new_name in dnd:
            raise XlMcpError(
                f"a name {new_name!r} already exists in scope {scope_name!r}")
        rt = defn.value
        del dnd[name]
        dnd.add(DefinedName(name=new_name, attr_text=rt))
        detail.update(renamed_from=name, renamed_to=new_name, scope=scope_name)

    else:  # update
        if not name or not refers_to:
            raise XlMcpError("update needs name and refers_to")
        scope_name, _dnd, defn = _find(wb, name, scope)
        rt = _normalize_refers_to(refers_to, what=f"the definition of {name}")
        defn.attr_text = rt
        detail.update(name=name, scope=scope_name, refers_to=_shown(rt))

    pkg._changed["name"] = detail
    return pkg.save(allow_loss=allow_loss, backup=backup,
                    verify_com=verify_com)


__all__ = ["manage_name", "NAME_ACTIONS"]
