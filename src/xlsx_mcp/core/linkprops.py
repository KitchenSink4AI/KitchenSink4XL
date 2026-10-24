"""core/linkprops.py: custom document properties that are linked to a cell.

A custom document property can be linked to a defined name
(docProps/custom.xml: <property name="Rate" linkTarget="LinkedRate">). Excel
refreshes the cell the name points at from the property when the workbook
opens, so a cell edit that leaves the property as it was is undone the moment
the file is opened, after the edit was answered "verified".

sync() runs on the package a save produced, whichever route produced it (the
byte-preserving cell edit or the full rewrite), BEFORE the verify step. For
each linked property whose cell holds something different from what the
original held, it writes the new value into the property, keeping the
property's own type (what Excel does when a linked cell changes). A value the
property cannot hold (text for a number, a formula, an error, an empty cell for
a number) is refused naming the property and the cell, unless allow_loss is
passed: then the property is left as it was and a warning names it.

Only a name that resolves to ONE cell, in the original and in the produced
package, is followed; a row or column inserted above the cell moves the cell
and the name together and changes no property.

A date property holds a UTC instant: the cell's date and time is numbered as
Excel numbers it (1 January 1900 is 1; 29 February 1900 is not a date) and
converted from the computer's time zone, as Excel does.

The property part and the workbook part are read as XML (core/xmlsafe.py,
namespace-aware, any encoding: UTF-8, UTF-16, UTF-32, CDATA, entities), never
by their bytes, so a link or a defined name spelled differently is found all
the same. A property part is written back in the encoding it had.
"""

from __future__ import annotations

import datetime as _dt
import decimal as _decimal
import html
import math
import re
import zipfile
from dataclasses import dataclass, field

from . import carry as _carry
from . import surgical as _s
from . import xmlsafe as _xmlsafe
from .errors import HazardRefused

#: Property types by the way a cell's content goes into them.
_TEXT_TYPES = frozenset({"lpwstr", "lpstr", "bstr"})
_REAL_TYPES = frozenset({"r8", "r4", "cy", "decimal"})
#: The exact range each real type holds, as (lowest, highest): binary32 for
#: r4, a 64-bit integer of ten-thousandths for cy (one more below zero than
#: above), the 96-bit decimal. Exact numbers, never floats: the float nearest
#: to the cy or decimal limit lies just past it. For r4 below the smallest
#: nonzero magnitude (_R4_TINY) the value becomes zero.
_R4_MAX = _decimal.Decimal(3.4028234663852886e38)
_REAL_BOUNDS = {
    "r4": (-_R4_MAX, _R4_MAX),
    "cy": (_decimal.Decimal("-922337203685477.5808"),
           _decimal.Decimal("922337203685477.5807")),
    "decimal": (_decimal.Decimal(-(2**96 - 1)), _decimal.Decimal(2**96 - 1)),
}
_R4_TINY = 1.401298464324817e-45
_INT_RANGES = {
    "i1": (-2**7, 2**7 - 1), "i2": (-2**15, 2**15 - 1),
    "i4": (-2**31, 2**31 - 1), "int": (-2**31, 2**31 - 1),
    "i8": (-2**63, 2**63 - 1),
    "ui1": (0, 2**8 - 1), "ui2": (0, 2**16 - 1),
    "ui4": (0, 2**32 - 1), "uint": (0, 2**32 - 1), "ui8": (0, 2**64 - 1),
}
_DATE_TYPES = frozenset({"date", "filetime"})

_CELL_REF = re.compile(
    r"^(?:(?P<sheet>'(?:[^']|'')+'|[^'!]+)!)?\$?(?P<col>[A-Za-z]{1,3})\$?(?P<row>[0-9]+)"
    r"(?::\$?(?P=col)\$?(?P=row))?$")
_T_RE = re.compile(rb"<(?:[\w.\-]+:)?t\b[^>]*?(?:/>|>(.*?)</(?:[\w.\-]+:)?t>)", re.S)
_V_RE = re.compile(rb"<(?:[\w.\-]+:)?v\b[^>]*?(?:/>|>(.*?)</(?:[\w.\-]+:)?v>)", re.S)
_IS_RE = re.compile(rb"<(?:[\w.\-]+:)?is\b[^>]*?>(.*?)</(?:[\w.\-]+:)?is>", re.S)
_F_RE = re.compile(rb"<(?:[\w.\-]+:)?f\b[^>]*?(?:/>|>(.*?)</(?:[\w.\-]+:)?f>)", re.S)
_RPH_RE = re.compile(rb"<(?:[\w.\-]+:)?rPh\b.*?</(?:[\w.\-]+:)?rPh>", re.S)


class _Unreadable(Exception):
    """A part could not be read well enough to compare a cell."""


@dataclass
class _Prop:
    name: str
    target: str
    tag: str
    text: str
    index: int  # position among the part's property elements


@dataclass
class SyncResult:
    changed: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


# ------------------------------------------------------------ reading parts

def _custom_part(pkg) -> str | None:
    for r in pkg.rels(""):
        if r.get("Type", "").endswith("/custom-properties"):
            part = _s.resolve_target("", r.get("Target", ""))
            return pkg.real(part)
    return pkg.real("docProps/custom.xml")


def _local(tag) -> str:
    """The local name of an element tag; "" for a comment or instruction."""
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def _may_hold_link(data: bytes) -> bool:
    """False only when the property part cannot hold a linkTarget attribute:
    an ASCII-compatible part that does not spell the word. Anything else (UTF-16
    or UTF-32, a part that does not start with "<") takes the parsed pass."""
    body = data[3:] if data.startswith(b"\xef\xbb\xbf") else data
    if _xmlsafe.byte_encoding(data) != "utf-8" or body[:1] != b"<":
        return True
    return b"linkTarget" in data


def _parse_properties(data: bytes):
    """(the root element with its comments and instructions kept, the linked
    properties in document order). xmlsafe refuses a part that declares a DTD
    and raises ParseError on one that is not well-formed."""
    root = _xmlsafe.fromstring(data, doc="the workbook", part="docProps/custom.xml",
                         remove_comments=False, remove_pis=False)
    out = []
    for i, el in enumerate(c for c in root if _local(c.tag) == "property"):
        target = el.get("linkTarget")
        if target is None:
            continue
        value = next((c for c in el if isinstance(c.tag, str)), None)
        tag = _local(value.tag) if value is not None else ""
        text = (value.text or "") if value is not None else ""
        out.append((_Prop(el.get("name", ""), target, tag, text, i), el, value))
    return root, out


_BOMS = (b"\xef\xbb\xbf", b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff", b"\xff\xfe", b"\xfe\xff")
_DECL = re.compile(r"<\?xml.*?\?>[ \t\r\n]*", re.S)
_DECL_ENCODING = re.compile(rb"<\?xml[^>]*?encoding\s*=\s*[\"']([A-Za-z0-9._\-]+)[\"']")


def _serialize(original: bytes, root) -> bytes:
    """The property part as `original` was spelled: its XML declaration
    verbatim, its byte order mark and its encoding."""
    kind = _xmlsafe.byte_encoding(original)
    bom = b""
    for mark in _BOMS:
        if original.startswith(mark) and (kind != "utf-8" or mark == _BOMS[0]):
            bom = mark
            break
    body = original[len(bom):]
    codec = kind
    if kind == "utf-8":
        m = _DECL_ENCODING.match(body)
        codec = m.group(1).decode("ascii") if m else "utf-8"
    text = body.decode(codec)
    head = _DECL.match(text)
    markup = _xmlsafe.etree.tostring(root.getroottree(), encoding="unicode")
    return bom + ((head.group(0) if head else "") + markup).encode(codec, "xmlcharrefreplace")


def _is_1904(wb_root) -> bool:
    for el in wb_root.iter():
        if _local(el.tag) == "workbookPr":
            return (el.get("date1904") or "").strip().lower() in ("1", "true")
    return False


def _workbook_root(pkg):
    try:
        return _xmlsafe.fromstring(pkg.read(pkg.workbook_part()), doc="the workbook",
                             part="the workbook part")
    except _xmlsafe.XML_ERRORS as exc:
        raise _Unreadable(f"the workbook part could not be read ({exc})") from exc


def _sheet_parts(pkg, wb_root) -> dict[str, str]:
    """{sheet title: worksheet part}, from the parsed workbook part."""
    wbp = pkg.workbook_part()
    ids = {r.get("Id"): _s.resolve_target(wbp, r.get("Target", "")) for r in pkg.rels(wbp)}
    out = {}
    for el in wb_root.iter():
        if _local(el.tag) != "sheet" or el.get("name") is None:
            continue
        rid = next((v for k, v in el.attrib.items() if k.rsplit("}", 1)[-1] == "id"), None)
        part = ids.get(rid)
        if part:
            out[el.get("name")] = part
    return out


def _name_cell(wb_root, name: str) -> tuple[str, str] | None:
    """The (sheet title, A1 address) a defined name points at when that is one
    cell; None for anything else (a range, a formula, a missing name, a name
    defined more than once)."""
    hits = [("".join(el.itertext())).strip()
            for el in wb_root.iter()
            if _local(el.tag) == "definedName"
            and (el.get("name") or "").lower() == name.lower()]
    if len(hits) != 1:
        return None
    ref = _CELL_REF.match(hits[0])
    if ref is None or ref.group("sheet") is None:
        return None
    sheet = ref.group("sheet")
    if sheet.startswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, f"{ref.group('col').upper()}{ref.group('row')}"


def _shared_strings(pkg) -> list[str]:
    wbp = pkg.workbook_part()
    parts = pkg.related(wbp, "/sharedStrings")
    if not parts or not pkg.has(parts[0]):
        return []
    data = pkg.read(parts[0])
    out = []
    for m in _s._SI_RE.finditer(data):
        body = m.group(2)
        if body is None:
            out.append("")
            continue
        body = _RPH_RE.sub(b"", body)
        out.append("".join(_s._sst_text(t.group(1) or b"") for t in _T_RE.finditer(body)))
    return out


class _Reader:
    """Cell contents of one package, by (sheet title, address)."""

    def __init__(self, pkg, parts: dict[str, str]):
        self.pkg = pkg
        self.parts = parts
        self._sst: list[str] | None = None
        self._data: dict[str, bytes] = {}

    def _strings(self) -> list[str]:
        if self._sst is None:
            self._sst = _shared_strings(self.pkg)
        return self._sst

    def cell(self, sheet: str, coord: str) -> tuple:
        """(kind, value, formula_text, has_cached): kind is empty, num, str,
        bool, err or formula (value is then the cached value's (kind, value)
        pair, or None)."""
        part = self.parts.get(sheet)
        if part is None:
            raise _Unreadable(f"no sheet {sheet}")
        if part not in self._data:
            self._data[part] = self.pkg.read(part)
        data = self._data[part]
        ref = _CELL_REF.match(coord)
        row = int(ref.group("row"))
        col = _s._col_index(ref.group("col").upper())
        try:
            scan = _s.scan_sheet(data, [row])
        except _s.NotSurgical as exc:
            raise _Unreadable(str(exc)) from exc
        loc = None
        if row in scan.rows:
            loc = next((c for c in scan.rows[row].cells if c.col == col), None)
        if loc is None or loc.selfclose:
            return ("empty", None, None, False)
        attrs = _s._attrs(loc.attrs)
        t = attrs.get("t", b"n").decode("ascii", "replace")
        body = data[loc.tag_end:loc.end]
        fm = _F_RE.search(body)
        vm = _V_RE.search(body)
        raw = html.unescape((vm.group(1) or b"").decode("utf-8", "replace")) \
            if vm is not None else None
        value = self._typed(t, raw, body)
        if fm is not None:
            ftext = _formula_text(fm.group(1) or b"")
            return ("formula", value, ftext, value is not None)
        if value is None:
            return ("empty", None, None, False)
        return (value[0], value[1], None, True)

    def _typed(self, t: str, raw: str | None, body: bytes):
        if t == "inlineStr":
            m = _IS_RE.search(body)
            if m is None:
                return None
            text = "".join(_s._sst_text(x.group(1) or b"")
                           for x in _T_RE.finditer(_RPH_RE.sub(b"", m.group(1))))
            return ("str", text)
        if raw is None or raw == "":
            return None
        if t == "s":
            try:
                return ("str", self._strings()[int(raw)])
            except (ValueError, IndexError) as exc:
                raise _Unreadable("a shared string index names no string") from exc
        if t == "str":
            return ("str", raw)
        if t == "b":
            return ("bool", raw.strip() in ("1", "true"))
        if t == "e":
            return ("err", raw)
        if t == "d":
            return ("str", raw)
        try:
            return ("num", float(raw))
        except ValueError as exc:
            raise _Unreadable("a number cell holds text") from exc


_CDATA = re.compile(r"<!\[CDATA\[(.*?)\]\]>", re.S)


def _formula_text(raw: bytes) -> str:
    """A formula element's text as an XML reader gives it: CDATA and
    entities resolved, line ends as single line feeds, nothing else changed."""
    text = raw.decode("utf-8", "replace")
    text = _CDATA.sub(lambda m: html.escape(m.group(1), quote=False), text)
    text = html.unescape(text).replace("\r\n", "\n").replace("\r", "\n")
    return text.strip()


def _bare(formula: str) -> str:
    """A formula without the leading "=" Excel ignores; every other character,
    a space inside a string literal included, is part of the formula."""
    return formula[1:] if formula.startswith("=") else formula


def _changed(old: tuple, new: tuple) -> bool:
    """Did the edit change what the cell holds?"""
    if old[0] == "formula" or new[0] == "formula":
        if old[0] != new[0]:
            return True
        of, nf = old[2], new[2]
        if of and nf and _bare(of) != _bare(nf):
            return True
        ov, nv = old[1], new[1]
        if ov is None and nv is None:
            return False
        if ov is None or nv is None:
            return True  # a result that was stored and is no longer
        return ov != nv
    return (old[0], old[1]) != (new[0], new[1])


# ------------------------------------------------------------- conversion

_XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


class _Cannot(Exception):
    pass


def _number_text(v: float) -> str:
    if math.isfinite(v) and v == int(v) and abs(v) < 1e15:
        return str(int(v))
    return _s.number_text(float(v))


def _utc_offset(local: _dt.datetime) -> _dt.timedelta:
    """The offset from UTC Excel applies to a wall-clock date in a date
    property. A date property holds a UTC instant and Excel shows (and, for a
    linked cell, fills in) the computer's local time for it. Measured in
    Excel, in a UTC+9 zone: the cell 1 January 1900 is stored as
    1899-12-31T15:00:00Z, so the zone rules in force now apply whatever the
    year. The same day of the current year stands in for the date."""
    year = _dt.datetime.now().year
    try:
        day = 28 if (local.month, local.day) == (2, 29) else local.day
        return local.replace(year=year, day=day).astimezone().utcoffset() or _dt.timedelta(0)
    except (OverflowError, OSError, ValueError):
        return _dt.timedelta(0)


def _property_text(tag: str, new: tuple, date1904: bool) -> str:
    """The text of the property's value element for the cell's new content,
    in the property's own type; _Cannot says why it cannot be held."""
    kind = new[0]
    if kind == "formula":
        raise _Cannot("the cell now holds a formula, whose result this server cannot calculate")
    if kind == "err":
        raise _Cannot("the cell now holds an error value")
    if tag in _TEXT_TYPES:
        if kind == "empty":
            return ""
        if kind == "str":
            if _XML_BAD.search(new[1]):
                raise _Cannot("the text holds a character an XML document cannot hold")
            return new[1]
        if kind == "num":
            return _number_text(new[1])
        raise _Cannot("the cell now holds a TRUE/FALSE value")
    if tag in _REAL_TYPES:
        if kind != "num":
            raise _Cannot("the cell no longer holds a number")
        v = new[1]
        if not math.isfinite(v):
            raise _Cannot(f"the number {_number_text(v)} is outside the range this "
                          f"type ({tag}) can hold")
        text = _number_text(v)
        bounds = _REAL_BOUNDS.get(tag)
        if bounds is not None:
            # Both the number the cell holds and the text that would be
            # written must be inside the exact range.
            lo, hi = bounds
            if not (lo <= _decimal.Decimal(v) <= hi
                    and lo <= _decimal.Decimal(text) <= hi):
                raise _Cannot(f"the number {text} is outside the range this "
                              f"type ({tag}) can hold")
        if tag == "r4" and v != 0 and abs(v) < _R4_TINY:
            raise _Cannot(f"the number {_number_text(v)} is too close to zero for "
                          f"this type ({tag}), which would store it as 0")
        return text
    if tag in _INT_RANGES:
        if kind != "num":
            raise _Cannot("the cell no longer holds a number")
        v = new[1]
        lo, hi = _INT_RANGES[tag]
        if not math.isfinite(v) or v != int(v) or not lo <= int(v) <= hi:
            raise _Cannot(f"the number {_number_text(v)} is not a whole number "
                          f"this type ({tag}) can hold")
        return str(int(v))
    if tag == "bool":
        if kind == "bool":
            return "true" if new[1] else "false"
        raise _Cannot("the cell no longer holds a TRUE/FALSE value")
    if tag in _DATE_TYPES:
        if kind != "num":
            raise _Cannot("the cell no longer holds a date")
        if not math.isfinite(new[1]):
            raise _Cannot("the number is not a date Excel can show")
        secs = round(new[1] * 86400)
        if date1904:
            base = _dt.datetime(1904, 1, 1)
            if secs < 0:
                raise _Cannot(f"the number {_number_text(new[1])} is before the first "
                              "date of the 1904 date system (1 January 1904)")
        else:
            # Excel numbers 1 January 1900 as 1 and counts a 29 February 1900
            # that never existed as 60, so a date before March 1900 is one
            # day later than the calendar count from 30 December 1899.
            base = _dt.datetime(1899, 12, 30)
            if secs < 86400:
                raise _Cannot(f"the number {_number_text(new[1])} is before the first "
                              "date of the 1900 date system (1 January 1900 is 1)")
            if secs < 60 * 86400:
                secs += 86400
            elif secs < 61 * 86400:
                raise _Cannot(f"the number {_number_text(new[1])} stands for "
                              "29 February 1900, a day that never existed (Excel "
                              "numbers it only to stay compatible with older programs)")
        try:
            when = base + _dt.timedelta(seconds=secs)
            when -= _utc_offset(when)
        except OverflowError as exc:
            raise _Cannot("the number is not a date Excel can show") from exc
        return when.strftime("%Y-%m-%dT%H:%M:%SZ")
    raise _Cannot(f"the property has a type ({tag or 'none'}) this server does not write")


def _what(new: tuple) -> str:
    kind = new[0]
    return {"empty": "nothing", "str": "text", "num": "a number", "bool": "TRUE/FALSE",
            "err": "an error value", "formula": "a formula"}[kind]


def _shown(sheet: str, coord: str) -> str:
    return f"{sheet}!{coord}"


# ------------------------------------------------------------------- sync

def sync(original: str, produced: str, *, allow_loss: bool) -> SyncResult:
    """Write each linked property's new value into `produced` (in place) when
    its cell changed from `original`; see the module docstring. Raises
    HazardRefused for a value a property cannot hold unless allow_loss."""
    result = SyncResult()
    with zipfile.ZipFile(produced) as zp:
        new_pkg = _s.Package(zp)
        part = _custom_part(new_pkg)
        if part is None:
            return result
        data = new_pkg.read(part)
        if not _may_hold_link(data):
            return result
        try:
            _root, found = _parse_properties(data)
        except _xmlsafe.XML_ERRORS as exc:
            result.warnings.append(
                "the document properties could not be read as XML "
                f"({exc}), so this server did not look for a property linked to a cell. "
                "If one is linked to a cell you edited, Excel may put its old value back "
                "when the workbook opens; check the properties in Excel (File, Info, "
                "Properties, Advanced Properties, Custom)")
            return result
        if not found:
            return result
        props = [f[0] for f in found]
        edits: list[tuple[_Prop, str, str]] = []
        refused: list[tuple[_Prop, str, str, str]] = []
        with zipfile.ZipFile(original) as zo:
            old_pkg = _s.Package(zo)
            try:
                old_names, new_names = _workbook_root(old_pkg), _workbook_root(new_pkg)
                date1904 = _is_1904(new_names)
                old_rd = _Reader(old_pkg, _sheet_parts(old_pkg, old_names))
                new_rd = _Reader(new_pkg, _sheet_parts(new_pkg, new_names))
            except _Unreadable as exc:
                old_names = new_names = None
                date1904 = False
                unreadable = str(exc)
            for prop in props:
                try:
                    if old_names is None:
                        raise _Unreadable(unreadable)
                    old_at = _name_cell(old_names, prop.target)
                    new_at = _name_cell(new_names, prop.target)
                    if old_at is None or new_at is None:
                        continue
                    old = old_rd.cell(*old_at)
                    new = new_rd.cell(*new_at)
                except (_Unreadable, KeyError, ValueError, zipfile.BadZipFile) as exc:
                    refused.append((prop, prop.target, "?",
                                    f"this server could not read the cell ({exc})"))
                    continue
                if not _changed(old, new):
                    continue
                shown = _shown(*new_at)
                try:
                    text = _property_text(prop.tag, new, date1904)
                except _Cannot as exc:
                    refused.append((prop, shown, _what(new), str(exc)))
                    continue
                if text != prop.text:
                    edits.append((prop, shown, text))
    for prop, shown, what, why in refused:
        result.warnings.append(
            f"the document property '{prop.name}' is linked to {shown} and was "
            f"left as it was ({why}). When the workbook opens, Excel puts the "
            "property's value back into a cell that holds a value; a cell that holds "
            "a formula keeps its result, and Excel updates the property when it next "
            "saves the file")
    if refused and not allow_loss:
        names = ", ".join(f"'{p.name}'" for p, *_ in refused)
        first = refused[0]
        exc = HazardRefused(
            f"The cell {first[1]} is linked to the document property {names} "
            "(a custom property that Excel keeps in step with the cell), and this "
            "edit cannot be written into the property. "
            f"{first[3][:1].upper() + first[3][1:]}. Nothing was changed. "
            "Two routes: pass allow_loss=true to save the edit and leave the "
            "property as it was (a cell that holds a value then shows the "
            "property's value again when Excel opens the file; a formula keeps "
            "its result), or change the property in Excel itself "
            "(File, Info, Properties, Advanced Properties, Custom).")
        exc.detail = {"properties": [p.name for p, *_ in refused],
                      "cells": sorted({r[1] for r in refused}),
                      "routes": ["allow_loss", "leave this workbook alone at the file tier"]}
        raise exc
    if not edits:
        return result
    out = _carry._Out(produced)
    real = out.real(part)
    blob = out.data[real]
    root, fresh = _parse_properties(blob)
    by_index = {f[0].index: f for f in fresh}
    for prop, shown, text in edits:
        cur = by_index.get(prop.index)
        if cur is None or cur[0].name != prop.name or cur[2] is None:
            continue
        cur[2].text = text
        result.changed.append({"property": prop.name, "cell": shown, "value": text})
    if not result.changed:
        return result
    try:
        out.put(real, _serialize(blob, root))
    except (UnicodeError, LookupError, ValueError) as exc:
        raise HazardRefused(
            f"The document properties part could not be written back in its own "
            f"encoding ({exc}), so the property linked to {edits[0][1]} could not be "
            "updated. Nothing was changed. Make this edit in Excel, which updates the "
            "property along with the cell.") from exc
    out.write(produced)
    return result
