"""core/dynarray.py: formulas written as dynamic arrays.

Excel 365 stores a formula that can return an array (=UNIQUE(A1:A5),
=SEQUENCE(3)) as an array formula on its anchor cell that points, through the
cell's cm attribute, at the XLDAPR "dynamic array" entry of xl/metadata.xml:

    <c r="C1" cm="1"><f t="array" ref="C1">_xlfn.UNIQUE(A1:A5)</f></c>

Written as a plain <f>, the same text is a LEGACY formula: Excel opens it as
=@UNIQUE(A1:A5), shows one value and never spills, with no warning anywhere.
Both save routes wrote it that way. mark_dynamic() runs on the produced
package, before verify-after-write sees it, and turns every formula cell this
save wrote that uses an array-returning function into the dynamic form,
creating xl/metadata.xml (or reusing its XLDAPR entry) as Excel does.
The spill range itself is left to Excel: the workbook is already flagged to
recalculate on open, and Excel sizes the spill when it calculates.
"""

from __future__ import annotations

import os
import re
import shutil
import zipfile

from . import surgical as _s

#: Functions whose result is (or can be) an array that spills. A formula
#: that uses one anywhere is written as a dynamic array; for a formula whose
#: result is a single value (=SUM(UNIQUE(A1:A5))) that is exactly what Excel
#: 365 itself writes, and it reads the same.
SPILLING_FUNCS = frozenset({
    "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY",
    "TOCOL", "TOROW", "HSTACK", "VSTACK", "WRAPROWS", "WRAPCOLS", "TAKE",
    "DROP", "CHOOSECOLS", "CHOOSEROWS", "EXPAND", "TEXTSPLIT", "MAKEARRAY",
    "MAP", "SCAN", "BYROW", "BYCOL", "GROUPBY", "PIVOTBY", "TRIMRANGE",
    "REGEXEXTRACT",
})

_STRING_RE = re.compile(r'"(?:[^"]|"")*"')
_CALL_RE = re.compile(r"(?<![A-Za-z0-9_.])((?:_xlfn\.|_xlws\.)*)([A-Za-z][A-Za-z0-9_.]*)\s*\(")

REL_METADATA = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                "relationships/sheetMetadata")
CT_METADATA = ("application/vnd.openxmlformats-officedocument.spreadsheetml."
               "sheetMetadata+xml")
_XLDAPR_EXT = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_XDA = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"

NEW_METADATA = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<metadata xmlns="{_NS_MAIN}" xmlns:xda="{_NS_XDA}">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" '
    'minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" '
    'merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
    'clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    f'<ext uri="{_XLDAPR_EXT}"><xda:dynamicArrayProperties fDynamic="1" '
    'fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>'
    '</metadata>').encode("utf-8")


class NotDynamic(Exception):
    """The package is not in a shape this step can mark safely."""


def is_spilling(formula: str) -> bool:
    """True when the formula calls an array-returning function."""
    if not isinstance(formula, str):
        return False
    body = _STRING_RE.sub('""', formula)
    return any(m.group(2).upper() in SPILLING_FUNCS
               for m in _CALL_RE.finditer(body))


def _xldapr_cm(meta: bytes) -> int | None:
    """The 1-based cellMetadata index of an existing dynamic-array entry,
    or None when the part has none (it may hold rich-value metadata only)."""
    types = re.findall(rb"<(?:\w+:)?metadataType\b([^>]*)>", meta)
    t_index = None
    for i, attrs in enumerate(types, start=1):
        if re.search(rb'\bname="XLDAPR"', attrs):
            t_index = i
            break
    if t_index is None:
        return None
    cm = re.search(rb"<(?:\w+:)?cellMetadata\b[^>]*>(.*?)</(?:\w+:)?cellMetadata>",
                   meta, re.S)
    if cm is None:
        return None
    for i, bk in enumerate(re.findall(rb"<(?:\w+:)?bk\b[^>]*>(.*?)</(?:\w+:)?bk>",
                                      cm.group(1), re.S), start=1):
        if re.search(rb'<(?:\w+:)?rc\b[^>]*\bt="%d"' % t_index, bk):
            return i
    return None


def _mark_cell(sheet: bytes, coord: str, cm: int) -> bytes:
    pat = re.compile(
        rb'<((?:[A-Za-z_][\w.\-]*:)?)c\b((?:[^>"\']|"[^"]*"|\'[^\']*\')*?\br="'
        + re.escape(coord.encode()) + rb'"(?:[^>"\']|"[^"]*"|\'[^\']*\')*)>(.*?)</\1c>',
        re.S)
    m = pat.search(sheet)
    if m is None:
        raise NotDynamic(f"cell {coord} was not found in the written sheet")
    prefix, attrs, inner = m.group(1), m.group(2), m.group(3)
    fm = re.search(rb"<" + re.escape(prefix) + rb"f\b([^>]*)>", inner)
    if fm is None or fm.group(0).endswith(b"/>"):
        raise NotDynamic(f"cell {coord} holds no formula text")
    f_attrs = fm.group(1)
    if b't="shared"' in f_attrs:
        raise NotDynamic(f"cell {coord} holds a shared formula")
    if b't="array"' not in f_attrs:
        new_f = (b"<" + prefix + b'f t="array" ref="' + coord.encode() + b'"'
                 + f_attrs + b">")
        inner = inner[:fm.start()] + new_f + inner[fm.end():]
    attrs = re.sub(rb'\scm="\d+"', b"", attrs)
    attrs = attrs + b' cm="%d"' % cm
    return (sheet[:m.start()] + b"<" + prefix + b"c" + attrs + b">" + inner
            + b"</" + prefix + b"c>" + sheet[m.end():])


def mark_dynamic(path: str, cells: dict[str, set[str]]) -> int:
    """Turn the formula cells `cells` ({sheet title: {"A1", ...}}) of the
    package at `path` into dynamic-array formulas, in place. Returns the
    number of cells marked. Raises NotDynamic when the package cannot be
    marked safely (the caller refuses the write rather than save a formula
    Excel would show as a single @ value)."""
    cells = {s: set(c) for s, c in cells.items() if c}
    if not cells:
        return 0
    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}
        pkg = _s.Package(zin)
        wbp = pkg.workbook_part()
        sheet_parts = pkg.sheet_parts()
        meta_rels = [r for r in pkg.rels(wbp)
                     if r.get("Type", "").endswith("/sheetMetadata")]
    order = [i.filename for i in infos]
    lower = {n.lower(): n for n in data}

    if meta_rels:
        meta_part = lower.get(_s.resolve_target(wbp, meta_rels[0].get("Target", "")).lower())
        if meta_part is None:
            raise NotDynamic("the workbook names a metadata part it does not hold")
        cm = _xldapr_cm(data[meta_part])
        if cm is None:
            raise NotDynamic(
                "the workbook's metadata part holds no dynamic-array entry")
    else:
        meta_part = posixpath_join(os.path.dirname(wbp), "metadata.xml")
        if meta_part.lower() in lower:
            raise NotDynamic("a metadata part exists without a relationship")
        data[meta_part] = NEW_METADATA
        order.append(meta_part)
        rp = _s.rels_path(wbp)
        rels = data.get(lower.get(rp.lower(), rp))
        if rels is None:
            raise NotDynamic("the workbook has no relationship part")
        rp = lower.get(rp.lower(), rp)
        used = {r.get("Id", "") for r in _s.parse_rels(rels)}
        n = 1
        while f"rId{n}" in used:
            n += 1
        el = (f'<Relationship Id="rId{n}" Type="{REL_METADATA}" '
              f'Target="metadata.xml"/>').encode()
        close = rels.rfind(b"</Relationships>")
        if close < 0:
            raise NotDynamic("the workbook relationship part is unreadable")
        data[rp] = rels[:close] + el + rels[close:]
        ct = data["[Content_Types].xml"]
        el = f'<Override PartName="/{meta_part}" ContentType="{CT_METADATA}"/>'.encode()
        close = ct.rfind(b"</Types>")
        data["[Content_Types].xml"] = ct[:close] + el + ct[close:]
        cm = 1

    marked = 0
    for title, coords in cells.items():
        part = sheet_parts.get(title)
        real = lower.get(part.lower()) if part else None
        if real is None:
            raise NotDynamic(f"sheet {title!r} was not found in the written package")
        sheet = data[real]
        for coord in sorted(coords):
            sheet = _mark_cell(sheet, coord.upper().replace("$", ""), cm)
            marked += 1
        data[real] = sheet

    by_name = {i.filename: i for i in infos}
    from .safesave import scratch_file
    fd, tmp = scratch_file(os.path.dirname(os.path.abspath(path)),
                           os.path.splitext(path)[1])
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
            for name in order:
                info = by_name.get(name)
                if info is None:
                    info = zipfile.ZipInfo(name)
                    info.compress_type = zipfile.ZIP_DEFLATED
                zw.writestr(info, data[name])
        shutil.move(tmp, path)
    except BaseException:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    return marked


def posixpath_join(a: str, b: str) -> str:
    return f"{a}/{b}" if a else b
