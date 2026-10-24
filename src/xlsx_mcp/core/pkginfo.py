"""core/pkginfo.py: what is in the package, read from the package itself.

The readers used to answer from openpyxl's model, and the model is blind to
much of an ordinary workbook: pictures need Pillow before openpyxl will load
them at all (manage_image listed none of four logos), dropdowns that
point at another sheet live in an x14 extension block openpyxl discards, an
Excel 4.0 macro sheet loads as a plain worksheet, and the model has no notion
of whether a cached result is current. These
functions read the package parts directly, need nothing beyond the standard
library, and never change the file.
"""

from __future__ import annotations

import html
import re
import zipfile

from . import surgical as _s

_EMU_PER_PX = 9525
_ANCHOR_RE = re.compile(
    rb"<((?:[A-Za-z_][\w.\-]*:)?)(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b"
    rb"((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>(.*?)</\1\2>", re.S)
_KINDS = ((b"pic", "picture"), (b"graphicFrame", "chart or other graphic"),
          (b"sp", "shape"), (b"grpSp", "group"), (b"cxnSp", "connector"),
          (b"contentPart", "ink"))


def _first_child_kind(body: bytes) -> str:
    """The object an anchor holds, by its first object element."""
    best = None
    for tag, kind in _KINDS:
        m = re.search(rb"<(?:[A-Za-z_][\w.\-]*:)?" + tag + rb"(?=[\s/>])", body)
        if m is not None and (best is None or m.start() < best[0]):
            best = (m.start(), kind)
    return best[1] if best else "unknown"


def _cell_of(block: bytes | None) -> str | None:
    if block is None:
        return None
    c = re.search(rb"<(?:[A-Za-z_][\w.\-]*:)?col>\s*(\d+)\s*<", block)
    r = re.search(rb"<(?:[A-Za-z_][\w.\-]*:)?row>\s*(\d+)\s*<", block)
    if c is None or r is None:
        return None
    return f"{_s._col_letters(int(c.group(1)) + 1)}{int(r.group(1)) + 1}"


def _ext_px(body: bytes) -> tuple[int | None, int | None]:
    m = re.search(rb"<(?:[A-Za-z_][\w.\-]*:)?ext\s+cx=\"(\d+)\"\s+cy=\"(\d+)\"", body)
    if m is None:
        return None, None
    return round(int(m.group(1)) / _EMU_PER_PX), round(int(m.group(2)) / _EMU_PER_PX)


def drawing_objects(path: str, sheet: str | None = None) -> dict:
    """{sheet title: [object]} for every object anchored in each sheet's
    drawing, in drawing order. A picture carries its name, alt text, anchor
    cells, placed size in pixels and the media part it shows."""
    out: dict = {}
    with zipfile.ZipFile(path) as zf:
        pkg = _s.Package(zf)
        for title, part in pkg.sheet_parts().items():
            if sheet is not None and title != sheet:
                continue
            objs: list = []
            for dp in pkg.related(part, "/drawing"):
                if not pkg.has(dp):
                    continue
                data = pkg.read(dp)
                rels = {r.get("Id"): r for r in pkg.rels(dp)}
                for m in _ANCHOR_RE.finditer(data):
                    kind_tag = m.group(2).decode()
                    attrs = _s._attrs(m.group(3))
                    body = m.group(4)
                    kind = _first_child_kind(body)
                    name = re.search(rb"<(?:[A-Za-z_][\w.\-]*:)?cNvPr\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>", body)
                    na = _s._attrs(name.group(1)) if name else {}
                    fr = re.search(rb"<(?:[A-Za-z_][\w.\-]*:)?from>(.*?)</(?:[A-Za-z_][\w.\-]*:)?from>", body, re.S)
                    to = re.search(rb"<(?:[A-Za-z_][\w.\-]*:)?to>(.*?)</(?:[A-Za-z_][\w.\-]*:)?to>", body, re.S)
                    w, h = _ext_px(body)
                    obj = {
                        "kind": kind,
                        "name": html.unescape(na.get("name", b"").decode("utf-8")) or None,
                        "anchor": {
                            "type": {"twoCellAnchor": "two-cell", "oneCellAnchor": "one-cell",
                                     "absoluteAnchor": "absolute"}[kind_tag],
                            "from": _cell_of(fr.group(1) if fr else None),
                            "to": _cell_of(to.group(1) if to else None),
                        },
                        "width": w or None, "height": h or None,
                    }
                    edit_as = attrs.get("editAs")
                    if edit_as:
                        obj["anchor"]["moves_and_sizes"] = {
                            b"oneCell": "moves with cells, does not resize",
                            b"twoCell": "moves and resizes with cells",
                            b"absolute": "fixed on the page"}.get(edit_as, edit_as.decode())
                    descr = na.get("descr")
                    if descr:
                        obj["description"] = html.unescape(descr.decode("utf-8"))
                    if kind == "picture":
                        emb = re.search(rb"<(?:[A-Za-z_][\w.\-]*:)?blip\b[^>]*?:(?:embed|link)=\"([^\"]+)\"", body)
                        if emb is not None:
                            rel = rels.get(emb.group(1).decode())
                            if rel is not None:
                                if rel.get("TargetMode") == "External":
                                    obj["linked"] = rel.get("Target")
                                else:
                                    mp = _s.resolve_target(dp, rel.get("Target", ""))
                                    obj["media"] = mp
                                    obj["format"] = mp.rsplit(".", 1)[-1].lower() if "." in mp else None
                                    real = pkg.real(mp)
                                    if real is not None:
                                        obj["bytes"] = next(i.file_size for i in pkg.infos if i.filename == real)
                    objs.append(obj)
            out[title] = objs
    return out


def sheet_kinds(path: str) -> dict:
    """{sheet title: kind} where kind is worksheet, chartsheet, macrosheet
    (Excel 4.0 macros) or dialogsheet (Excel 5 dialogs)."""
    out: dict = {}
    with zipfile.ZipFile(path) as zf:
        pkg = _s.Package(zf)
        wbp = pkg.workbook_part()
        types = {r.get("Id"): r.get("Type", "") for r in pkg.rels(wbp)}
        data = pkg.read(wbp)
        for m in re.finditer(rb"<(?:[\w.\-]+:)?sheet\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>", data):
            a = _s._attrs(m.group(1))
            name = a.get("name")
            rid = next((v for k, v in a.items() if k.endswith(":id")), None)
            if name is None or rid is None:
                continue
            t = types.get(rid.decode("utf-8"), "")
            kind = ("macrosheet" if t.endswith("/xlMacrosheet") or t.endswith("/xlIntlMacrosheet")
                    else "dialogsheet" if t.endswith("/dialogsheet")
                    else "chartsheet" if t.endswith("/chartsheet") else "worksheet")
            out[html.unescape(name.decode("utf-8"))] = kind
    return out


def builtin_names(path: str) -> list[tuple[str, str, str, bool]]:
    """(scope, name, refers_to, hidden) for each of Excel's built-in names in
    the file (_xlnm.Print_Area, _xlnm.Print_Titles, _xlnm._FilterDatabase and
    the like), as stored. openpyxl folds them into sheet settings on load, so
    a read of the loaded model never sees them as names."""
    out: list = []
    with zipfile.ZipFile(path) as zf:
        pkg = _s.Package(zf)
        data = pkg.read(pkg.workbook_part())
    titles = [html.unescape(_s._attrs(m.group(1)).get("name", b"").decode(
        "utf-8")) for m in re.finditer(
        rb"<(?:[\w.\-]+:)?sheet\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>", data)]
    for m in re.finditer(rb"<(?:[\w.\-]+:)?definedName\b((?:[^>\"']|\"[^\"]*\""
                         rb"|'[^']*')*)>(.*?)</(?:[\w.\-]+:)?definedName>",
                         data, re.S):
        a = _s._attrs(m.group(1))
        name = html.unescape(a.get("name", b"").decode("utf-8"))
        if not name.startswith("_xlnm."):
            continue
        sid = a.get("localSheetId")
        scope = "workbook"
        if sid is not None and sid.isdigit() and int(sid) < len(titles):
            scope = titles[int(sid)]
        out.append((scope, name, html.unescape(m.group(2).decode("utf-8")),
                    a.get("hidden") in (b"1", b"true")))
    return out


def x14_only_cf_rules(path: str) -> dict:
    """{sheet title: count} of conditional-format rules stored only in
    Excel's x14 extension block (a rule that refers to another sheet, an
    icon set with custom icons, and the like). A data bar's x14 companion
    of a rule in the main section is not counted: it is that rule."""
    out: dict = {}
    with zipfile.ZipFile(path) as zf:
        pkg = _s.Package(zf)
        for title, part in pkg.sheet_parts().items():
            if not pkg.has(part):
                continue
            data = pkg.read(part)
            if b"cfRule" not in data:
                continue
            refs = set(re.findall(rb"<(?:[\w.\-]+:)?id>\s*([^<\s]+)\s*<", data))
            n = 0
            for m in re.finditer(rb"<x14:cfRule\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>",
                                 data):
                rid = _s._attrs(m.group(1)).get("id")
                if rid is None or rid not in refs:
                    n += 1
            if n:
                out[title] = n
    return out


def xlm_macro_names(path: str) -> list[str]:
    """Defined names that are Excel 4.0 macros (xlm="1"), such as Auto_Open."""
    with zipfile.ZipFile(path) as zf:
        pkg = _s.Package(zf)
        data = pkg.read(pkg.workbook_part())
    out = []
    for m in re.finditer(rb"<(?:[\w.\-]+:)?definedName\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>", data):
        a = _s._attrs(m.group(1))
        if a.get("xlm") in (b"1", b"true") and a.get("name"):
            out.append(html.unescape(a["name"].decode("utf-8")))
    return out


_X14_DV_RE = re.compile(
    rb"<((?:[A-Za-z_][\w.\-]*:)?)dataValidation\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>(.*?)</\1dataValidation>",
    re.S)


def x14_validations(path: str, sheet: str | None = None) -> dict:
    """{sheet title: [validation]} for the data validations stored in a
    worksheet's x14 extension block (Excel writes a list that points at
    another sheet there). openpyxl discards these at load."""
    out: dict = {}
    with zipfile.ZipFile(path) as zf:
        pkg = _s.Package(zf)
        for title, part in pkg.sheet_parts().items():
            if sheet is not None and title != sheet:
                continue
            if not pkg.has(part):
                continue
            data = pkg.read(part)
            i = data.find(b"dataValidations")
            if i < 0:
                continue
            found = []
            for ext in re.finditer(rb"<(?:[A-Za-z_][\w.\-]*:)?ext\b[^>]*uri=\"\{CCE6A557-97BC-4[Bb]89-ADB6-D9C93CAAB3DF\}\"[^>]*>(.*?)</(?:[A-Za-z_][\w.\-]*:)?ext>",
                                   data, re.S):
                for m in _X14_DV_RE.finditer(ext.group(1)):
                    a = {k: html.unescape(v.decode("utf-8")) for k, v in _s._attrs(m.group(2)).items()}
                    body = m.group(3)
                    f1 = re.search(rb"formula1>\s*<(?:[A-Za-z_][\w.\-]*:)?f>(.*?)</", body, re.S)
                    f2 = re.search(rb"formula2>\s*<(?:[A-Za-z_][\w.\-]*:)?f>(.*?)</", body, re.S)
                    sq = re.search(rb"<(?:[A-Za-z_][\w.\-]*:)?sqref>(.*?)</", body, re.S)
                    found.append({
                        "range": html.unescape(sq.group(1).decode("utf-8")) if sq else None,
                        "type": a.get("type", "none"),
                        "operator": a.get("operator"),
                        "formula1": html.unescape(f1.group(1).decode("utf-8")) if f1 else None,
                        "formula2": html.unescape(f2.group(1).decode("utf-8")) if f2 else None,
                        "allow_blank": a.get("allowBlank") in ("1", "true"),
                        "error_message": a.get("error"),
                        "prompt": a.get("prompt"),
                    })
            if found:
                out[title] = found
    return out


def calc_state(path: str) -> dict:
    """What the package says about its cached results: the recalculate-on-
    open flag, the calculation mode, the engine id, and the application that
    last wrote the file."""
    with zipfile.ZipFile(path) as zf:
        pkg = _s.Package(zf)
        data = pkg.read(pkg.workbook_part())
        app = None
        for r in pkg.rels(""):
            if r.get("Type", "").endswith("/extended-properties"):
                p = _s.resolve_target("", r.get("Target", ""))
                if pkg.has(p):
                    m = re.search(rb"<(?:[\w.\-]+:)?Application>(.*?)</", pkg.read(p), re.S)
                    if m:
                        app = html.unescape(m.group(1).decode("utf-8", "replace")).strip()
    m = re.search(rb"<(?:[\w.\-]+:)?calcPr\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>", data)
    a = _s._attrs(m.group(1)) if m else {}
    return {
        "recalculates_on_open": a.get("fullCalcOnLoad") in (b"1", b"true"),
        "calc_mode": (a.get("calcMode") or b"auto").decode(),
        "calc_engine_id": a.get("calcId", b"").decode() or None,
        "last_saved_by": app,
    }


def _note_count(data: bytes) -> int:
    """Legacy notes in a comments part, leaving out the fallback copies
    Excel writes for threaded comments (their author is "tc={id}")."""
    authors = [html.unescape(m.group(1).decode("utf-8", "replace"))
               for m in re.finditer(rb"<(?:[\w.\-]+:)?author>(.*?)</", data, re.S)]
    n = 0
    for m in re.finditer(rb"<(?:[\w.\-]+:)?comment\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>", data):
        a = _s._attrs(m.group(1))
        try:
            who = authors[int(a.get("authorId", b"0"))]
        except (ValueError, IndexError):
            who = ""
        if not who.startswith("tc="):
            n += 1
    return n


def hidden_objects(path: str) -> dict:
    """{sheet title: {kind: count}} for what a cell grid does not show:
    pictures, charts, shapes, notes, threaded comments, x14 validations,
    sparkline groups, slicers. Only nonzero kinds, only sheets with any."""
    out: dict = {}
    draw = drawing_objects(path)
    x14 = x14_validations(path)
    with zipfile.ZipFile(path) as zf:
        pkg = _s.Package(zf)
        for title, part in pkg.sheet_parts().items():
            counts: dict = {}
            for o in draw.get(title, []):
                key = {"picture": "pictures", "chart or other graphic": "charts_or_graphics",
                       "shape": "shapes", "group": "shape_groups",
                       "connector": "connectors"}.get(o["kind"], "other_objects")
                counts[key] = counts.get(key, 0) + 1
            for cp in pkg.related(part, "/comments"):
                if pkg.has(cp):
                    counts["notes"] = _note_count(pkg.read(cp))
            for tp in pkg.related(part, "/threadedComment"):
                if pkg.has(tp):
                    counts["threaded_comments"] = len(re.findall(rb"<(?:[\w.\-]+:)?threadedComment\b", pkg.read(tp)))
            if title in x14:
                counts["x14_validations"] = len(x14[title])
            if pkg.has(part):
                data = pkg.read(part)
                n = len(re.findall(rb"<(?:[\w.\-]+:)?sparklineGroup\b", data))
                if n:
                    counts["sparkline_groups"] = n
                n = len(re.findall(rb"<(?:[\w.\-]+:)?slicer\s+r:id", data))
                if n:
                    counts["slicer_lists"] = n
            counts = {k: v for k, v in counts.items() if v}
            if counts:
                out[title] = counts
    return out


__all__ = ["drawing_objects", "sheet_kinds", "xlm_macro_names", "builtin_names",
           "x14_only_cf_rules",
           "x14_validations", "calc_state", "hidden_objects"]
