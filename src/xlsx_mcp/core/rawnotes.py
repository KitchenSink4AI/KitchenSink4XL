"""core/rawnotes.py: legacy comments (notes) on the byte-preserving route.

A note lives in three places, none of them the cell: the sheet's comments
part (the text and author), the sheet's VML drawing (the yellow box, its
position and size), and the sheet's <legacyDrawing> element plus two
relationships that tie them together. Adding, editing or deleting a note is
therefore a splice into those parts, and every other byte of them is kept.

Cases handled:
  - edit a note's text or author: the <text> of its <comment> (and authorId);
    the box stays where it was.
  - add a note to a sheet that already has notes: a <comment> and a VML
    shape for it.
  - add the first note to a sheet with no comments part and no legacy
    drawing: a new comments part, a new VML part, the two relationships, the
    <legacyDrawing> element and the content types.
  - delete a note while others remain on the sheet.

Everything else raises surgical.NotSurgical and the edit takes the full
rewrite: deleting a sheet's last note (the parts would have to go too), a
note on a cell that also holds a threaded comment (the note is Excel's
fallback copy of that thread), a sheet whose existing VML drawing also holds
form controls, and any part whose shape this module does not recognize.
"""

from __future__ import annotations

import posixpath
import re

from . import surgical as _s

REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_V = b"urn:schemas-microsoft-com:vml"
NS_O = b"urn:schemas-microsoft-com:office:office"
NS_X = b"urn:schemas-microsoft-com:office:excel"

_ANY_TAG = re.compile(rb"<(/?)((?:[A-Za-z_][\w.\-]*:)?)([A-Za-z_][\w.\-]*)((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>")
_COMMENT_RE = re.compile(
    rb"<((?:[A-Za-z_][\w.\-]*:)?)comment(?=[\s/>])((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>(.*?)</\1comment>", re.S)
_AUTHOR_RE = re.compile(rb"<((?:[A-Za-z_][\w.\-]*:)?)author>(.*?)</\1author>|<(?:[A-Za-z_][\w.\-]*:)?author/>", re.S)
_TEXT_RE = re.compile(rb"<((?:[A-Za-z_][\w.\-]*:)?)text(?=[\s/>])(?:(?:[^>\"']|\"[^\"]*\"|'[^']*')*)>.*?</\1text>", re.S)
_SHAPE_RE = re.compile(rb"<v:shape(?=[\s>])(?:(?:[^>\"']|\"[^\"]*\"|'[^']*')*)>.*?</v:shape>", re.S)


def _xml_doc_prefix(data: bytes, local: bytes) -> bytes:
    m = re.search(rb"<((?:[A-Za-z_][\w.\-]*:)?)" + local + rb"(?=[\s/>])", data)
    if m is None:
        raise _s.NotSurgical("a comments part is unreadable")
    return m.group(1)


def _text_el(prefix: bytes, text: str) -> bytes:
    p = prefix.decode("ascii")
    space = ' xml:space="preserve"' if text != text.strip() or "\n" in text else ""
    return f"<{p}text><{p}t{space}>{_s.xml_text(text)}</{p}t></{p}text>".encode("utf-8")


def _vml_prefixes_ok(vml: bytes) -> bool:
    head = vml[:2048]
    return (b'xmlns:v="' + NS_V + b'"' in head and b'xmlns:o="' + NS_O + b'"' in head
            and b'xmlns:x="' + NS_X + b'"' in head)


def _shape_cell(shape: bytes) -> tuple[int, int] | None:
    if b'ObjectType="Note"' not in shape:
        return None
    r = re.search(rb"<x:Row>\s*(\d+)\s*</x:Row>", shape)
    c = re.search(rb"<x:Column>\s*(\d+)\s*</x:Column>", shape)
    if r is None or c is None:
        return None
    return int(r.group(1)) + 1, int(c.group(1)) + 1


def _shape_xml(shape_id: int, row: int, col: int, z: int) -> bytes:
    """An Excel-style note box for a cell (1-based row and column)."""
    r0, c0 = row - 1, col - 1
    anchor = f"{c0 + 1}, 15, {max(r0 - 1, 0)}, 10, {c0 + 3}, 15, {r0 + 3}, 4"
    return (f'<v:shape id="_x0000_s{shape_id}" type="#_x0000_t202" '
            f"style='position:absolute;margin-left:59.25pt;margin-top:1.5pt;"
            f"width:108pt;height:59.25pt;z-index:{z};visibility:hidden' "
            f'fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/>'
            f'<v:shadow on="t" color="black" obscured="t"/>'
            f'<v:path o:connecttype="none"/><v:textbox style=\'mso-direction-alt:auto\'>'
            f"<div style='text-align:left'></div></v:textbox>"
            f'<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
            f"<x:Anchor>{anchor}</x:Anchor><x:AutoFill>False</x:AutoFill>"
            f"<x:Row>{r0}</x:Row><x:Column>{c0}</x:Column></x:ClientData>"
            f"</v:shape>").encode("ascii")


_SHAPETYPE = (b'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
              b'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
              b'<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>')


def _max_shape_ids(pkg: _s.Package, out_parts: dict) -> tuple[int, int]:
    """(largest idmap block, largest numeric shape id) over every VML part."""
    block = 0
    sid = 0
    for name in list(pkg.names) + [n for n in out_parts if n not in pkg.names]:
        if not name.lower().endswith(".vml"):
            continue
        data = out_parts.get(name, pkg.read(name))
        for m in re.finditer(rb'data="([\d,\s]+)"', data):
            for piece in m.group(1).split(b","):
                if piece.strip().isdigit():
                    block = max(block, int(piece))
        for m in re.finditer(rb'id="_x0000_s(\d+)"', data):
            sid = max(sid, int(m.group(1)))
    return block, sid


def _free_name(pkg: _s.Package, out_parts: dict, pattern: str) -> str:
    n = 1
    while True:
        name = pattern.format(n)
        if not pkg.has(name) and name not in out_parts:
            return name
        n += 1


def _top_level_insert_point(data: bytes, sd_close: int, before_locals: tuple) -> int:
    """Offset of the first top-level worksheet child named in before_locals
    after the cell table, else the worksheet's closing tag."""
    depth = 0
    for m in _ANY_TAG.finditer(data, sd_close):
        closing = m.group(1) == b"/"
        local = m.group(3).decode("ascii", "replace")
        selfclose = m.group(4).rstrip().endswith(b"/")
        if closing:
            if depth == 0:
                return m.start()  # </worksheet>
            depth -= 1
            continue
        if depth == 0 and local in before_locals:
            return m.start()
        if not selfclose:
            depth += 1
    raise _s.NotSurgical("a worksheet is not closed")


_AFTER_LEGACY = ("legacyDrawingHF", "drawingHF", "picture", "oleObjects", "AlternateContent",
                 "controls", "webPublishItems", "tableParts", "extLst")


def apply(pkg: _s.Package, sheet_part: str, sheet_data: bytes, notes: dict,
          out_parts: dict, ct_edits: list, removed: set) -> bytes:
    """Apply note changes {(row, col): (new Comment or None, old note key)} for
    one sheet. Returns the (possibly updated) worksheet bytes; writes the
    comments/VML/rels parts into out_parts and content-type additions into
    ct_edits as (kind, key, content type)."""
    rels = pkg.rels(sheet_part)
    c_rel = next((r for r in rels if r.get("Type", "").endswith("/comments")), None)
    v_rel = next((r for r in rels if r.get("Type", "").endswith("/vmlDrawing")), None)
    t_rel = next((r for r in rels if r.get("Type", "").endswith("/threadedComment")), None)
    threaded: set = set()
    if t_rel is not None:
        tp = _s.resolve_target(sheet_part, t_rel.get("Target", ""))
        if pkg.has(tp):
            for m in re.finditer(rb'\bref="([A-Za-z]+\d+)"', pkg.read(tp)):
                rc = _s._split_coord(m.group(1))
                if rc:
                    threaded.add(rc)
    for rc in notes:
        if rc in threaded:
            raise _s.NotSurgical(
                "this edit changes a note on a cell that holds a threaded comment")
    if c_rel is None:
        if v_rel is not None:
            raise _s.NotSurgical(
                "this edit adds a note to a sheet whose legacy drawing holds "
                "other objects")
        return _first_notes(pkg, sheet_part, sheet_data, notes, out_parts, ct_edits)
    if v_rel is None:
        raise _s.NotSurgical("a sheet's notes have no drawing part")
    cp = pkg.real(_s.resolve_target(sheet_part, c_rel.get("Target", "")))
    vp = pkg.real(_s.resolve_target(sheet_part, v_rel.get("Target", "")))
    if cp is None or vp is None:
        raise _s.NotSurgical("a sheet's notes parts are missing")
    cdata = pkg.read(cp)
    vdata = pkg.read(vp)
    if not _vml_prefixes_ok(vdata):
        raise _s.NotSurgical("a sheet's notes drawing uses an unfamiliar layout")
    prefix = _xml_doc_prefix(cdata, b"comments")
    # authors
    authors: list[str] = []
    for m in _AUTHOR_RE.finditer(cdata):
        authors.append(_s._xml_unescape(m.group(2) or b""))
    new_authors: list[str] = []

    def author_id(name: str) -> int:
        if name in authors:
            return authors.index(name)
        if name in new_authors:
            return len(authors) + new_authors.index(name)
        new_authors.append(name)
        return len(authors) + len(new_authors) - 1

    existing = {}
    for m in _COMMENT_RE.finditer(cdata):
        a = _s._attrs(m.group(2))
        rc = _s._split_coord(a.get("ref", b""))
        if rc is not None:
            existing[rc] = m
    shapes = {}
    for m in _SHAPE_RE.finditer(vdata):
        rc = _shape_cell(m.group(0))
        if rc is not None:
            shapes[rc] = m
    c_edits: list = []
    v_edits: list = []
    deletes = [rc for rc, (new, _had) in notes.items() if new is None]
    remaining = (set(existing) - set(deletes)) | {
        rc for rc, (new, _had) in notes.items() if new is not None}
    if deletes and not remaining:
        return _remove_notes(pkg, sheet_part, sheet_data, c_rel, v_rel, cp, vp,
                             vdata, out_parts, ct_edits, removed)
    block, max_sid = _max_shape_ids(pkg, out_parts)
    z = len(shapes)
    adds: list = []
    for rc, (new, _had) in sorted(notes.items()):
        m = existing.get(rc)
        if new is None:
            if m is None:
                continue
            c_edits.append((m.start(), m.end(), b""))
            sm = shapes.get(rc)
            if sm is not None:
                v_edits.append((sm.start(), sm.end(), b""))
            continue
        _s_text = new.text or ""
        old_text = _had[0] if _had else None
        if m is not None:
            whole = m.group(0)
            tm = _TEXT_RE.search(whole)
            if tm is None:
                raise _s.NotSurgical("a note's text is unreadable")
            if _s_text == old_text:
                rebuilt = whole  # same words: keep the note's own formatting
            else:
                rebuilt = whole[:tm.start()] + _text_el(prefix, _s_text) + whole[tm.end():]
            tag_end = whole.find(b">") + 1
            head = whole[:tag_end]
            head2 = _s._set_attr(head, b"authorId", str(author_id(new.author or "")).encode())
            rebuilt = head2 + rebuilt[tag_end:]
            c_edits.append((m.start(), m.end(), rebuilt))
            continue
        coord = f"{_s._col_letters(rc[1])}{rc[0]}"
        p = prefix.decode("ascii")
        el = (f'<{p}comment ref="{coord}" authorId="{author_id(new.author or "")}">'
              .encode("utf-8") + _text_el(prefix, _s_text) + f"</{p}comment>".encode())
        adds.append((rc, el))
        max_sid += 1
        z += 1
        v_edits.append((None, None, _shape_xml(max_sid, rc[0], rc[1], z)))
    # place the new comments in row order inside commentList
    if adds:
        order = sorted(existing.items())
        for rc, el in adds:
            after = next((mm for rr, mm in order if rr > rc), None)
            if after is not None:
                c_edits.append((after.start(), after.start(), el))
            else:
                close = cdata.rfind(b"</" + prefix + b"commentList>")
                if close < 0:
                    raise _s.NotSurgical("a comments part has no comment list")
                c_edits.append((close, close, el))
    if new_authors:
        close = cdata.find(b"</" + prefix + b"authors>")
        p = prefix.decode("ascii")
        payload = "".join(f"<{p}author>{_s.xml_text(a)}</{p}author>" for a in new_authors).encode("utf-8")
        if close >= 0:
            c_edits.append((close, close, payload))
        else:
            m = re.search(rb"<" + re.escape(prefix) + rb"authors\s*/>", cdata)
            if m is None:
                raise _s.NotSurgical("a comments part has no author list")
            c_edits.append((m.start(), m.end(), f"<{p}authors>".encode() + payload + f"</{p}authors>".encode()))
    new_shapes = b"".join(e[2] for e in v_edits if e[0] is None)
    v_real = [e for e in v_edits if e[0] is not None]
    if new_shapes:
        close = vdata.rfind(b"</xml>")
        if close < 0:
            raise _s.NotSurgical("a notes drawing is not closed")
        if b'id="_x0000_t202"' not in vdata:
            new_shapes = _SHAPETYPE + new_shapes
        v_real.append((close, close, new_shapes))
    if c_edits:
        out_parts[cp] = _s._apply_edits(cdata, c_edits)
    if v_real:
        out_parts[vp] = _s._apply_edits(vdata, v_real)
    return sheet_data


def _remove_notes(pkg: _s.Package, sheet_part: str, sheet_data: bytes,
                  c_rel: dict, v_rel: dict, cp: str, vp: str, vdata: bytes,
                  out_parts: dict, ct_edits: list, removed: set) -> bytes:
    """The sheet's last note goes: so do its comments part, its notes
    drawing (when that drawing holds nothing but notes), the two
    relationships and the <legacyDrawing> element, as Excel does."""
    for m in _SHAPE_RE.finditer(vdata):
        if _shape_cell(m.group(0)) is None:
            raise _s.NotSurgical(
                "this edit deletes a sheet's last note, and its drawing holds "
                "other objects")
    if re.search(rb"<v:(?:rect|oval|line|polyline|roundrect|image|group)\b", vdata):
        raise _s.NotSurgical(
            "this edit deletes a sheet's last note, and its drawing holds "
            "other objects")
    vid = v_rel.get("Id", "")
    m = re.search(rb"<((?:[A-Za-z_][\w.\-]*:)?)legacyDrawing\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)/>",
                  sheet_data)
    if m is None or not any(k.endswith(":id") and v.decode("utf-8") == vid
                            for k, v in _s._attrs(m.group(2)).items()):
        raise _s.NotSurgical("a sheet's notes drawing is not where it should be")
    data = sheet_data[:m.start()] + sheet_data[m.end():]
    rp = pkg.real(_s.rels_path(sheet_part))
    rdata = out_parts.get(rp, pkg.read(rp))
    edits = []
    for rm in _s._REL_RE.finditer(rdata):
        rid = _s._attrs(rm.group(1)).get("Id", b"").decode("utf-8")
        if rid in (vid, c_rel.get("Id", "")):
            edits.append((rm.start(), rm.end(), b""))
    out_parts[rp] = _s._apply_edits(rdata, edits)
    removed.update({cp, vp})
    ct_edits.append(("remove", "/" + cp, ""))
    return data


def _first_notes(pkg: _s.Package, sheet_part: str, sheet_data: bytes,
                 notes: dict, out_parts: dict, ct_edits: list) -> bytes:
    adds = {rc: new for rc, (new, _had) in notes.items() if new is not None}
    if not adds:
        return sheet_data
    cp = _free_name(pkg, out_parts, "xl/comments{}.xml")
    vp = _free_name(pkg, out_parts, "xl/drawings/vmlDrawing{}.vml")
    authors: list[str] = []
    comments = []
    for rc, new in sorted(adds.items()):
        name = new.author or ""
        if name not in authors:
            authors.append(name)
        coord = f"{_s._col_letters(rc[1])}{rc[0]}"
        comments.append(f'<comment ref="{coord}" authorId="{authors.index(name)}">'.encode("utf-8")
                        + _text_el(b"", new.text or "") + b"</comment>")
    cxml = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            b'<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>'
            + b"".join(f"<author>{_s.xml_text(a)}</author>".encode("utf-8") for a in authors)
            + b"</authors><commentList>" + b"".join(comments) + b"</commentList></comments>")
    block, max_sid = _max_shape_ids(pkg, out_parts)
    block += 1
    sid = block * 1024
    shapes = b""
    for z, rc in enumerate(sorted(adds), start=1):
        sid += 1
        shapes += _shape_xml(sid, rc[0], rc[1], z)
    vxml = (b'<xml xmlns:v="urn:schemas-microsoft-com:vml" '
            b'xmlns:o="urn:schemas-microsoft-com:office:office" '
            b'xmlns:x="urn:schemas-microsoft-com:office:excel">'
            + f'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="{block}"/></o:shapelayout>'.encode()
            + _SHAPETYPE + shapes + b"</xml>")
    out_parts[cp] = cxml
    out_parts[vp] = vxml
    # relationships from the sheet
    rp_name = _s.rels_path(sheet_part)
    rp = pkg.real(rp_name)
    if rp is not None:
        rdata = out_parts.get(rp, pkg.read(rp))
    else:
        rp = rp_name
        rdata = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                 b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    used = {r.get("Id", "") for r in _s.parse_rels(rdata)}

    def new_id() -> str:
        n = 1
        while f"rId{n}" in used:
            n += 1
        used.add(f"rId{n}")
        return f"rId{n}"

    vid, cid = new_id(), new_id()
    base = posixpath.dirname(sheet_part)
    vt = posixpath.relpath(vp, base)
    ctgt = posixpath.relpath(cp, base)
    payload = (f'<Relationship Id="{vid}" Type="{REL_VML}" Target="{vt}"/>'
               f'<Relationship Id="{cid}" Type="{REL_COMMENTS}" Target="{ctgt}"/>').encode()
    close = rdata.rfind(b"</Relationships>")
    if close < 0:
        raise _s.NotSurgical("a sheet's relationship part is unreadable")
    out_parts[rp] = rdata[:close] + payload + rdata[close:]
    # the <legacyDrawing> element
    scan = _s.scan_sheet(sheet_data, [])
    prefix = scan.prefix
    # scan.sd_close is where </sheetData> STARTS (or the end of a
    # self-closing <sheetData/>); the worksheet's later children begin after
    # that closing tag.
    sd_close = scan.sd_close if scan.sd_selfclose else \
        sheet_data.index(b">", scan.sd_close) + 1
    # the root's binding for the relationships namespace, declared when it
    # has none (never over a prefix the root already uses)
    data, r_pfx, shift = _s.worksheet_rel_prefix(sheet_data)
    sd_close += shift
    at = _top_level_insert_point(data, sd_close, _AFTER_LEGACY)
    el = b"<" + prefix + b"legacyDrawing " + r_pfx + b':id="' + vid.encode() + b'"/>'
    data = data[:at] + el + data[at:]
    ct_edits.append(("override", "/" + cp, CT_COMMENTS))
    ct_edits.append(("default", "vml", CT_VML))
    return data


def apply_content_types(pkg: _s.Package, out_parts: dict, ct_edits: list) -> None:
    if not ct_edits:
        return
    name = pkg.real("[Content_Types].xml")
    data = out_parts.get(name, pkg.read(name))
    payload = b""
    for kind, key, ctype in ct_edits:
        if kind == "remove":
            m = re.search(rb"<(?:[\w.\-]+:)?Override\b[^>]*?PartName=\""
                          + re.escape(key.encode()) + rb"\"[^>]*>", data, re.I)
            if m is not None:
                data = data[:m.start()] + data[m.end():]
                out_parts[name] = data
            continue
        if kind == "default":
            if re.search(rb'Extension="' + re.escape(key.encode()) + rb'"', data, re.I):
                continue
            payload += f'<Default Extension="{key}" ContentType="{ctype}"/>'.encode()
        else:
            if f'PartName="{key}"'.encode() in data:
                continue
            payload += f'<Override PartName="{_s.xml_attr(key)}" ContentType="{ctype}"/>'.encode()
    if not payload:
        return
    close = data.rfind(b"</Types>")
    if close < 0:
        raise _s.NotSurgical("the content types part is unreadable")
    out_parts[name] = data[:close] + payload + data[close:]


__all__ = ["apply", "apply_content_types"]
