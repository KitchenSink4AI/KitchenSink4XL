"""core/rawlinks.py: cell hyperlinks on the byte-preserving route.

A hyperlink is not in its cell either: the sheet's <hyperlinks> block holds
one <hyperlink ref=...> per link, an in-workbook link names its target in a
location attribute, and a web or file link points at an external
relationship of the sheet. Adding, changing or removing a link is a splice
into that block (and the sheet's relationship part); everything else in the
sheet and the package stays as it was.

Anything this module does not recognise (a link spanning a range, a
<hyperlinks> block it cannot find the end of) raises surgical.NotSurgical and
the edit takes the full rewrite.
"""

from __future__ import annotations

import re

from . import surgical as _s

REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

#: Worksheet children that come after <hyperlinks> in the schema order.
_AFTER_HYPERLINKS = (
    "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks",
    "colBreaks", "customProperties", "cellWatches", "ignoredErrors",
    "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF",
    "picture", "oleObjects", "AlternateContent", "controls",
    "webPublishItems", "tableParts", "extLst")

_LINK_RE = re.compile(rb"<((?:[A-Za-z_][\w.\-]*:)?)hyperlink\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)/>")


def apply(pkg: _s.Package, sheet_part: str, sheet_data: bytes, links: dict,
          out_parts: dict) -> bytes:
    """Apply hyperlink changes {(row, col): Hyperlink or None} to one sheet's
    bytes (after its cell splice) and its relationship part."""
    from . import rawnotes as _rawnotes

    scan = _s.scan_sheet(sheet_data, [])
    prefix = scan.prefix
    p = prefix.decode("ascii")
    sd_end = scan.sd_close if scan.sd_selfclose else sheet_data.index(b">", scan.sd_close) + 1
    data = sheet_data
    r_pfx = b"r"
    if any(hl is not None and hl.target for hl in links.values()):
        # a web or file link needs the root's binding for the relationships
        # namespace, declared when it has none and never over a prefix the
        # root already uses; a removal or an in-workbook link needs
        # none, so the root stays as it was
        data, r_pfx, shift = _s.worksheet_rel_prefix(data)
        sd_end += shift
    # existing links, by cell
    block = re.search(rb"<" + re.escape(prefix) + rb"hyperlinks(?=[\s/>])[^>]*>", data[sd_end:])
    existing: dict = {}
    if block is not None:
        b_start = sd_end + block.start()
        b_open_end = sd_end + block.end()
        if data[b_open_end - 2:b_open_end] == b"/>":
            b_close = None
        else:
            b_close = data.find(b"</" + prefix + b"hyperlinks>", b_open_end)
            if b_close < 0:
                raise _s.NotSurgical("a worksheet's hyperlink list is not closed")
            for m in _LINK_RE.finditer(data, b_open_end, b_close):
                a = _s._attrs(m.group(2))
                ref = a.get("ref", b"")
                rc = _s._split_coord(ref)
                if rc is None:
                    if b":" in ref:
                        # a link over a range: only the rewrite models those
                        for (r, c) in links:
                            box = _s._ref_box(ref.decode("ascii", "replace"))
                            if box and box[0] <= r <= box[2] and box[1] <= c <= box[3]:
                                raise _s.NotSurgical("this edit changes a hyperlink that spans a range")
                    continue
                existing[rc] = (m, a)
    # relationship part of the sheet
    rp_name = _s.rels_path(sheet_part)
    rp = pkg.real(rp_name)
    if rp is not None:
        rdata = out_parts.get(rp, pkg.read(rp))
    else:
        rp = rp_name
        rdata = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                 b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    rels = _s.parse_rels(rdata)
    used = {r.get("Id", "") for r in rels}
    rel_edits: list = []
    new_rels: list = []

    def new_rid() -> str:
        n = 1
        while f"rId{n}" in used:
            n += 1
        used.add(f"rId{n}")
        return f"rId{n}"

    def drop_rel(rid: str) -> None:
        for m in _s._REL_RE.finditer(rdata):
            if _s._attrs(m.group(1)).get("Id", b"").decode("utf-8") == rid:
                rel_edits.append((m.start(), m.end(), b""))

    edits: list = []
    adds: list = []
    for (r, c), hl in sorted(links.items()):
        old = existing.get((r, c))
        if old is not None:
            m, a = old
            rid = next((v.decode("utf-8") for k, v in a.items() if k.endswith(":id")), None)
            if rid:
                drop_rel(rid)
            edits.append((m.start(), m.end(), b""))
        if hl is None:
            continue
        coord = f"{_s._col_letters(c)}{r}"
        attrs = [f'ref="{coord}"']
        if hl.target:
            rid = new_rid()
            new_rels.append(f'<Relationship Id="{rid}" Type="{REL_HYPERLINK}" '
                            f'Target="{_s.xml_attr(hl.target)}" TargetMode="External"/>')
            attrs.append(f'{r_pfx.decode()}:id="{rid}"')
        if hl.location:
            attrs.append(f'location="{_s.xml_attr(hl.location)}"')
        if hl.display:
            attrs.append(f'display="{_s.xml_attr(hl.display)}"')
        if hl.tooltip:
            attrs.append(f'tooltip="{_s.xml_attr(hl.tooltip)}"')
        adds.append(f"<{p}hyperlink {' '.join(attrs)}/>".encode("utf-8"))
    remaining = len(existing) - sum(1 for k in links if k in existing) + len(adds)
    payload = b"".join(adds)
    if block is None:
        if payload:
            at = _rawnotes._top_level_insert_point(data, sd_end, _AFTER_HYPERLINKS)
            edits.append((at, at, f"<{p}hyperlinks>".encode() + payload + f"</{p}hyperlinks>".encode()))
    elif b_close is None:
        if payload:
            edits.append((b_start, b_open_end, f"<{p}hyperlinks>".encode() + payload
                          + f"</{p}hyperlinks>".encode()))
    elif remaining == 0:
        # the last link went: the empty block goes too (Excel never writes one)
        end = b_close + len(b"</" + prefix + b"hyperlinks>")
        edits = [e for e in edits if not (b_start <= e[0] and e[1] <= end)]
        edits.append((b_start, end, b""))
    elif payload:
        edits.append((b_close, b_close, payload))
    if new_rels:
        close = rdata.rfind(b"</Relationships>")
        if close < 0:
            raise _s.NotSurgical("a sheet's relationship part is unreadable")
        rel_edits.append((close, close, "".join(new_rels).encode("utf-8")))
    if rel_edits:
        out_parts[rp] = _s._apply_edits(rdata, rel_edits)
    return _s._apply_edits(data, edits) if edits else data


__all__ = ["apply"]
