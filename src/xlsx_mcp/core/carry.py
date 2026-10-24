"""core/carry.py: what the full rewrite puts back from the original.

The full openpyxl rewrite (structural edits: inserting rows, adding sheets,
merging, tables ...) writes a brand-new package from the model, and the model
does not hold everything. carry_forward() runs on the rewrite's output, BEFORE
verify-after-write sees it, and puts back the parts that do not depend on
anything the edit changed, byte-for-byte from the original:

  - docProps/custom.xml: custom document properties, which is also where
    Purview sensitivity labels and SharePoint properties live. LibreOffice
    writes an EMPTY one on every save; dropping it made every LibreOffice
    file uneditable (the inventory check refused the save, correctly).
  - customXml/: SharePoint/OneDrive property sets, add-in stores and the
    Power Query M payload. They hang off the workbook relationships and
    nothing in the worksheets points at them.
  - docProps/app.xml: company, manager and the rest of the extended
    properties, when the sheet list is unchanged (it lists the sheets).
  - external link parts and their relationships: the rewrite re-serializes
    them and kept only one of the two relationships Excel writes for an
    external book, leaving the link pointing at a relationship that no
    longer existed (Excel refused the file).
  - printer settings: each sheet's printerSettings blob and the r:id on its
    pageSetup element.
  - cached formula results: every formula the edit did not write gets its
    original <v> back (followed through inserted and deleted rows/columns);
    the rewrite strips them all.
  - row heights: the rewrite marks every row that has a height as a custom
    height, which locks Excel's automatic heights at a stale value;
    rows that were automatic stay automatic.
  - the sensitivity label: docMetadata/LabelInfo.xml (Microsoft Purview),
    reached from the root relationships. Dropping it failed every rewrite
    on a labelled workbook.
  - the password to modify: workbook.xml's <fileSharing>, which the model
    does not hold; dropping it silently removed the protection.

Every step is conservative: it acts only when the produced package is in the
shape the step expects, and anything it cannot place is left for
verify-after-write, which still refuses on a lost part.
"""

from __future__ import annotations

import html
import os
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from typing import Callable

from . import surgical as _s

#: Hazard families (core/hazard.py keys) carry_forward always puts back, so
#: the rewrite route's gate does not refuse on them; verify-after-write
#: still refuses if one goes missing. Document properties are carried only
#: when the sheet list is unchanged, which the package decides per save.
CARRIED_KEYS = frozenset({"power_query"})

CT_CUSTOM = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CT_PRINTER = "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"
REL_PRINTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass
class CarryReport:
    carried: list = field(default_factory=list)   # what was put back
    cached_values: int = 0                        # formula results restored
    cached_dropped: int = 0                       # results that could not be
    auto_heights: int = 0                         # rows kept automatic

    def as_dict(self) -> dict:
        out: dict = {}
        if self.carried:
            out["carried_over"] = list(self.carried)
        if self.cached_values:
            out["cached_values_restored"] = self.cached_values
        if self.auto_heights:
            out["automatic_row_heights_kept"] = self.auto_heights
        return out


class _Out:
    """The produced package, edited in memory part by part."""

    def __init__(self, path: str):
        with zipfile.ZipFile(path) as zf:
            self.infos = zf.infolist()
            self.data = {i.filename: zf.read(i.filename) for i in self.infos}
        self.order = [i.filename for i in self.infos]
        self.dirty = False

    def has(self, name: str) -> bool:
        low = name.lower()
        return any(n.lower() == low for n in self.data)

    def real(self, name: str) -> str | None:
        low = name.lower()
        return next((n for n in self.data if n.lower() == low), None)

    def put(self, name: str, data: bytes) -> None:
        if name not in self.data:
            self.order.append(name)
        self.data[name] = data
        self.dirty = True

    def rels(self, part: str) -> list[dict]:
        rp = self.real(_s.rels_path(part))
        return _s.parse_rels(self.data[rp]) if rp else []

    def add_rel(self, owner: str, rtype: str, target: str,
                mode: str | None = None) -> str:
        rp_name = _s.rels_path(owner)
        rp = self.real(rp_name)
        if rp is None:
            rp = rp_name
            data = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    b'<Relationships xmlns="http://schemas.openxmlformats.org/'
                    b'package/2006/relationships"></Relationships>')
        else:
            data = self.data[rp]
        used = {r.get("Id", "") for r in _s.parse_rels(data)}
        n = 1
        while f"rId{n}" in used:
            n += 1
        rid = f"rId{n}"
        extra = f' TargetMode="{mode}"' if mode else ""
        el = (f'<Relationship Id="{rid}" Type="{_s.xml_attr(rtype)}" '
              f'Target="{_s.xml_attr(target)}"{extra}/>').encode("utf-8")
        close = data.rfind(b"</Relationships>")
        if close < 0:
            raise _s.NotSurgical("a relationship part is unreadable")
        self.put(rp, data[:close] + el + data[close:])
        return rid

    def add_override(self, part: str, ctype: str) -> None:
        ct = self.data["[Content_Types].xml"]
        if f'PartName="/{part}"'.encode() in ct:
            return
        el = (f'<Override PartName="/{_s.xml_attr(part)}" '
              f'ContentType="{_s.xml_attr(ctype)}"/>').encode("utf-8")
        close = ct.rfind(b"</Types>")
        self.put("[Content_Types].xml", ct[:close] + el + ct[close:])

    def add_default(self, ext: str, ctype: str) -> None:
        ct = self.data["[Content_Types].xml"]
        if re.search(rb'Extension="' + re.escape(ext.encode()) + rb'"', ct, re.I):
            return
        el = (f'<Default Extension="{ext}" ContentType="{_s.xml_attr(ctype)}"/>'
              ).encode("utf-8")
        close = ct.rfind(b"</Types>")
        self.put("[Content_Types].xml", ct[:close] + el + ct[close:])

    def write(self, path: str) -> None:
        by_name = {i.filename: i for i in self.infos}
        from .safesave import scratch_file
        fd, tmp = scratch_file(os.path.dirname(os.path.abspath(path)),
                               os.path.splitext(path)[1])
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
                for name in self.order:
                    info = by_name.get(name)
                    if info is None:
                        info = zipfile.ZipInfo(name)
                        info.compress_type = zipfile.ZIP_DEFLATED
                    zw.writestr(info, self.data[name])
            shutil.move(tmp, path)
        except BaseException:
            try:
                os.remove(tmp)
            except OSError:
                pass
            raise


def _content_types(pkg: _s.Package) -> tuple[dict, dict]:
    ct = pkg.read("[Content_Types].xml")
    defaults = {}
    overrides = {}
    for m in re.finditer(rb"<(?:[\w.\-]+:)?(Default|Override)\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>", ct):
        a = _s._attrs(m.group(2))
        ctype = html.unescape(a.get("ContentType", b"").decode("utf-8"))
        if m.group(1) == b"Default":
            defaults[a.get("Extension", b"").decode("utf-8").lower()] = ctype
        else:
            overrides[html.unescape(a.get("PartName", b"").decode("utf-8")).lstrip("/").lower()] = ctype
    return defaults, overrides


def _ctype_of(part: str, defaults: dict, overrides: dict) -> str | None:
    hit = overrides.get(part.lower())
    if hit:
        return hit
    ext = part.rsplit(".", 1)[-1].lower() if "." in part.rsplit("/", 1)[-1] else ""
    return defaults.get(ext)


def carry_forward(original: str, produced: str, *,
                  sheets_unchanged: bool,
                  custom_props_unchanged: bool,
                  external_links_unchanged: bool,
                  remap: Callable[[str, str], str | None],
                  keep: Callable[[str, str, str], bool],
                  titles: dict[str, str] | None = None) -> CarryReport:
    """Put the model-independent parts and per-cell caches of the original
    back into the produced package (in place). `remap` follows an original
    address through the edit's row/column inserts and deletes; `keep(sheet,
    old, new)` says whether the formula at `new` is still the one that was
    at `old` (its cached value is then put back); `titles` maps a sheet's
    original title to its title in the produced package (a rename). Never
    raises: a step that cannot run is skipped and verify-after-write judges
    the result."""
    rep = CarryReport()
    try:
        with zipfile.ZipFile(original) as zin:
            src = _s.Package(zin)
            out = _Out(produced)
            defaults, overrides = _content_types(src)
            for step in (_custom_props, _custom_xml, _app_props,
                         _external_links, _printer_settings,
                         _label_info, _file_sharing, _abs_path):
                try:
                    step(src, out, rep, defaults, overrides,
                         sheets_unchanged=sheets_unchanged,
                         custom_props_unchanged=custom_props_unchanged,
                         external_links_unchanged=external_links_unchanged)
                except Exception:  # noqa: BLE001 - verify judges the result
                    continue
            try:
                _sheet_caches(src, out, rep, remap, keep, titles or {})
            except Exception:  # noqa: BLE001
                pass
            try:
                _drawing_links(src, out, rep, defaults, overrides,
                               titles=titles or {})
            except Exception:  # noqa: BLE001 - verify judges the result
                pass
        if out.dirty:
            out.write(produced)
    except Exception:  # noqa: BLE001
        return rep
    return rep


# ------------------------------------------------------------------- steps

def _custom_props(src, out, rep, defaults, overrides, *, custom_props_unchanged, **_):
    want = [r for r in src.rels("") if r.get("Type", "").endswith("/custom-properties")]
    if not want:
        return
    part = _s.resolve_target("", want[0].get("Target", ""))
    if not src.has(part):
        return
    data = src.read(part)
    have = [r for r in out.rels("") if r.get("Type", "").endswith("/custom-properties")]
    if have:
        if not custom_props_unchanged:
            return
        target = out.real(_s.resolve_target("", have[0].get("Target", "")))
        if target is not None and out.data[target] != data:
            out.put(target, data)
            rep.carried.append("custom document properties")
        return
    if out.has(part):
        return
    out.put(part, data)
    out.add_rel("", want[0].get("Type"), part)
    out.add_override(part, _ctype_of(part, defaults, overrides) or CT_CUSTOM)
    rep.carried.append("custom document properties")


def _custom_xml(src, out, rep, defaults, overrides, **_):
    wbp = src.workbook_part()
    rels = [r for r in src.rels(wbp) if r.get("Type", "").endswith("/customXml")]
    if not rels:
        return
    owbp = out.real("xl/workbook.xml") or "xl/workbook.xml"
    have = {_s.resolve_target(owbp, r.get("Target", "")).lower()
            for r in out.rels(owbp) if r.get("Type", "").endswith("/customXml")}
    copied = 0
    todo = []
    for r in rels:
        part = _s.resolve_target(wbp, r.get("Target", ""))
        if not src.has(part) or part.lower() in have:
            continue
        todo.append((r, part))
    # every part reachable from those items (their itemProps), with rels
    seen = set()
    stack = [p for _r, p in todo]
    while stack:
        part = stack.pop()
        if part.lower() in seen or not src.has(part):
            continue
        seen.add(part.lower())
        if not out.has(part):
            out.put(src.real(part), src.read(part))
            ct = _ctype_of(part, defaults, overrides)
            if ct and part.lower() in overrides:
                out.add_override(src.real(part), ct)
            copied += 1
        rp = _s.rels_path(part)
        if src.has(rp) and not out.has(rp):
            out.put(src.real(rp), src.read(rp))
        for rr in src.rels(part):
            if rr.get("TargetMode") != "External":
                stack.append(_s.resolve_target(part, rr.get("Target", "")))
    for r, part in todo:
        target = os.path.relpath(part, os.path.dirname(owbp)).replace("\\", "/")
        out.add_rel(owbp, r.get("Type"), target)
    if copied:
        rep.carried.append("custom XML parts (SharePoint/OneDrive properties, "
                           "add-in data, Power Query)")



REL_LABELS_SUFFIX = "/relationships/classificationlabels"
CT_LABELS = "application/vnd.ms-office.classificationlabels+xml"


def _label_info(src, out, rep, defaults, overrides, **_):
    """The sensitivity-label part and its root relationship. The label
    describes the whole file and nothing in it points into the sheets, so
    it is put back byte-for-byte whatever the edit was."""
    want = [r for r in src.rels("") if r.get("Type", "").endswith(REL_LABELS_SUFFIX)]
    if not want:
        return
    have = {_s.resolve_target("", r.get("Target", "")).lower()
            for r in out.rels("") if r.get("Type", "").endswith(REL_LABELS_SUFFIX)}
    copied = False
    for r in want:
        part = _s.resolve_target("", r.get("Target", ""))
        if not src.has(part):
            continue
        real = src.real(part)
        if not out.has(part):
            out.put(real, src.read(part))
            out.add_override(real, _ctype_of(part, defaults, overrides) or CT_LABELS)
            copied = True
        if part.lower() not in have:
            out.add_rel("", r.get("Type"), real)
            copied = True
    if copied:
        rep.carried.append("sensitivity label (docMetadata/LabelInfo.xml)")


_FILE_SHARING_EL = re.compile(rb"<fileSharing\b(?:[^>\"']|\"[^\"]*\"|'[^']*')*?/>")


def _file_sharing(src, out, rep, defaults, overrides, **_):
    """The password to modify (<fileSharing>), which the openpyxl model does
    not hold. Put back unchanged, in schema position: after fileVersion and
    before workbookPr. Only the unprefixed form Excel writes is handled; a
    prefixed one is left alone (verify then judges the file as it is)."""
    wbp = src.workbook_part()
    owbp = out.real("xl/workbook.xml")
    if owbp is None or not src.has(wbp):
        return
    m = _FILE_SHARING_EL.search(src.read(wbp))
    if m is None:
        return
    data = out.data[owbp]
    if b"<fileSharing" in data:
        return
    el = m.group(0)
    k = data.find(b"<workbookPr")
    if k < 0:
        fv = re.search(rb"<fileVersion\b[^>]*?/>", data)
        if fv is not None:
            k = fv.end()
        else:
            root = re.search(rb"<workbook\b[^>]*>", data)
            if root is None or root.group(0).endswith(b"/>"):
                return
            k = root.end()
    out.put(owbp, data[:k] + el + data[k:])
    rep.carried.append("password to modify")


_ALT_ABSPATH = re.compile(
    rb"<(?P<p>[A-Za-z_][\w.\-]*):AlternateContent\b[^>]*>"
    rb"(?:(?!</(?P=p):AlternateContent>).)*?absPath\b"
    rb"(?:(?!</(?P=p):AlternateContent>).)*?</(?P=p):AlternateContent>", re.S)
_NS_MC = b"http://schemas.openxmlformats.org/markup-compatibility/2006"
_NS_X15 = b"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"


def _abs_path(src, out, rep, defaults, overrides, **_):
    """The folder the workbook was last saved in (x15ac:absPath,
    inside an mc:AlternateContent block of workbook.xml). Excel resolves a
    relative external link against it when the file is opened from another
    folder; dropping it changed where such links pointed. Put back after
    workbookPr, with the namespaces it needs declared on the block itself."""
    wbp = src.workbook_part()
    owbp = out.real("xl/workbook.xml")
    if owbp is None or not src.has(wbp):
        return
    m = _ALT_ABSPATH.search(src.read(wbp))
    if m is None:
        return
    data = out.data[owbp]
    if b"absPath" in data:
        return
    block = m.group(0)
    p = m.group("p")
    head_end = block.index(b">")
    head = block[:head_end]
    if b"xmlns:" + p + b"=" not in head:
        head += b" xmlns:" + p + b'="' + _NS_MC + b'"'
    if b'Requires="x15"' in block and b"xmlns:x15=" not in head:
        head += b' xmlns:x15="' + _NS_X15 + b'"'
    block = head + block[head_end:]
    wp = re.search(rb"<workbookPr\b(?:[^>\"']|\"[^\"]*\"|'[^']*')*?(?:/>|>.*?</workbookPr>)",
                   data, re.S)
    if wp is None:
        return
    out.put(owbp, data[:wp.end()] + block + data[wp.end():])
    rep.carried.append("the folder the workbook was saved in (for relative links)")


def _app_props(src, out, rep, defaults, overrides, *, sheets_unchanged, **_):
    if not sheets_unchanged:
        return
    want = [r for r in src.rels("") if r.get("Type", "").endswith("/extended-properties")]
    have = [r for r in out.rels("") if r.get("Type", "").endswith("/extended-properties")]
    if not want or not have:
        return
    sp = _s.resolve_target("", want[0].get("Target", ""))
    op = out.real(_s.resolve_target("", have[0].get("Target", "")))
    if not src.has(sp) or op is None:
        return
    data = src.read(sp)
    if out.data[op] != data:
        out.put(op, data)
        rep.carried.append("document properties (app.xml)")


def _external_links(src, out, rep, defaults, overrides, *,
                    external_links_unchanged, **_):
    if not external_links_unchanged:
        return
    wbp = src.workbook_part()
    owbp = out.real("xl/workbook.xml") or "xl/workbook.xml"
    s_links = _ordered_links(src.read(wbp), src.rels(wbp), wbp)
    o_links = _ordered_links(out.data[owbp], out.rels(owbp), owbp)
    if not s_links or len(s_links) != len(o_links):
        return
    n = 0
    for sp, op in zip(s_links, o_links):
        if not src.has(sp):
            continue
        real_op = out.real(op)
        if real_op is None:
            continue
        data = src.read(sp)
        if out.data[real_op] != data:
            out.put(real_op, data)
            n += 1
        srp = _s.rels_path(sp)
        orp = _s.rels_path(real_op)
        if src.has(srp):
            sdata = src.read(srp)
            real_orp = out.real(orp) or orp
            if out.data.get(real_orp) != sdata:
                out.put(real_orp, sdata)
                n += 1
    if n:
        rep.carried.append("external link definitions")


def _ordered_links(wb_xml: bytes, rels: list, owner: str) -> list[str]:
    by_id = {r.get("Id"): _s.resolve_target(owner, r.get("Target", "")) for r in rels}
    out = []
    for m in re.finditer(rb"<(?:[\w.\-]+:)?externalReference\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>", wb_xml):
        a = _s._attrs(m.group(1))
        rid = next((v for k, v in a.items() if k.endswith(":id") or k == "id"), None)
        if rid is None:
            return []
        part = by_id.get(rid.decode("utf-8"))
        if part is None:
            return []
        out.append(part)
    return out


_PAGESETUP_RE = re.compile(rb"<((?:[A-Za-z_][\w.\-]*:)?)pageSetup\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>")


def _printer_settings(src, out, rep, defaults, overrides, **_):
    s_sheets = src.sheet_parts()
    o_pkg = _OutView(out)
    o_sheets = o_pkg.sheet_parts()
    n = 0
    for title, sp in s_sheets.items():
        op = o_sheets.get(title)
        if op is None or not src.has(sp):
            continue
        op = out.real(op)
        if op is None:
            continue
        sdata = src.read(sp)
        m = _PAGESETUP_RE.search(sdata)
        if m is None:
            continue
        rid = next((v for k, v in _s._attrs(m.group(2)).items() if k.endswith(":id")), None)
        if rid is None:
            continue
        rel = next((r for r in src.rels(sp) if r.get("Id") == rid.decode("utf-8")), None)
        if rel is None or not rel.get("Type", "").endswith("/printerSettings"):
            continue
        blob_part = _s.resolve_target(sp, rel.get("Target", ""))
        if not src.has(blob_part):
            continue
        odata = out.data[op]
        om = _PAGESETUP_RE.search(odata)
        if om is None or any(k.endswith(":id") for k in _s._attrs(om.group(2))):
            continue
        target_part = src.real(blob_part)
        if out.has(target_part):
            continue
        out.put(target_part, src.read(blob_part))
        rel_target = os.path.relpath(target_part, os.path.dirname(op)).replace("\\", "/")
        new_rid = out.add_rel(op, REL_PRINTER, rel_target)
        odata = out.data[op]
        odata, prefix, _shift = _s.worksheet_rel_prefix(odata)
        om = _PAGESETUP_RE.search(odata)
        tag = _s._set_attr(om.group(0), prefix + b":id", new_rid.encode())
        out.put(op, odata[:om.start()] + tag + odata[om.end():])
        ext = target_part.rsplit(".", 1)[-1]
        ct = _ctype_of(target_part, defaults, overrides) or CT_PRINTER
        if target_part.lower() in overrides:
            out.add_override(target_part, ct)
        else:
            out.add_default(ext, ct)
        n += 1
    if n:
        rep.carried.append("printer settings")


class _OutView:
    """surgical.Package's sheet lookup over the in-memory produced package."""

    def __init__(self, out: _Out):
        self.out = out

    def has(self, name):
        return self.out.has(name)

    def read(self, name):
        return self.out.data[self.out.real(name)]

    def rels(self, part):
        return self.out.rels(part)

    def workbook_part(self):
        return _s.Package.workbook_part(self)

    def sheet_parts(self):
        return _s.Package.sheet_parts(self)


# ------------------------------------------------- picture and shape links

_HLINK_RE = re.compile(
    rb"<(?P<p>(?:[A-Za-z_][\w.\-]*:)?)(?P<tag>hlinkClick|hlinkHover)\b"
    rb"(?P<attrs>(?:[^>/\"']|\"[^\"]*\"|'[^']*')*)"
    rb"(?:/>|>.*?</(?P=p)(?P=tag)>)", re.S)
_RID_ATTR = re.compile(rb"(\b[\w.\-]+:id=\")([^\"]*)(\")")
_HYPERLINK = ("http://schemas.openxmlformats.org/officeDocument/2006/"
              "relationships/hyperlink")


def _drawing_links(src, out, rep, defaults, overrides, *, titles=None, **_):
    """A picture's hyperlink after a full rewrite. openpyxl keeps the link's
    relationship id in the drawing XML (a:hlinkClick r:id) but writes only
    the image relationships, so the id names nothing (Excel repairs or
    refuses the file) or, worse, names the image. Each such link goes back
    under a fresh id with the target the original drawing gave it; a link
    the original does not explain is removed."""
    view = _OutView(out)
    back = {new: old for old, new in (titles or {}).items()}
    src_sheets = src.sheet_parts()
    for title, part in view.sheet_parts().items():
        old_part = src_sheets.get(back.get(title, title))
        if old_part is None:
            continue
        new_dr = _s.Package.related(view, part, "/drawing")
        old_dr = src.related(old_part, "/drawing")
        if len(new_dr) != 1 or len(old_dr) != 1 or not out.has(new_dr[0]) \
                or not src.has(old_dr[0]):
            continue
        dp, odp = out.real(new_dr[0]), old_dr[0]
        data = out.data[dp]
        if b"hlink" not in data:
            continue
        links = {r.get("Id"): r for r in src.rels(odp)
                 if r.get("Type", "").endswith("/hyperlink")}
        now = {r.get("Id"): r for r in out.rels(dp)}
        kept = dropped = 0

        def fix(m):
            nonlocal kept, dropped
            idm = _RID_ATTR.search(m.group("attrs"))
            if idm is None:
                return m.group(0)
            rid = idm.group(2).decode("utf-8", "replace")
            cur = now.get(rid)
            if cur is not None and cur.get("Type", "").endswith("/hyperlink"):
                return m.group(0)  # a link the drawing defines
            orig = links.get(rid)
            if orig is None:
                dropped += 1
                return b""
            new_id = out.add_rel(dp, _HYPERLINK, orig.get("Target", ""),
                                 orig.get("TargetMode"))
            now[new_id] = {"Id": new_id, "Type": _HYPERLINK}
            kept += 1
            whole = m.group(0)
            return whole.replace(idm.group(0), idm.group(1)
                                 + new_id.encode("ascii") + idm.group(3), 1)

        fixed = _HLINK_RE.sub(fix, data)
        if fixed != data:
            out.put(dp, fixed)
        if kept:
            rep.carried.append(f"{kept} picture link(s) on {title}")
        if dropped:
            rep.carried.append(f"{dropped} picture link(s) on {title} "
                               "removed: the original names no target")


# ----------------------------------------------------- per-cell caches, rows

_CELL_RE = re.compile(
    rb"<(?P<p>(?:[A-Za-z_][\w.\-]*:)?)c(?=[\s/>])"
    rb"(?P<attrs>(?:[^>/\"']|\"[^\"]*\"|'[^']*')*)"
    rb"(?:/>|>(?P<body>.*?)</(?P=p)c>)", re.S)
_V_RE = re.compile(rb"<((?:[A-Za-z_][\w.\-]*:)?)v\b[^>]*?(?:/>|>(.*?)</\1v>)", re.S)
_F_RE = re.compile(rb"<((?:[A-Za-z_][\w.\-]*:)?)f\b")
_ROW_RE = re.compile(rb"<((?:[A-Za-z_][\w.\-]*:)?)row\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>")


def _orig_formula_values(data: bytes) -> tuple[dict, dict]:
    """addr -> (t attribute or None, <v> element bytes) for formula cells
    that carry a cached value; and row -> had customHeight."""
    vals: dict = {}
    if b"<f" in data or b":f" in data:
        for m in _CELL_RE.finditer(data):
            body = m.group("body")
            if not body or not _F_RE.search(body):
                continue
            vm = _V_RE.search(body)
            if vm is None or vm.group(2) in (None, b""):
                continue
            a = _s._attrs(m.group("attrs"))
            r = a.get("r")
            if r is None:
                continue
            vals[r.decode("ascii")] = (a.get("t"), vm.group(0))
    rows: dict = {}
    for m in _ROW_RE.finditer(data):
        a = _s._attrs(m.group(2))
        if "r" in a and "ht" in a:
            num = _s._row_number(a["r"])  # tolerates <row r="1.0">
            if num is not None:
                rows[num] = a.get("customHeight") in (b"1", b"true")
    return vals, rows


def _present(data: bytes, wanted: dict) -> tuple[int, int]:
    """(wanted formula cells that carry a cached value, wanted cells that are
    formulas) in the produced part."""
    n = 0
    total = 0
    for m in _CELL_RE.finditer(data):
        body = m.group("body")
        if not body or not _F_RE.search(body):
            continue
        a = _s._attrs(m.group("attrs"))
        r = a.get("r")
        if r is None or r.decode("ascii") not in wanted:
            continue
        total += 1
        vm = _V_RE.search(body)
        if vm is not None and vm.group(2) not in (None, b""):
            n += 1
    return n, total


def _sheet_caches(src, out, rep, remap, keep, titles=None) -> None:
    s_sheets = src.sheet_parts()
    o_sheets = _OutView(out).sheet_parts()
    for title, sp in s_sheets.items():
        # a renamed sheet's cells keep their results under the new title
        op = o_sheets.get((titles or {}).get(title, title))
        if op is None or not src.has(sp):
            continue
        op = out.real(op)
        if op is None:
            continue
        sdata = src.read(sp)
        if sdata[:2] in (b"\xff\xfe", b"\xfe\xff"):
            continue
        vals, rows = _orig_formula_values(sdata)
        wanted: dict = {}
        for addr, cached in vals.items():
            new = remap(title, addr)
            if new is None or not keep(title, addr, new):
                continue
            wanted[new] = cached
        auto_rows = set()
        for r, custom in rows.items():
            if custom:
                continue
            new = remap(title, f"A{r}")
            if new is not None:
                auto_rows.add(int(re.sub(r"^[A-Z]+", "", new)))
        if not wanted and not auto_rows:
            continue
        odata = out.data[op]
        n_vals = 0
        n_rows = 0

        def fix_cell(m):
            nonlocal n_vals
            a = _s._attrs(m.group("attrs"))
            r = a.get("r")
            body = m.group("body")
            if r is None or body is None:
                return m.group(0)
            cached = wanted.get(r.decode("ascii"))
            if cached is None or not _F_RE.search(body):
                return m.group(0)
            t, v_el = cached
            vm = _V_RE.search(body)
            if vm is not None and vm.group(2) not in (None, b""):
                return m.group(0)
            if vm is not None:
                body = body[:vm.start()] + body[vm.end():]
            prefix = m.group("p")
            p = prefix
            v_el = re.sub(rb"^<(?:[A-Za-z_][\w.\-]*:)?v", b"<" + p + b"v", v_el)
            v_el = re.sub(rb"</(?:[A-Za-z_][\w.\-]*:)?v>$", b"</" + p + b"v>", v_el)
            tag = b"<" + prefix + b"c" + m.group("attrs") + b">"
            if t is not None and t != b"n":
                tag = _s._set_attr(tag, b"t", t)
            n_vals += 1
            return tag + body + v_el + b"</" + prefix + b"c>"

        if wanted:
            odata = _CELL_RE.sub(fix_cell, odata)

        def fix_row(m):
            nonlocal n_rows
            a = _s._attrs(m.group(2))
            try:
                r = int(a.get("r", b"0"))
            except ValueError:
                return m.group(0)
            if r in auto_rows and a.get("customHeight") in (b"1", b"true"):
                n_rows += 1
                return _s._set_attr(m.group(0), b"customHeight", None)
            return m.group(0)

        if auto_rows:
            odata = _ROW_RE.sub(fix_row, odata)
        have, formulas = _present(odata, wanted)
        rep.cached_dropped += max(0, formulas - have)
        if n_vals or n_rows:
            out.put(op, odata)
            rep.cached_values += n_vals
            rep.auto_heights += n_rows


__all__ = ["carry_forward", "CarryReport"]
