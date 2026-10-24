"""Made-up workbooks for the XML and zip safety tests, built in code.

Nothing here comes from a real file. `make_xlsx` writes a small package that
Excel opens: one sheet holding a shared string, an inline string and two
numbers (one with a custom number format), a legacy note with its VML anchor,
a defined name and core document properties. `chart=True` adds a drawing that
anchors a bar chart. `parts` replaces or adds parts by name.

`hostile(kind, attack, ...)` returns one part of that package with a document
type declaration put into it. Where the part has a made-up text (or an
attribute) of its own, the attack's entity reference replaces it, so a reader
that expands the entity shows the expansion in what it reads back. Every
target a declaration names is a made-up file in the test's own folder or a
listener on 127.0.0.1: nothing here names an outside host."""

from __future__ import annotations

import zipfile
from pathlib import Path

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
OD = "application/vnd.openxmlformats-officedocument"

#: What a small internal entity expands to. A read that shows these words
#: expanded an entity it should never have read.
EXPANSION = "made-up words"
#: What the made-up local target file of an external entity holds. It must
#: never appear in anything the server answers.
MARKER = "MADE-UP-MARKER-5c1e"


def _content_types(chart: bool) -> str:
    over = [
        ("/xl/workbook.xml", f"{OD}.spreadsheetml.sheet.main+xml"),
        ("/xl/worksheets/sheet1.xml", f"{OD}.spreadsheetml.worksheet+xml"),
        ("/xl/sharedStrings.xml", f"{OD}.spreadsheetml.sharedStrings+xml"),
        ("/xl/styles.xml", f"{OD}.spreadsheetml.styles+xml"),
        ("/xl/comments1.xml", f"{OD}.spreadsheetml.comments+xml"),
        ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
        ("/docProps/app.xml", f"{OD}.extended-properties+xml"),
    ]
    if chart:
        over += [("/xl/drawings/drawing1.xml", f"{OD}.drawing+xml"),
                 ("/xl/charts/chart1.xml", f"{OD}.drawingml.chart+xml")]
    body = "".join(f'<Override PartName="{n}" ContentType="{t}"/>' for n, t in over)
    return (f'{DECL}<Types xmlns="{CT}">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Default Extension="vml" ContentType="{OD}.vmlDrawing"/>'
            f"{body}</Types>")


def _rels(*rels: tuple[str, str, str]) -> str:
    body = "".join(f'<Relationship Id="{i}" Type="{t}" Target="{g}"/>' for i, t, g in rels)
    return f'{DECL}<Relationships xmlns="{PKG_REL}">{body}</Relationships>'


def _sheet(chart: bool) -> str:
    drawing = '<drawing r:id="rId3"/>' if chart else ""
    return (f'{DECL}<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><dimension ref="A1:B2"/>'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/><sheetData>'
            '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>'
            '<row r="2"><c r="A2" t="inlineStr"><is><t>Made-up inline text</t></is></c>'
            '<c r="B2" s="1"><v>7.5</v></c></row></sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            f'{drawing}<legacyDrawing r:id="rId1"/></worksheet>')


_STYLES = (
    f'{DECL}<styleSheet xmlns="{MAIN}">'
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>'
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>")

_VML = (
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:x="urn:schemas-microsoft-com:office:excel">'
    '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
    'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
    '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
    '<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;'
    'margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" '
    'o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>'
    '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto">'
    '<div style="text-align:left">Made-up anchor text</div></v:textbox>'
    '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
    '<x:Anchor>2, 15, 0, 2, 4, 15, 3, 16</x:Anchor><x:AutoFill>False</x:AutoFill>'
    '<x:Row>0</x:Row><x:Column>1</x:Column></x:ClientData></v:shape></xml>')

_DRAWING = (
    f'{DECL}<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{DML}"><xdr:twoCellAnchor>'
    "<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>"
    "<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff>"
    "<xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
    '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Made-up chart"/>'
    "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/>"
    '<a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="' + CHART + '">'
    f'<c:chart xmlns:c="{CHART}" xmlns:r="{REL}" r:id="rId1"/></a:graphicData></a:graphic>'
    "</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>")

_CHART = (
    f'{DECL}<c:chartSpace xmlns:c="{CHART}" xmlns:a="{DML}" xmlns:r="{REL}"><c:chart>'
    '<c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/>'
    '<c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>'
    "<c:tx><c:v>Made-up series</c:v></c:tx>"
    "<c:val><c:numRef><c:f>Sheet1!$B$1:$B$2</c:f></c:numRef></c:val></c:ser>"
    '<c:axId val="10"/><c:axId val="20"/></c:barChart>'
    '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
    '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>'
    '<c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
    '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx></c:plotArea>'
    '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>')


def base_parts(chart: bool = False) -> dict[str, str]:
    """Every part of the made-up workbook, by name."""
    sheet_rels = [("rId1", f"{REL}/vmlDrawing", "../drawings/vmlDrawing1.vml"),
                  ("rId2", f"{REL}/comments", "../comments1.xml")]
    if chart:
        sheet_rels.append(("rId3", f"{REL}/drawing", "../drawings/drawing1.xml"))
    parts = {
        "[Content_Types].xml": _content_types(chart),
        "_rels/.rels": _rels(
            ("rId1", f"{REL}/officeDocument", "xl/workbook.xml"),
            ("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/"
                     "core-properties", "docProps/core.xml"),
            ("rId3", f"{REL}/extended-properties", "docProps/app.xml")),
        "docProps/core.xml": (
            f'{DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/'
            'metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">'
            "<dc:title>Made-up title</dc:title><dc:creator>Made-up Author</dc:creator>"
            "</cp:coreProperties>"),
        "docProps/app.xml": (
            f'{DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
            'extended-properties"><Application>Microsoft Excel</Application></Properties>'),
        "xl/workbook.xml": (
            f'{DECL}<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews>'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
            '<definedNames><definedName name="MadeUpTotal">Sheet1!$B$1</definedName>'
            "</definedNames></workbook>"),
        "xl/_rels/workbook.xml.rels": _rels(
            ("rId1", f"{REL}/worksheet", "worksheets/sheet1.xml"),
            ("rId2", f"{REL}/sharedStrings", "sharedStrings.xml"),
            ("rId3", f"{REL}/styles", "styles.xml")),
        "xl/worksheets/sheet1.xml": _sheet(chart),
        "xl/worksheets/_rels/sheet1.xml.rels": _rels(*sheet_rels),
        "xl/sharedStrings.xml": (
            f'{DECL}<sst xmlns="{MAIN}" count="1" uniqueCount="1"><si><t>Made-up label</t></si></sst>'),
        "xl/styles.xml": _STYLES,
        "xl/comments1.xml": (
            f'{DECL}<comments xmlns="{MAIN}"><authors><author>Made-up Author</author></authors>'
            '<commentList><comment ref="B1" authorId="0"><text><r><t>Made-up note</t></r></text>'
            "</comment></commentList></comments>"),
        "xl/drawings/vmlDrawing1.vml": _VML,
    }
    if chart:
        parts["xl/drawings/drawing1.xml"] = _DRAWING
        parts["xl/drawings/_rels/drawing1.xml.rels"] = _rels(
            ("rId1", f"{REL}/chart", "../charts/chart1.xml"))
        parts["xl/charts/chart1.xml"] = _CHART
    return parts


def make_xlsx(path: Path, *, parts: dict[str, bytes | str] | None = None,
              chart: bool = False, extra: list[tuple[str, bytes | str]] | None = None) -> Path:
    """Write the made-up workbook. `parts` replaces or adds parts by name;
    `extra` appends members as given (a repeated name stays repeated)."""
    base: dict[str, bytes | str] = dict(base_parts(chart))
    base.update(parts or {})
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in base.items():
            z.writestr(name, data.encode("utf-8") if isinstance(data, str) else data)
        for name, data in extra or []:
            z.writestr(name, data.encode("utf-8") if isinstance(data, str) else data)
    return path


# ---------------------------------------------------------------- the parts

#: kind -> (part name, root element name, the made-up text or attribute value
#: an entity reference replaces, whether that slot is text or an attribute,
#: needs the chart base). None: the part has no slot of its own, so an
#: internal entity is declared and not referenced.
KINDS: dict[str, tuple[str, str, str | None, str, bool]] = {
    "workbook": ("xl/workbook.xml", "workbook", "Sheet1!$B$1", "text", False),
    "sheet": ("xl/worksheets/sheet1.xml", "worksheet", "Made-up inline text", "text", False),
    "sharedStrings": ("xl/sharedStrings.xml", "sst", "Made-up label", "text", False),
    "styles": ("xl/styles.xml", "styleSheet", "0.000", "attr", False),
    "comments": ("xl/comments1.xml", "comments", "Made-up note", "text", False),
    "vml": ("xl/drawings/vmlDrawing1.vml", "xml", "Made-up anchor text", "text", False),
    "drawing": ("xl/drawings/drawing1.xml", "xdr:wsDr", "Made-up chart", "attr", True),
    "chart": ("xl/charts/chart1.xml", "c:chartSpace", "Made-up series", "text", True),
    "workbook_rels": ("xl/_rels/workbook.xml.rels", "Relationships", None, "text", False),
    "content_types": ("[Content_Types].xml", "Types", None, "text", False),
    "core_props": ("docProps/core.xml", "cp:coreProperties", "Made-up title", "text", False),
}

#: The part kinds a read of the workbook model loads.
READ_KINDS = ("workbook", "sheet", "sharedStrings", "styles", "comments", "drawing", "chart",
              "workbook_rels", "content_types", "core_props")

ATTACKS = ("billion_laughs", "quadratic_blowup", "small_internal_entity", "external_entity_file",
           "external_entity_http", "parameter_entity_file", "parameter_entity_http",
           "dtd_no_entities")


def _laughs() -> tuple[str, str]:
    ents = '<!ENTITY a "made up made up ">' + "".join(
        f'<!ENTITY {chr(98 + i)} "{("&" + chr(97 + i) + ";") * 10}">' for i in range(9))
    return ents, "&j;"


def declaration(attack: str, *, target_url: str = "", dtd_url: str = "",
                http_base: str = "", slot: str = "text") -> tuple[str, str | None]:
    """(the internal subset's text, the reference that replaces the slot or
    None). An external reference never goes into an attribute (XML forbids
    it), so an attribute slot keeps its text for those."""
    if attack == "billion_laughs":
        return _laughs()
    if attack == "quadratic_blowup":
        return f'<!ENTITY q "{"Q" * 40000}">', "&q;" * 40000
    if attack == "small_internal_entity":
        return f'<!ENTITY e "{EXPANSION}">', "&e;"
    if attack == "external_entity_file":
        return f'<!ENTITY e SYSTEM "{target_url}">', ("&e;" if slot == "text" else None)
    if attack == "external_entity_http":
        return (f'<!ENTITY e SYSTEM "{http_base}/made-up-target.txt">',
                "&e;" if slot == "text" else None)
    if attack == "parameter_entity_file":
        return f'<!ENTITY % p SYSTEM "{dtd_url}"> %p;', None
    if attack == "parameter_entity_http":
        return f'<!ENTITY % p SYSTEM "{http_base}/made-up.dtd"> %p;', None
    if attack == "dtd_no_entities":
        return "", None
    raise AssertionError(attack)


def hostile(kind: str, attack: str, *, target_url: str = "", dtd_url: str = "",
            http_base: str = "", encoding: str = "utf-8") -> bytes:
    """One part of the made-up workbook with the attack's declaration in its
    prolog, encoded as `encoding` ("utf-8", "utf-16-le" or "utf-16-be"; the
    UTF-16 forms carry a byte order mark and say UTF-16 in the declaration)."""
    name, root, slot_text, slot, chart = KINDS[kind]
    xml = base_parts(chart)[name]
    subset, ref = declaration(attack, target_url=target_url, dtd_url=dtd_url,
                              http_base=http_base, slot=slot)
    doctype = f"<!DOCTYPE {root} [{subset}]>" if subset else f"<!DOCTYPE {root}>"
    if xml.startswith(DECL):
        xml = DECL + doctype + xml[len(DECL):]
    else:
        xml = doctype + xml
    if ref is not None and slot_text is not None:
        assert xml.count(slot_text) >= 1, (kind, slot_text)
        xml = xml.replace(slot_text, ref, 1)
    if encoding == "utf-8":
        return xml.encode("utf-8")
    xml = xml.replace('encoding="UTF-8"', 'encoding="UTF-16"', 1)
    bom = b"\xff\xfe" if encoding == "utf-16-le" else b"\xfe\xff"
    return bom + xml.encode(encoding)


# --------------------------------------------------------- zip packages

def make_package(path: Path, members: list[tuple[str, bytes | str]]) -> Path:
    """Write `members` exactly as given and in order: a name may repeat, and
    a backslash in a name is stored as it is (zipfile.ZipInfo() would turn
    it into "/" on Windows)."""
    import warnings

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # zipfile warns about a repeated name
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in members:
                info = zipfile.ZipInfo("made-up", date_time=(2026, 1, 1, 0, 0, 0))
                info.filename = name
                info.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(info, data.encode("utf-8") if isinstance(data, str) else data)
    return path


def base_members(chart: bool = False) -> list[tuple[str, str]]:
    """The made-up workbook's parts as (name, text), in package order."""
    return list(base_parts(chart).items())


#: What a bomb part repeats: spaces in an XML part (white space between
#: elements), more one-letter strings in the shared-string part, zero bytes
#: in any other member.
_FILL = {"xl/sharedStrings.xml": b"<si><t>a</t></si>"}


def _stream(z: zipfile.ZipFile, name: str, size: int, head: bytes, tail: bytes,
            fill: bytes) -> None:
    body = (size - len(head) - len(tail)) // len(fill) * len(fill)
    unit = fill * max(1, (1 << 20) // len(fill))
    with z.open(name, "w", force_zip64=size > (3 << 29)) as fh:
        fh.write(head)
        while body > 0:
            piece = unit if body >= len(unit) else unit[:body]
            fh.write(piece)
            body -= len(piece)
        fh.write(tail)


def make_bomb_xlsx(path: Path, bombs: dict[str, int], *, chart: bool = False) -> Path:
    """The made-up workbook with each member named in `bombs` written as a
    part of about that many bytes that deflates at about 400 to 1,000 to 1.
    It is built in 1 MiB steps, so a 1 GiB part never sits in memory. A part
    of the workbook stays well-formed XML: the fill goes just before its
    root's end tag. Any other name is a new member of zero bytes."""
    parts = base_parts(chart)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            if name not in bombs:
                z.writestr(name, data)
        for name, size in bombs.items():
            if name in parts:
                xml = parts[name].encode("utf-8")
                cut = xml.rindex(b"</")
                head, tail = xml[:cut], xml[cut:]
                fill = _FILL.get(name, b" ")
            else:
                head, tail, fill = b"", b"", b"\x00"
            _stream(z, name, size, head, tail, fill)
    return path


def made_up_targets(folder: Path) -> tuple[Path, Path]:
    """The made-up local files an external entity and a parameter entity
    point at. Their text must never come back from the server."""
    target = folder / "made_up_target.txt"
    target.write_text(MARKER, encoding="utf-8")
    dtd = folder / "made_up.dtd"
    dtd.write_text(f'<!ENTITY leak "{MARKER}">', encoding="utf-8")
    return target, dtd
