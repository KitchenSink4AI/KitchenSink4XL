"""One hardened way to parse the XML inside a workbook package.

Every XML part this server parses goes through this module: no DTD loading,
no network access, entities never substituted, and libxml2's default size
limits (huge_tree off). openpyxl, which builds the workbook model, is called
only through `load_workbook` here, which checks each part openpyxl reads
before openpyxl sees it.

A part that carries a document type declaration (<!DOCTYPE ...>) is refused
with a plain sentence, never parsed for its content. The Open Packaging
Conventions do not allow a DTD in a package's XML (ECMA-376 Part 2), so a
file Excel wrote never has one. Refusing is the only safe reading:
- expanding the entities a DTD declares is how a "billion laughs" part makes
  a reader build gigabytes, and how an external entity reads another file or
  a network address;
- leaving them unexpanded is no answer either. Measured on openpyxl 3.1.5:
  its worksheet and shared-string reader (the standard library's expat)
  EXPANDED a small internal entity into the cell text, and its other readers
  (lxml with entities off) dropped an entity's text from an element but
  still expanded one inside an attribute.

The check runs at three places: a look at each XML part's prolog (the text
before its root element, the only place XML allows the declaration) when a
package is opened, so the refusal comes before any parsing whichever op
asked (after zipsafe.check_members, which refuses a decompression bomb or a
part named twice from the ZIP directory alone); the same look at every part
openpyxl reads, as it reads it (a part a relationship names under any name
included); and a look at the parsed document's own DOCTYPE after every parse
here.

The prolog look asks libxml2 itself, so it decodes the part exactly as the
parsers do, in every encoding they read (UTF-7 included, where "<!DOCTYPE"
is not the ASCII bytes); it stops at the root element's start tag. A prolog
libxml2 cannot read is walked as text: the bytes for an ASCII-compatible
encoding, the decoded characters for UTF-16 and UTF-32.
"""

from __future__ import annotations

import os
import zipfile
import zlib
from xml.etree.ElementTree import ParseError

from lxml import etree

from . import partnames as _partnames
from . import zipsafe as _zipsafe

#: The parser options every parse takes. dtd_validation and
#: attribute_defaults would read the DTD; all three entity switches are off.
_FLAGS = dict(
    resolve_entities=False,
    no_network=True,
    load_dtd=False,
    dtd_validation=False,
    attribute_defaults=False,
    huge_tree=False,
)

_FIXED = ("resolve_entities", "no_network", "load_dtd", "dtd_validation", "attribute_defaults")

#: The tree the parse sites here were written against (the standard
#: library's ElementTree): XML comments and processing instructions are not
#: part of it.
_ET_LAYOUT = dict(remove_comments=True, remove_pis=True)

#: The XML exceptions a caller may need to name without importing a parser.
XML_ERRORS: tuple[type[BaseException], ...] = (etree.LxmlError, ParseError)

#: The member names whose content is XML in an Open XML package.
_XML_SUFFIXES = (".xml", ".rels", ".vml")

#: How much of a part the prolog look reads at first; a longer prolog (a
#: long comment before the root) is read on in larger steps.
_HEAD = 1 << 16


def _checked(overrides: dict) -> dict:
    for key in _FIXED:
        if key in overrides:
            raise ValueError(f"xmlsafe: {key} is fixed")
    return {**_FLAGS, **overrides}


def parser(**overrides) -> etree.XMLParser:
    """A new hardened parser (lxml parsers are not shared between threads).
    `overrides` may set layout options (remove_blank_text, remove_comments,
    remove_pis, strip_cdata); never an entity or DTD switch."""
    return etree.XMLParser(**_checked(overrides))


def stream_parse(source, *, events, tag=None, **layout):
    """lxml's iterparse over a binary stream with the hardened settings (no
    DTD, no entities, no network, no huge trees); `layout` may set layout
    options only, as for parser()."""
    return etree.iterparse(source, events=events, tag=tag, **_checked(layout))


def pull_parser(*, events, **layout):
    """lxml's feed parser with the hardened settings; `layout` as above."""
    return etree.XMLPullParser(events=events, **_checked(layout))


def xpath(expr: str, **kw):
    """A compiled XPath, run over trees a hardened parse built (it parses
    nothing itself)."""
    return etree.XPath(expr, **kw)


#: lxml's base error and its syntax error, for callers that read with the
#: functions above.
LXML_ERROR = etree.LxmlError
SYNTAX_ERROR = etree.XMLSyntaxError


# ------------------------------------------------------------- the refusal

def dtd_refusal(doc: str, part: str):
    """The UNSUPPORTED_CONTENT refusal for a part that declares a DTD."""
    from .errors import UnsafePackage

    return UnsafePackage(
        f"{doc}: its part {part} contains a document type declaration "
        "(<!DOCTYPE ...>), which an Excel workbook never has (the Open XML package rules "
        "do not allow one). It was not read: such a declaration can make an XML reader "
        "expand text without limit or read other files. Nothing was changed. If another "
        "program made this file, save it again from that program as an Excel workbook "
        "(.xlsx).")


def is_amplification(exc: BaseException) -> bool:
    """libxml2's stop for an entity that expands past its amplification
    limit, or that refers to itself (measured with lxml 6.1.3 / libxml2
    2.11.9: "Maximum entity amplification factor exceeded", "Detected an
    entity reference loop"); only a DTD's entities reach either."""
    text = str(exc).lower()
    return "entity" in text and ("amplification" in text or "reference loop" in text)


def has_doctype(root) -> bool:
    """True when the document that `root` belongs to declared a DTD."""
    info = root.getroottree().docinfo
    return bool(info.doctype) or info.internalDTD is not None or info.externalDTD is not None


# --------------------------------------------------------- the prolog look

def byte_encoding(data: bytes) -> str:
    """The encoding family of an XML part's bytes, from its byte order mark
    or its first character ("<" or the BOM-less first byte pair):
    "utf-32-le", "utf-32-be", "utf-16-le", "utf-16-be", or "utf-8" for
    every ASCII-compatible encoding."""
    head = data[:4]
    if head.startswith(b"\xff\xfe\x00\x00") or head == b"<\x00\x00\x00":
        return "utf-32-le"
    if head.startswith(b"\x00\x00\xfe\xff") or head == b"\x00\x00\x00<":
        return "utf-32-be"
    if head.startswith(b"\xff\xfe") or head.startswith(b"<\x00"):
        return "utf-16-le"
    if head.startswith(b"\xfe\xff") or head.startswith(b"\x00<"):
        return "utf-16-be"
    return "utf-8"


def text_of(data: bytes) -> str:
    """The part's characters, decoded by its byte order mark or first bytes
    (undecodable bytes replaced)."""
    enc = byte_encoding(data)
    for bom in (b"\xef\xbb\xbf", b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff",
                b"\xff\xfe", b"\xfe\xff"):
        if data.startswith(bom) and (enc != "utf-8" or bom == b"\xef\xbb\xbf"):
            data = data[len(bom):]
            break
    return data.decode(enc, errors="replace")


_MARK = "<!DOCTYPE"


def _walk(text, lt, pi, pi_end, com, com_end, mark, complete: bool):
    """Walk the prolog (XML declaration, comments, processing instructions,
    white space) up to the root element's start tag. True when "<!DOCTYPE"
    comes first; False at the root element; None when `text` ends inside
    the prolog and is not the whole part."""
    i = 0
    while True:
        i = text.find(lt, i)
        if i < 0:
            return None if not complete else False
        if not complete and len(text) - i < len(mark):
            return None  # too little left to tell a declaration from the root
        if text.startswith(mark, i):
            return True
        if text.startswith(com, i):
            j = text.find(com_end, i + len(com))
        elif text.startswith(pi, i):
            j = text.find(pi_end, i + len(pi))
        else:
            return False  # the root element: no declaration before it
        if j < 0:
            return None if not complete else False
        i = j + 1


def _walk_declares(data: bytes, complete: bool):
    if byte_encoding(data) == "utf-8":
        return _walk(data, b"<", b"<?", b"?>", b"<!--", b"-->", _MARK.encode("ascii"), complete)
    return _walk(text_of(data), "<", "<?", "?>", "<!--", "-->", _MARK, complete)


class _Doctype(Exception):
    pass


class _Root(Exception):
    pass


class _PrologTarget:
    """A parser target that stops at the first thing after the prolog: the
    document type declaration (before its internal subset is read) or the
    root element's start tag."""

    def doctype(self, *_args):
        raise _Doctype

    def start(self, *_args):
        raise _Root

    def close(self):
        return None


def _libxml2_says(head: bytes, complete: bool):
    """True / False as libxml2 reads the prolog; None when the head ends
    inside the prolog or libxml2 cannot read the prolog at all. huge_tree
    only lets this look read past a long comment: it builds no tree and
    expands nothing, stopping where the internal subset would begin."""
    p = etree.XMLParser(target=_PrologTarget(), **_checked({"huge_tree": True}))
    try:
        p.feed(head)
        if complete:
            p.close()
    except _Doctype:
        return True
    except _Root:
        return False
    except etree.XMLSyntaxError:
        return "error"
    return None


def _verdict(head: bytes, complete: bool):
    said = _libxml2_says(head, complete)
    if said == "error":
        return _walk_declares(head, complete)
    return said


def declares_dtd(data: bytes) -> bool:
    """True when the part declares a DTD in its prolog. Reads only as much
    of the part as its prolog takes."""
    n = _HEAD
    while True:
        complete = n >= len(data)
        said = _verdict(data[:n], complete)
        if said is not None or complete:
            return said is True
        n *= 8


def stream_declares_dtd(fh) -> bool:
    """declares_dtd for a part read from a stream (a zip member): reads its
    prolog and no further."""
    buf = b""
    want = _HEAD
    while True:
        more = fh.read(want - len(buf))
        buf += more
        complete = len(buf) < want or not more
        said = _verdict(buf, complete)
        if said is not None or complete:
            return said is True
        want *= 8


# ------------------------------------------------------------- the parsing

def _syntax(exc: BaseException) -> ParseError:
    """A part that is not well-formed, as the ParseError these parse sites
    raised before (same exception class, so the same refusal code)."""
    err = ParseError(str(exc))
    pos = getattr(exc, "position", None)
    if pos is not None:
        err.position = pos
    return err


def fromstring(data: bytes | str, *, doc: str = "the workbook", part: str = "an XML part",
               **overrides):
    """Parse one part's bytes with the hardened parser, as ElementTree read
    them (no comments or processing instructions in the tree). A part that
    declares a DTD is refused; one that is not well-formed raises
    ParseError. A str (markup this server builds) is parsed as UTF-8."""
    if isinstance(data, str):
        data = data.encode("utf-8")
    if declares_dtd(data):
        raise dtd_refusal(doc, part)
    try:
        root = etree.fromstring(data, parser=parser(**{**_ET_LAYOUT, **overrides}))
    except etree.XMLSyntaxError as exc:
        if is_amplification(exc):
            raise dtd_refusal(doc, part) from None
        raise _syntax(exc) from None
    if has_doctype(root):
        raise dtd_refusal(doc, part)
    return root


class _WalkTarget:
    def __init__(self, start, end, data=None):
        self._start, self._end, self._data = start, end, data

    def doctype(self, *_args):
        raise _Doctype

    def start(self, tag, attrib):
        if self._start is not None:
            self._start(tag, attrib)

    def end(self, tag):
        if self._end is not None:
            self._end(tag)

    def data(self, text):
        if self._data is not None:
            self._data(text)

    def close(self):
        return None


def walk(source, *, start=None, end=None, data=None, doc: str = "the workbook",
         part: str = "an XML part") -> None:
    """A streaming parse that builds no tree: start(tag, attrib) and
    end(tag) are called for each element (tags in "{namespace}local" form),
    and data(text) for each run of character data.
    `source` is bytes or a binary stream. A part that declares a DTD is
    refused; one that is not well-formed raises ParseError after the
    callbacks for everything before the fault."""
    p = etree.XMLParser(target=_WalkTarget(start, end, data), **_checked({}))
    try:
        if isinstance(source, (bytes, bytearray, memoryview)):
            view = memoryview(source)
            for i in range(0, len(view), 1 << 20):
                p.feed(bytes(view[i:i + (1 << 20)]))
        else:
            while True:
                chunk = source.read(1 << 20)
                if not chunk:
                    break
                p.feed(chunk)
        p.close()
    except _Doctype:
        raise dtd_refusal(doc, part) from None
    except etree.XMLSyntaxError as exc:
        if is_amplification(exc):
            raise dtd_refusal(doc, part) from None
        raise _syntax(exc) from None


def stream_elements(data: bytes, tags, *, doc: str = "the workbook",
                    part: str = "an XML part"):
    """Each element of one part whose "{namespace}local" tag is in `tags`, in
    document order, as it ends (lxml's iterparse with the hardened settings,
    run in C: only the matching elements reach Python). The caller may
    clear() a matching element once read, and delete its earlier siblings, to
    keep a large part's memory flat. A part that declares a DTD is refused
    before it is parsed; one that is not well-formed raises ParseError after
    the elements before the fault."""
    import io
    if declares_dtd(data):
        raise dtd_refusal(doc, part)
    try:
        for _event, el in etree.iterparse(io.BytesIO(data), events=("end",),
                                          tag=tuple(tags), **_checked(_ET_LAYOUT)):
            yield el
    except etree.XMLSyntaxError as exc:
        if is_amplification(exc):
            raise dtd_refusal(doc, part) from None
        raise _syntax(exc) from None


class _Rejoined:
    """A part's byte stream after the prolog look: the bytes the look read
    come first, then the rest of the stream. A part that declares a DTD is
    refused here, before any of it is parsed."""

    def __init__(self, source, doc: str, part: str):
        if isinstance(source, (bytes, bytearray, memoryview)):
            data = bytes(source)
            if declares_dtd(data):
                raise dtd_refusal(doc, part)
            self._head, self._rest = data, None
            return
        buf = b""
        want = _HEAD
        while True:
            more = source.read(want - len(buf))
            buf += more
            complete = len(buf) < want or not more
            said = _verdict(buf, complete)
            if said is not None or complete:
                break
            want *= 8
        if said is True:
            raise dtd_refusal(doc, part)
        self._head, self._rest = buf, (None if complete else source)

    def read(self, n: int = -1) -> bytes:
        if self._head:
            if n is None or n < 0:
                out, self._head = self._head, b""
                return out + (self._rest.read() if self._rest is not None else b"")
            out, self._head = self._head[:n], self._head[n:]
            return out
        if self._rest is None:
            return b""
        return self._rest.read(n)


def iterparse(source, *, doc: str = "the workbook", part: str = "an XML part"):
    """The standard library's iterparse for one part ("end" events, no
    comments or processing instructions), after the prolog look: a part that
    declares a DTD is refused before any of it is parsed, so the parser never
    meets an entity declaration (the guarantee openpyxl's worksheet and
    shared-string reads get from _CheckedArchive). `source` is bytes or a
    binary stream; the caller may clear what it has read. A part that is not
    well-formed raises ParseError after the elements before the fault."""
    from xml.etree.ElementTree import iterparse as _iterparse

    return _iterparse(_Rejoined(source, doc, part))
# ------------------------------------------------------------- the package

def _member(info) -> str:
    return getattr(info, "filename", info)


def check_package(zf: zipfile.ZipFile, *, doc: str) -> None:
    """Refuse the package when its directory shows a decompression bomb or a
    part named twice (zipsafe.check_members: nothing inflated), or when any
    of its XML parts (.xml, .rels, .vml) declares a DTD. Reads only each XML
    part's prolog. A member zipfile cannot read is left to the reader that
    meets it, which fails on it as it did before this check."""
    _zipsafe.check_members(zf, doc=doc)
    for info in zf.infolist():
        if not info.filename.lower().endswith(_XML_SUFFIXES) or info.is_dir():
            continue
        try:
            with zf.open(info) as fh:
                found = stream_declares_dtd(fh)
        except (zipfile.BadZipFile, zlib.error, RuntimeError, NotImplementedError,
                EOFError, OSError, ValueError):
            continue
        if found:
            raise dtd_refusal(doc, info.filename)


def check_file(path: str, *, doc: str | None = None) -> None:
    """check_package for the workbook file at `path`. A file that is not a
    zip is left to the caller's own open, which reports it as before."""
    try:
        zf = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError, ValueError):
        return
    with zf:
        check_package(zf, doc=doc or path)


class _CheckedArchive:
    """openpyxl's view of the package: each part it reads is checked for a
    document type declaration before openpyxl sees a byte of it. Everything
    else is the zip file itself."""

    def __init__(self, zf: zipfile.ZipFile, doc: str):
        self._zf = zf
        self._doc = doc
        self._names = None

    def read(self, name, pwd=None):
        data = self._zf.read(name, pwd)
        if declares_dtd(data):
            raise dtd_refusal(self._doc, _member(name))
        if isinstance(name, str) and name.lower().endswith(".rels") and b"%" in data:
            # A Target spelled with a percent code for a plain character
            # (worksheets/sheet%31.xml) names the part sheet1.xml; openpyxl
            # looks the literal text up and silently skips the sheet.
            if self._names is None:
                self._names = frozenset(self._zf.namelist())
            data = _partnames.respell_rels(name, data, self._names)
        return data

    def open(self, name, mode="r", pwd=None, **kwargs):
        if mode == "r":
            with self._zf.open(name, "r", pwd=pwd) as head:
                if stream_declares_dtd(head):
                    raise dtd_refusal(self._doc, _member(name))
        return self._zf.open(name, mode, pwd=pwd, **kwargs)

    def __getattr__(self, attr):
        return getattr(self._zf, attr)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self._zf.close()


#: Extensions of a macro-enabled workbook or template: the save keeps the VBA
#: project and the macro-enabled content type for these.
MACRO_EXTENSIONS = (".xlsm", ".xltm")


def keeps_vba(path) -> bool:
    """True when a workbook with this name is macro-enabled (.xlsm, .xltm), so
    it is loaded with keep_vba and saved with the macro-enabled content type."""
    return os.fspath(path).lower().endswith(MACRO_EXTENSIONS)


def load_workbook(path, *, read_only: bool = False, keep_vba: bool = False,
                  data_only: bool = False, keep_links: bool = True,
                  rich_text: bool = False):
    """openpyxl's load_workbook, with every part it reads checked first
    (_CheckedArchive), so openpyxl never parses a document type declaration,
    and the package's directory checked before openpyxl inflates anything
    (zipsafe.check_members). The only way this server builds an openpyxl
    model from a file."""
    from openpyxl.reader.excel import ExcelReader

    reader = ExcelReader(path, read_only=read_only, keep_vba=keep_vba, data_only=data_only,
                         keep_links=keep_links, rich_text=rich_text)
    archive = reader.archive
    doc = os.fspath(path) if not hasattr(path, "read") else "the workbook"
    reader.archive = _CheckedArchive(archive, doc)
    try:
        _zipsafe.check_members(archive, doc=doc)
        try:
            reader.read()
        except KeyError:
            # a listed sheet whose relationship id names nothing: openpyxl
            # raises a bare KeyError; the refusal names the sheet instead
            if not hasattr(path, "read"):
                class _None:
                    sheetnames: list = []
                refuse_unloaded_sheets(_None, os.fspath(path))
            raise
    except BaseException:
        archive.close()
        raise
    if not hasattr(path, "read"):
        refuse_unloaded_sheets(reader.wb, os.fspath(path))
    return reader.wb


_R_NAMESPACES = ("http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                 "http://purl.oclc.org/ooxml/officeDocument/relationships")


def refuse_unloaded_sheets(wb, path: str) -> None:
    """A workbook that lists a sheet the reader could not load is refused by
    name, never answered as a workbook with fewer sheets (or none). openpyxl
    skips, without a word, a sheet whose relationship names no part of the
    package; a read then reported ok with 0 sheets and an edit wrote the rest
    back without it. Two cases: a listed sheet whose target names no part
    (spelled so Excel itself does not find it), and a workbook that lists
    sheets and loaded none."""
    from . import opcxml as _opc
    from . import partnames as _pn
    from .errors import UnsupportedStructure
    from pathlib import Path as _P
    try:
        zf = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError, ValueError):
        return
    try:
        with zf:
            names = {n.lower() for n in zf.namelist()}
            names |= {_pn.normalize(n).lower() for n in zf.namelist()}
            wbpart = "xl/workbook.xml"
            if "_rels/.rels" in names:
                for rel in _opc.relationships(zf.read("_rels/.rels"), doc=_P(path).name,
                                              part="_rels/.rels"):
                    if rel.type.endswith("/officeDocument") and not rel.external:
                        wbpart = _opc.resolve_part("", rel.target)
                        break
            real = next((n for n in zf.namelist() if n.lower() == wbpart.lower()), None)
            if real is None:
                return
            root = fromstring(zf.read(real), doc=_P(path).name, part=real)
            rels_name = _opc.rels_part_for(real)
            real_rels = next((n for n in zf.namelist()
                              if n.lower() == rels_name.lower()), None)
            rels = {r.id: r for r in _opc.relationships(
                zf.read(real_rels), doc=_P(path).name, part=real_rels)} if real_rels else {}
            listed = []
            for el in root.iter():
                if not isinstance(el.tag, str):
                    continue
                ns, _s, local = (el.tag[1:].partition("}") if el.tag.startswith("{")
                                 else ("", "", el.tag))
                if local != "sheet" or ns not in _opc.MAIN_NAMESPACES:
                    continue
                rid = next((el.get("{%s}id" % n) for n in _R_NAMESPACES
                            if el.get("{%s}id" % n)), None)
                listed.append((el.get("name") or "", rid))
    except XML_ERRORS:
        return
    except (KeyError, OSError, ValueError, zipfile.BadZipFile):
        return
    # Every <sheet> entry counts, fail closed: one whose relationship id is
    # missing or empty, names no relationship, or names a part that is not in
    # the package is a sheet Excel cannot open, and a read must not show the
    # workbook without it.
    bad = []
    expected_loaded = 0
    for title, rid in listed:
        rel = rels.get(rid) if rid else None
        if not rid:
            bad.append((title, "it has no relationship id"))
        elif rel is None or rel.external:
            bad.append((title, f"its relationship id {rid!r} names no relationship "
                               "to a part of the package"))
        else:
            target = _opc.resolve_part(wbpart, rel.target)
            if target.lower() not in names:
                bad.append((title, f"its relationship target {rel.target!r} names "
                                   "no part of the package"))
            elif rel.type.endswith(("/worksheet", "/chartsheet")):
                expected_loaded += 1
    if bad:
        title, why = bad[0]
        more = f" (and {len(bad) - 1} more)" if len(bad) > 1 else ""
        raise UnsupportedStructure(
            f"{_P(path).name}: the workbook lists the sheet {title!r}, but {why}{more}. "
            "Excel does not open a workbook whose sheet has no part, and a read here "
            "would show a workbook with sheets missing, so the file was refused and "
            "not changed. Save it again from the program that made it, or repair it "
            "in Excel, then retry")
    if listed and (not wb.sheetnames or len(wb.sheetnames) < expected_loaded):
        raise UnsupportedStructure(
            f"{_P(path).name}: the workbook lists {len(listed)} sheet(s) but "
            + ("none" if not wb.sheetnames else f"only {len(wb.sheetnames)}")
            + " could be loaded, so the file was refused rather than read as a "
            "workbook with sheets missing; nothing was changed. Save it again from "
            "the program that made it, or repair it in Excel, then retry")


def install_string_table(read_string_table) -> None:
    """Give openpyxl's workbook reader this server's shared-string table
    reader (core/xltext.py, which parses the table with iterparse here).
    openpyxl's reader modules are imported in this module only."""
    from openpyxl.reader import excel as _excel
    from openpyxl.reader import strings as _strings

    _strings.read_string_table = read_string_table
    _excel.read_string_table = read_string_table


__all__ = [
    "XML_ERRORS", "ParseError", "parser", "fromstring", "walk", "iterparse", "declares_dtd",
    "stream_declares_dtd", "has_doctype", "is_amplification", "dtd_refusal",
    "byte_encoding", "text_of", "check_package", "check_file", "load_workbook",
    "install_string_table",
]
