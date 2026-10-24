"""core/netreach.py: what in a workbook can reach the network when Excel
calculates it.

A recalculation in Excel runs every formula, and a few worksheet functions
fetch from the network as they calculate: WEBSERVICE (often inside FILTERXML),
IMAGE, STOCKHISTORY, TRANSLATE, DETECTLANGUAGE, COPILOT and Python in Excel
(PY). RTD asks a real-time data server that may sit on another machine, the
CUBE functions query an OLAP source, and an external data connection (Power
Query, a web or database query) can refresh from its server. None of that is
a macro, so Excel's automation security does not stop it.

The COM tier therefore refuses to calculate such a workbook unless the call
passes allow_network=true: this server never makes a network request the
caller did not ask for. The check reads the file only and never opens the
workbook in Excel.

How it reads, and why it fails closed:

- XML is read with a real XML parser (lxml with the hardened settings of
  core.xmlsafe: no DTD, no entities, no network), never with patterns over
  the markup, so character references, CDATA sections, comments and
  processing instructions mean exactly what the XML rules say. Every text
  node and every attribute value of every element is searched. The raw text
  of each part is searched as well, so a call found in either reading counts.
- EVERY member of the package is read, not only the parts its relationships
  reach, so a decoy workbook part or a part named in another letter case
  hides nothing.
- Each member's reader is chosen by its content type in [Content_Types].xml
  (an Override for its name first, then the Default for its extension),
  never by its name or its first bytes. Every member of an XML type goes
  through the parser and the raw-text search. Only a member of a known
  binary type (a picture, a font, a printer setting, a VBA project, an
  embedded object) is not parsed, and its bytes are still searched. Excel
  loads a part through the relationship that reaches it, whatever its
  content type says (measured), so a part of another type that a
  relationship loads as XML content is parsed too, and counts as unread.
- Whenever the file cannot be read with certainty, the answer is "may reach
  the network": a member that cannot be read (a checksum error, an encrypted
  or unknown compression), an XML part that is not well-formed or that
  declares an encoding other than UTF-8 or UTF-16, a member with no content
  type or one this check does not know, a part named twice once letter case
  and escapes are ignored, a relationship that cannot be resolved, and a file
  too large to read in full. Such a file is refused
  unless allow_network, with the reason in `unread`. A partial reading is
  never reported as clean.

In an .xlsx or .xlsm package a formula (a cell, a defined name, a
conditional-format or validation rule and the other formula places) counts a
call by the name Excel stores (calls_in: _xlfn. and _xlws. prefixes; a
function name inside a text literal or after a dot is not a call, measured).
Every other text node, attribute value and the raw text are searched for the
spellings Excel calculates in a package (_STORED_CALL_RE). A CSV, an
OpenDocument spreadsheet and other text are searched for every spelling
(_TEXT_CALL_RE), the name after any prefix or namespace.
"""

from __future__ import annotations

import codecs
import os
import posixpath
import re
import zipfile
import zlib
from pathlib import Path
from urllib.parse import unquote

from . import xmlsafe as _xmlsafe

#: Functions that can reach the network when Excel calculates them, by the
#: name a formula stores (newer ones carry the _xlfn. prefix in the file).
NETWORK_FUNCTIONS = ("WEBSERVICE", "IMAGE", "STOCKHISTORY", "TRANSLATE",
                     "DETECTLANGUAGE", "COPILOT", "PY", "RTD", "CUBEVALUE",
                     "CUBEMEMBER", "CUBESET", "CUBESETCOUNT", "CUBERANKEDMEMBER",
                     "CUBEMEMBERPROPERTY", "CUBEKPIMEMBER")
#: The functions Excel added after 2007: a package stores them as _xlfn.NAME
#: (PY as _xlfn._xlws.PY). The older ones (RTD, the CUBE functions) are
#: stored bare.
_NEWER = ("WEBSERVICE", "IMAGE", "STOCKHISTORY", "TRANSLATE", "DETECTLANGUAGE", "COPILOT", "PY")
_OLDER = tuple(f for f in NETWORK_FUNCTIONS if f not in _NEWER)

_CALL_RE = re.compile(
    r"(?<![A-Za-z0-9_.\\])((?:_xlfn\.)?(?:_xlws\.)?)(" + "|".join(NETWORK_FUNCTIONS)
    + r")\s*\(", re.IGNORECASE)
#: The same functions in a TEXT file (a CSV, an OpenDocument spreadsheet's
#: XML). A name may follow a namespace or prefix there: OpenDocument stores
#: WEBSERVICE as COM.MICROSOFT.WEBSERVICE (measured: Excel's own SaveAs to
#: .ods, and Excel fetches from it), and _xlfn. or _xlws. may sit before a
#: name anywhere. So only a letter, digit or underscore right before the
#: name rules it out. PY is two letters and ordinary text holds ".py(" often
#: enough, so it counts bare, or after _xlws. or MICROSOFT. only.
_TEXT_CALL_RE = re.compile(
    r"(?<![A-Za-z0-9_])(" + "|".join(f for f in NETWORK_FUNCTIONS if f != "PY") + r")\s*\("
    r"|(?<![A-Za-z0-9_.])(PY)\s*\("
    r"|(?<![A-Za-z0-9_])(?:_xlws|MICROSOFT)\.(PY)\s*\(", re.IGNORECASE)
#: The spellings Excel calculates in an .xlsx or .xlsm package, wherever the
#: text sits: a newer function after _xlfn. or _xlws. (whatever comes before
#: the prefix), an older one bare or after _xlfn. (measured in Excel 16: a
#: bare WEBSERVICE, or one after COM.MICROSOFT. or another dotted prefix, is
#: an unknown name there and fetches nothing). Ordinary text such as
#: "image (1)" in a cell is therefore not taken for a call.
_STORED_CALL_RE = re.compile(
    r"_xl(?:fn|ws)\.(?:_xlws\.)?(" + "|".join(_NEWER) + r")\s*\("
    r"|(?<![A-Za-z0-9_.])(?:_xlfn\.)?(" + "|".join(_OLDER) + r")\s*\(", re.IGNORECASE)
_NUM_REF_RE = re.compile(r"&#(?:[xX]0*([0-9A-Fa-f]{1,6})|0*([0-9]{1,7}));")
#: Text that starts like an XML document (a declaration, or a root element).
_XML_START_RE = re.compile(r"\s*(?:<\?xml\b|<[A-Za-z_][\w.\-]*(?::[\w.\-]+)?[\s/>])")
_MARKUP_ENTITY = {38: "&amp;", 60: "&lt;", 62: "&gt;", 34: "&quot;", 39: "&apos;"}

#: The SpreadsheetML namespaces (transitional and strict).
_SML = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "http://purl.oclc.org/ooxml/spreadsheetml/main")
_REL_NS = ("http://schemas.openxmlformats.org/package/2006/relationships",)
_R_NS = ("http://schemas.openxmlformats.org/officeDocument/2006/relationships",
         "http://purl.oclc.org/ooxml/officeDocument/relationships")
#: Elements whose text Excel reads as a formula, and attributes likewise.
_RULE_TAGS = frozenset({"formula", "formula1", "formula2"})
_FORMULA_TAGS = frozenset({"f", "definedName", "calculatedColumnFormula", "totalsRowFormula"}) | _RULE_TAGS
#: Member names read as XML in a ZIP that is not an Excel package (an
#: OpenDocument spreadsheet, whose reader takes its parts by fixed names).
_XML_SUFFIXES = (".xml", ".rels", ".vml")

#: The list of an Excel package's content types, and its namespace.
_CT_PART = "[Content_Types].xml"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_CT_ITER = (f"{{{_CT_NS}}}Default", f"{{{_CT_NS}}}Override")
#: The workbook part of an .xlsb, which Excel has to look at itself.
_XLSB_MAIN = "application/vnd.ms-excel.sheet.binary.macroenabled.main"
#: The closed list of content types that hold no XML: pictures and other
#: media, fonts, printer settings, VBA projects and their signatures, ActiveX
#: and OLE objects, embedded documents, signature certificates and the
#: other binary parts Excel writes. A member of one of these types is not
#: parsed; its bytes are still searched.
_BINARY_FAMILIES = ("image/", "audio/", "video/", "font/")
_BINARY_TYPES = frozenset({
    "application/vnd.openxmlformats-officedocument.spreadsheetml.printersettings",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.customproperty",
    "application/vnd.openxmlformats-officedocument.oleobject",
    "application/vnd.openxmlformats-officedocument.obfuscatedfont",
    "application/vnd.openxmlformats-package.digital-signature-certificate",
    "application/vnd.openxmlformats-package.digital-signature-origin",
    "application/vnd.ms-package.obfuscated-opentype",
    "application/vnd.ms-office.vbaproject",
    "application/vnd.ms-office.vbaprojectsignature",
    "application/vnd.ms-office.vbaprojectsignatureagile",
    "application/vnd.ms-office.vbaprojectsignaturev3",
    "application/vnd.ms-office.activex",
    "application/vnd.ms-excel.attachedtoolbars",
    "application/vnd.ms-excel.datamodel", "application/vnd.openxmlformats-officedocument.model+data",
    "application/x-font-ttf", "application/x-fontdata",
    # embedded documents
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
    "application/vnd.openxmlformats-officedocument.presentationml.template",
    "application/vnd.ms-excel", "application/vnd.ms-excel.sheet.macroenabled.12",
    "application/vnd.ms-excel.sheet.binary.macroenabled.12",
    "application/vnd.ms-excel.template.macroenabled.12",
    "application/msword", "application/vnd.ms-word.document.macroenabled.12",
    "application/vnd.ms-powerpoint", "application/vnd.ms-powerpoint.presentation.macroenabled.12",
    "application/vnd.ms-powerpoint.slideshow.macroenabled.12",
    "application/pdf",
})
#: Families of the Open XML types that are XML whatever their suffix
#: (SpreadsheetML, DrawingML and the package's own; VML is matched by name).
#: The relationship types (by their last segment, letter case ignored) that
#: reach binary content: pictures and media, printer settings, VBA projects
#: and their signatures, OLE and ActiveX objects, embedded packages, fonts,
#: thumbnails, signature certificates and origins.
_BINARY_RELS = frozenset({
    "image", "hdphoto", "media", "audio", "video", "printersettings", "vbaproject",
    "vbaprojectsignature", "vbaprojectsignatureagile", "vbaprojectsignaturev3", "oleobject",
    "package", "font", "activexcontrolbinary", "attachedtoolbars", "thumbnail", "certificate",
    "origin", "customproperty", "model", "powerpivotdata"})
_XML_FAMILIES = ("application/vnd.openxmlformats-officedocument.spreadsheetml.",
                 "application/vnd.openxmlformats-officedocument.drawingml.",
                 "application/vnd.openxmlformats-package.")


def _type_kind(content_type: str | None) -> str:
    """How a member of this content type is read: "xml" (the parser and the
    raw-text search), "binary" (a type on the closed list above: its bytes
    are searched) or "unknown" (no type, or one this check does not know:
    the file cannot be read with certainty)."""
    if not content_type:
        return "unknown"
    t = content_type.split(";", 1)[0].strip().lower()
    if t.endswith(("+xml", "/xml")) or "vml" in t:
        return "xml"
    if t in _BINARY_TYPES or t.startswith(_BINARY_FAMILIES):
        return "binary"
    if t.startswith(_XML_FAMILIES):
        return "xml"
    return "unknown"

#: How many places the refusal names before it counts the rest.
_SHOW = 8
#: The most a file may hold, in all its members as read, for the check to
#: read it in full. A larger file counts as one that may reach the network.
_MAX_READ = 1024 * 1024 * 1024
_CHUNK = 1024 * 1024
_TEXT_CHUNK = 8 * 1024 * 1024
_OVERLAP = 64

#: What file_reach says about a file: its formulas were read from the
#: package (an .xlsx, .xlsm and their templates), from its text (a CSV or
#: other text file, an OpenDocument spreadsheet), or cannot be read here
#: (an .xlsb, a legacy .xls, an encrypted workbook), so Excel has to look.
READ_PACKAGE, READ_TEXT, READ_IN_EXCEL = "package", "text", "excel"

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
#: What reading a ZIP member can raise: a checksum error (BadZipFile), a
#: damaged stream, an encrypted member (RuntimeError), an unknown compression.
_READ_ERRORS = (zipfile.BadZipFile, zlib.error, EOFError, OSError, RuntimeError,
                NotImplementedError, ValueError, KeyError)


class _Unread(Exception):
    """The check cannot read this with certainty; the message says why."""


def _num_ref(m) -> str:
    code = int(m.group(1), 16) if m.group(1) else int(m.group(2))
    if code in _MARKUP_ENTITY:
        return _MARKUP_ENTITY[code]
    if code == 0 or code > 0x10FFFF or 0xD800 <= code <= 0xDFFF:
        return m.group(0)
    return chr(code)


def _views(text: str, xml: bool = True) -> list[str]:
    """The ways the raw-text search reads a text, the text itself first: the
    text as it is, and the text with character references resolved. A view
    only adds places to look. Nothing is ever removed from the text (the XML
    itself is read by the parser, not here). `xml` is kept for callers; both
    kinds of text are read the same way."""
    out = [text]
    if "&#" in text:
        resolved = _NUM_REF_RE.sub(_num_ref, text)
        if resolved != text:
            out.append(resolved)
    return out


#: A text can hold a call only when it holds a function name: a quick test
#: before the patterns run. Characters that Python's case-insensitive
#: matching takes for ASCII letters are folded first, so the test never
#: rules out a text the pattern would match.
_LOW_NAMES = tuple(n.lower() for n in NETWORK_FUNCTIONS if n != "PY")
_PY_RE = re.compile(r"py\s*\(")
_FOLD = str.maketrans({"\u0130": "i", "\u0131": "i", "\u017f": "s", "\u212a": "k"})


def _may_name(text: str) -> bool:
    """False only when no pattern here can match anywhere in text."""
    low = text.lower() if text.isascii() else text.translate(_FOLD).lower()
    return any(n in low for n in _LOW_NAMES) or ("py" in low and _PY_RE.search(low) is not None)


def _names(pattern: re.Pattern, text: str, found: list[str]) -> None:
    if not _may_name(text):
        return
    for m in pattern.finditer(text):
        name = next(g for g in m.groups() if g).upper()
        if name not in found:
            found.append(name)


def calls_in(formula: str) -> list[str]:
    """The network-capable functions a formula calls, in order, each once
    (PY only in its stored form _xlfn._xlws.PY, so a function of one's own
    named PY is not taken for it)."""
    from .calc import _blank_literals
    found: list[str] = []
    for m in _CALL_RE.finditer(_blank_literals(formula)):
        name = m.group(2).upper()
        if name == "PY" and "_xlws." not in m.group(1).lower():
            continue
        if name not in found:
            found.append(name)
    return found


def _text_calls(chunks, xml: bool | None = None, pattern: re.Pattern = _TEXT_CALL_RE) -> list[str]:
    """The network-capable functions named anywhere in a text, read in chunks
    (a name cut by a chunk edge is seen in the overlap), in every view of
    _views. Nothing that looks like markup is stripped: a CSV is only text,
    and XML is read by the parser (_XmlScan) besides this raw search."""
    found: list[str] = []
    tail = ""
    for text in chunks:
        window = tail + text
        for view in _views(window, bool(xml)):
            _names(pattern, view, found)
        tail = window[-_OVERLAP:]
    return found


# ------------------------------------------------------------ reading XML

def _local(tag) -> str:
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1] if tag.startswith("{") else tag.rsplit(":", 1)[-1]


#: lxml's own tag iteration ("{*}" is any namespace or none) finds these
#: elements in C; an XPath predicate on local-name() visits every element
#: in Python-sized steps and cost several times the parse on large sheets.
_FORMULA_ITER = tuple("{*}" + t for t in sorted(_FORMULA_TAGS))
_CONNECTION_ITER = tuple(f"{{{ns}}}connection" for ns in _SML)
_SHEET_ITER = tuple(f"{{{ns}}}sheet" for ns in _SML)
_REL_ITER = tuple(f"{{{ns}}}Relationship" for ns in _REL_NS)
_X_TEXT = _xmlsafe.xpath("descendant-or-self::*/text()", smart_strings=False)
_X_TEXT_NODES = _xmlsafe.xpath("descendant-or-self::*/text()")
_X_ATTR = _xmlsafe.xpath("descendant-or-self::*/@*", smart_strings=False)
_X_OWN_TEXT = _xmlsafe.xpath("text()", smart_strings=False)
_X_FORMULA_ATTR = _xmlsafe.xpath("descendant-or-self::*/@formula", smart_strings=False)
#: Elements read in batches as the parser reaches them and then let go, so
#: a large part never sits in memory whole: rows, shared strings, pivot
#: cache records and OpenDocument table rows, under the parents named here.
_BATCH_TAGS = (f"{{{_SML[0]}}}row", f"{{{_SML[1]}}}row", f"{{{_SML[0]}}}si", f"{{{_SML[1]}}}si",
               f"{{{_SML[0]}}}r", f"{{{_SML[1]}}}r",
               "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-row")
_BATCH_PARENTS = frozenset({"sheetData", "sst", "pivotCacheRecords", "table", "table-rows"})
_BATCH = 512


class _XmlScan:
    """The parser's view of one XML part: every text node and every
    attribute value of every element is searched with `pattern` (an
    element's text is also searched joined across any child elements inside
    it); formula places are read with calls_in and named. kind is "sheet"
    (title = the sheet's name), "workbook" or "other". The parse is lxml's
    with the hardened settings (no DTD, no entities, no network); comments
    and processing instructions are dropped and CDATA is plain text, as the
    XML rules have it."""

    def __init__(self, member: str, pattern: re.Pattern, kind: str = "other",
                 title: str = "", formulas: bool = True):
        self.member, self.pattern, self.kind, self.title = member, pattern, kind, title
        self.formulas = formulas
        self.places: list[tuple[str, list[str]]] = []  # (where, functions) in order
        self.loose: list[str] = []  # names found anywhere (formula places included)
        self.connections: list[str] = []
        self.rels: list[dict] = []  # Relationship elements (a .rels part)
        self.sheets: list[dict] = []  # <sheet> elements (the workbook part)
        self.types: list[tuple[str, dict]] = []  # Default/Override ([Content_Types].xml)

    def parse(self, source) -> None:
        """Parse `source` (a file-like object). Raises _Unread or lxml's
        error for a part that is not well-formed or declares a DTD; what was
        read before the fault has been searched."""
        ctx = _xmlsafe.stream_parse(source, events=("end",), tag=_BATCH_TAGS,
                                    remove_comments=True, remove_pis=True)
        counts: dict = {}
        root = None
        try:
            for _ev, el in ctx:
                parent = el.getparent()
                if parent is None or _local(parent.tag) not in _BATCH_PARENTS:
                    continue
                n = counts.get(parent, 0) + 1
                if n >= _BATCH:
                    self._search(parent)
                    del parent[:]
                    n = 0
                counts[parent] = n
        finally:
            root = ctx.root
            if root is not None:
                self._search(root)
        if root is None:
            raise _Unread("it is not well-formed XML (it has no root element)")
        if _xmlsafe.has_doctype(root):
            raise _Unread("it declares a document type, which no workbook part has")

    def _search(self, node) -> None:
        texts = _X_TEXT(node)
        if _may_name("".join(texts)):
            # each text node, then an element's text joined across the child
            # elements inside it (a non-blank tail marks one)
            _names(self.pattern, "\x00".join(texts), self.loose)
            joined = []
            for t in _X_TEXT_NODES(node):
                if t.is_tail and not t.isspace():
                    holder = t.getparent().getparent()
                    if holder is not None and holder not in joined:
                        joined.append(holder)
            for el in joined:
                _names(self.pattern, "".join(_X_OWN_TEXT(el)), self.loose)
            if self.formulas:
                for el in node.iter(*_FORMULA_ITER):
                    self._formula(el)
        attrs = _X_ATTR(node)
        if _may_name("".join(attrs)):
            _names(self.pattern, "\x00".join(attrs), self.loose)
            if self.formulas:
                for value in _X_FORMULA_ATTR(node):
                    if "(" in value:
                        self._place(f"a formula in the part {self.member}", calls_in(value))
        for el in node.iter(*_CONNECTION_ITER):
            name = el.get("name") or "(unnamed)"
            if name not in self.connections:
                self.connections.append(name)
        for el in node.iter(*_REL_ITER):
            self.rels.append({_local(k): v for k, v in el.attrib.items()})
        for el in node.iter(*_CT_ITER):
            self.types.append((_local(el.tag), dict(el.attrib)))
        if self.kind == "workbook":
            for el in node.iter(*_SHEET_ITER):
                self.sheets.append({("rid" if k.startswith("{") and k[1:].split("}")[0] in _R_NS
                                     and _local(k) == "id" else _local(k)): v
                                    for k, v in el.attrib.items()})

    # -- formula places
    def _quoted(self) -> str:
        t = self.title
        return t if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", t) else "'" + t.replace("'", "''") + "'"

    def _formula(self, el) -> None:
        text = "".join(el.itertext())
        if "(" not in text:
            return
        funcs = calls_in(text)
        if not funcs:
            return
        local = _local(el.tag)
        parent = el.getparent()
        if self.kind == "sheet" and local == "f" and parent is not None and _local(parent.tag) == "c":
            ref = el.get("ref") if el.get("t") in ("shared", "array") else None
            self._place(f"{self._quoted()}!{ref or parent.get('r', '?')}", funcs)
        elif self.kind == "sheet" and local in _RULE_TAGS | {"f"}:
            self._place(f"a conditional format or data validation rule on {self.title}", funcs)
        elif self.kind == "workbook" and local == "definedName":
            self._place(f"the defined name {el.get('name', '?')}", funcs)
        else:
            self._place(f"a formula in the part {self.member}", funcs)

    def _place(self, where: str, funcs: list[str]) -> None:
        if not funcs:
            return
        for w, fs in self.places:
            if w == where:
                fs.extend(f for f in funcs if f not in fs)
                return
        self.places.append((where, list(funcs)))


def _reason(exc: BaseException) -> str:
    text = " ".join(str(exc).split())
    if isinstance(exc, zipfile.BadZipFile) and "CRC" in text:
        return "its stored checksum does not match its content"
    if isinstance(exc, RuntimeError) and "encrypted" in text.lower():
        return "it is encrypted inside the ZIP package"
    if isinstance(exc, NotImplementedError):
        return "it is stored with a compression this server cannot read"
    return f"it could not be read ({text[:120] or type(exc).__name__})"


def _family_codec(family: str) -> str:
    return {"utf-8": "utf-8-sig", "utf-16-le": "utf-16", "utf-16-be": "utf-16"}[family]


_BOM16 = {"utf-16-le": b"\xff\xfe", "utf-16-be": b"\xfe\xff"}
#: How far into a part its XML declaration must end: 64 KiB. The declaration
#: is the first thing in a part, and every workbook writer makes it a few
#: dozen bytes long; the XML grammar lets its white space run on without
#: limit, so the encoding it names could sit past any chunk. The declaration
#: is read across chunks up to this bound. One that has not ended by then is
#: a doubt (the part is refused as not read with certainty), never a part
#: without a declaration.
_DECLARATION_BOUND = 64 * 1024
_DECLARATION_START = re.compile(r"<\?xml[ \t\r\n]")
_DECLARED_ENCODING = re.compile(
    r"""\bencoding[ \t\r\n]*=[ \t\r\n]*(["'])([A-Za-z][A-Za-z0-9._-]*)\1""")
_DECLARED_OK = {"utf-8": {"utf-8", "utf8"},
                "utf-16-le": {"utf-16", "utf16", "utf-16le", "utf16le"},
                "utf-16-be": {"utf-16", "utf16", "utf-16be", "utf16be"}}


def _declaration(pending: list[bytes], rest) -> tuple[str | None, str | None]:
    """(the encoding the part's XML declaration names, a doubt) from the
    part's first _DECLARATION_BOUND bytes: (None, None) when the part has no
    declaration or one that names no encoding; a doubt when a declaration
    starts but does not end within the bound, or names an encoding in a form
    that cannot be read. Reads the chunks already taken (pending) and, while
    they hold less than the bound, takes more from `rest` into pending, so
    the reads after this one still see every byte."""
    head, i = b"", 0
    while len(head) < _DECLARATION_BOUND:
        if i == len(pending):
            more = next(rest, None)
            if more is None:
                break
            pending.append(more)
        head += pending[i][:_DECLARATION_BOUND - len(head)]
        i += 1
    text = _xmlsafe.text_of(head)
    if not _DECLARATION_START.match(text):
        return None, None
    end = text.find("?>")
    if end < 0:
        return None, (f"its XML declaration does not end within its first "
                      f"{_DECLARATION_BOUND // 1024} KiB, so the encoding it names "
                      "cannot be read with certainty")
    found = _DECLARED_ENCODING.search(text, 0, end)
    if found:
        return found.group(2), None
    if "encoding" in text[:end]:
        return None, "its XML declaration names an encoding in a form that cannot be read"
    return None, None


def _syntax(exc: BaseException) -> str:
    if isinstance(exc, _Unread):
        return str(exc)
    if isinstance(exc, _xmlsafe.SYNTAX_ERROR) and _xmlsafe.is_amplification(exc):
        return "it declares a document type, which no workbook part has"
    return f"it is not well-formed XML ({' '.join(str(exc).split())[:120]})"


class _Source:
    """The bytes of one part for the parser, each chunk passed to `seen`
    (the raw-text search) as the parser takes it."""

    def __init__(self, pending: list[bytes], rest, seen):
        self._queue, self._rest, self._seen = list(pending), rest, seen
        self._buf, self._pos = b"", 0

    def _pull(self) -> bool:
        if self._queue:
            raw = self._queue.pop(0)
        else:
            raw = next(self._rest, None)
            if raw is None:
                return False
        self._seen(raw)
        self._buf, self._pos = self._buf[self._pos:] + raw, 0
        return True

    def read(self, n: int = -1) -> bytes:
        if n is None or n < 0:
            while self._pull():
                pass
            n = len(self._buf) - self._pos
        while len(self._buf) - self._pos < n and self._pull():
            pass
        out = self._buf[self._pos:self._pos + n]
        self._pos += len(out)
        return out

    def drain(self) -> None:
        """What the parser did not read (it stopped at a fault) still goes
        through the raw-text search."""
        while self._pull():
            self._buf, self._pos = b"", 0


def _read_xml(chunks, scan: _XmlScan, raw_pattern: re.Pattern) -> list[str]:
    """Read one XML part (an iterable of byte chunks) through the parser view
    (scan) and the raw-text search with raw_pattern, whose finds go into
    scan.loose. Returns why the part cannot be read with certainty, if it
    cannot: UTF-32, an encoding other than UTF-8 or UTF-16 declared, text
    that does not decode, XML that is not well-formed, a document type. Each
    doubt stops only the reading it concerns: the raw-text search always
    reads the whole part, and the parser reads as far as it can (in the
    declared encoding, whatever it is), so a call is still named where it
    can be. A ZIP read error propagates."""
    doubts: list[str] = []
    it = iter(chunks)
    head = b""
    for chunk in it:
        head += chunk
        if len(head) >= 4:
            break
    family = _xmlsafe.byte_encoding(head)
    prefix = b""
    if family in _BOM16 and not head.startswith(_BOM16[family]):
        prefix = _BOM16[family]  # UTF-16 without a mark, as XML allows: the parser gets the mark
    pending = [prefix + head]
    # the declaration, as the parser reads it (the prolog only)
    pull = _xmlsafe.pull_parser(events=("start",))
    root = None
    try:
        data, fed, step = pending[0], 0, 1 << 16
        while root is None:
            if fed >= len(data):
                nxt = next(it, None)
                if nxt is None:
                    pull.close()
                    for _ev, root in pull.read_events():
                        break
                    break
                pending.append(nxt)
                data, fed = nxt, 0
            pull.feed(data[fed:fed + step])
            fed += step
            for _ev, root in pull.read_events():
                break
    except _xmlsafe.LXML_ERROR as exc:
        doubts.append(_syntax(exc))
    said = root.getroottree().docinfo.encoding if root is not None else None
    del root, pull
    # The parser reports the declared encoding only once it has read the
    # whole part on some libxml2 versions, and not at all for an encoding it
    # does not support; the declaration itself is read as well, and either
    # one naming another encoding is a doubt.
    written, unsure = _declaration(pending, it)
    if unsure:
        doubts.append(unsure)
    codecs_ = [_family_codec(family) if family in _DECLARED_OK else family]
    declared = next((d for d in (said, written)
                     if d is not None and family in _DECLARED_OK
                     and d.lower() not in _DECLARED_OK[family]), None)
    if family not in _DECLARED_OK:
        doubts.append("it is stored as UTF-32 text, and a workbook part is UTF-8 or UTF-16")
    elif declared is not None:
        doubts.append(f"it declares the encoding {declared}, and a workbook part is UTF-8 "
                      "or UTF-16")
        try:
            codecs_.append(codecs.lookup(declared).name)  # also as declared, to name what it hides
        except LookupError:
            pass
    # the raw text, in the encoding its bytes show and in the one it declares: each view only adds
    readers = [[c, codecs.getincrementaldecoder(c)(errors="strict"), ""] for c in codecs_]

    def seen(raw: bytes, final: bool = False) -> None:
        for reader in readers:
            codec, decoder, tail = reader
            try:
                text = decoder.decode(raw, final)
            except UnicodeDecodeError:
                if codec == codecs_[0]:
                    doubts.append(f"its text does not decode as {codec.upper()}")
                decoder = reader[1] = codecs.getincrementaldecoder(codec)(errors="replace")
                text = decoder.decode(raw, final)
            window = tail + text
            _names(raw_pattern, window, scan.loose)
            reader[2] = window[-_OVERLAP:]

    source = _Source(pending, it, seen)
    try:
        scan.parse(source)
    except (_Unread, _xmlsafe.LXML_ERROR) as exc:
        doubts.append(_syntax(exc))
    source.drain()
    seen(b"", final=True)
    out: list[str] = []
    for d in doubts:
        if d not in out:
            out.append(d)
    return out


class _Member:
    """A ZIP member's bytes in chunks, read to the end even when its stored
    checksum is wrong (Excel does not check it, measured): the check is made
    here instead, and a mismatch is kept in crc_bad."""

    def __init__(self, zf: zipfile.ZipFile, info: zipfile.ZipInfo):
        self.zf, self.info, self.crc_bad = zf, info, False

    def chunks(self):
        crc = 0
        with self.zf.open(self.info) as fh:
            if hasattr(fh, "_expected_crc"):
                fh._expected_crc = None  # checked below, after every byte is read
            while True:
                raw = fh.read(_CHUNK)
                if not raw:
                    break
                crc = zlib.crc32(raw, crc)
                yield raw
        if crc != self.info.CRC:
            self.crc_bad = True

    def head(self, n: int = 4096) -> bytes:
        """The first bytes, without the checksum test (which needs every byte)."""
        with self.zf.open(self.info) as fh:
            if hasattr(fh, "_expected_crc"):
                fh._expected_crc = None
            return fh.read(n)

    def doubts(self) -> list[str]:
        return ["its stored checksum does not match its content"] if self.crc_bad else []


def _looks_like_xml(head: bytes) -> bool:
    if _xmlsafe.byte_encoding(head) != "utf-8":
        return True  # a "<" of UTF-16 or UTF-32 width, or their marks
    body = head[3:] if head.startswith(b"\xef\xbb\xbf") else head
    return body.lstrip(b" \t\r\n")[:1] == b"<"


# ------------------------------------------------------------ the package

def _norm(name: str) -> str:
    """A part name as the package rules compare it: without a leading slash,
    escapes decoded, letter case ignored."""
    return unquote(name.lstrip("/")).casefold()


def _resolve(source: str, target: str) -> str:
    """The part name a relationship's target names, from its source part."""
    target = target.split("#", 1)[0]
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


class _Reach:
    """What a file's reading found: places with their functions, data
    connections, and the reasons the reading is not certain."""

    def __init__(self):
        self.formulas: list[dict] = []
        self.connections: list[str] = []
        self.unread: list[str] = []

    def add(self, where: str, funcs: list[str]) -> None:
        if not funcs:
            return
        for f in self.formulas:
            if f["where"] == where:
                f["functions"].extend(fn for fn in funcs if fn not in f["functions"])
                return
        self.formulas.append({"where": where, "functions": list(funcs)})

    def doubt(self, why: str) -> None:
        if why not in self.unread:
            self.unread.append(why)

    def take(self, scan: _XmlScan, loose_where: str | None = None) -> None:
        named: set[str] = set()
        for where, funcs in scan.places:
            self.add(where, funcs)
            named.update(funcs)
        rest = [f for f in scan.loose if f not in named]
        self.add(loose_where or f"the part {scan.member}", rest)
        for c in scan.connections:
            if c not in self.connections:
                self.connections.append(c)

    def result(self) -> dict | None:
        if not (self.formulas or self.connections or self.unread):
            return None
        out = {"formulas": self.formulas, "connections": self.connections}
        if self.unread:
            out["unread"] = self.unread
        return out


class _Types:
    """A package's content types: Override by part name (letter case and
    escapes ignored), then Default by extension (letter case ignored)."""

    def __init__(self):
        self.over: dict[str, str] = {}
        self.default: dict[str, str] = {}

    def of(self, name: str) -> str | None:
        t = self.over.get(_norm(name))
        if t is None:
            seg = name.rsplit("/", 1)[-1]
            if "." in seg:
                t = self.default.get(seg.rsplit(".", 1)[1].casefold())
        return t


def _content_types(zf, info, reach: _Reach) -> tuple[_Types, _XmlScan | None]:
    """The content types listed in [Content_Types].xml (info; None when the
    package has none), its doubts recorded in reach, and the scan of that
    part (searched like any other)."""
    types = _Types()
    if info is None:
        reach.doubt("the package has no list of its parts' content types")
        return types, None
    scan = _XmlScan(info.filename, _STORED_CALL_RE, formulas=False)
    if not _parse_member(zf, info, scan, _STORED_CALL_RE, reach):
        return types, None
    for local, attrs in scan.types:
        a = {_local(k): v for k, v in attrs.items()}
        ct = (a.get("ContentType") or "").strip()
        key = a.get("Extension") if local == "Default" else a.get("PartName")
        if not ct or not key:
            reach.doubt(f"the part {info.filename} has an entry without a content type, "
                        "an extension or a part name")
            continue
        table, key = ((types.default, key.casefold()) if local == "Default"
                      else (types.over, _norm(key)))
        if table.get(key, ct).lower() != ct.lower():
            reach.doubt(f"the part {info.filename} gives {key} two content types")
        table.setdefault(key, ct)
    return types, scan


def _structure(zf: zipfile.ZipFile, infos: list, by_norm: dict, reach: _Reach):
    """(workbook part, {sheet part: sheet name}) from the parsed content
    types and relationships; every doubt is recorded in reach. The workbook
    part is None when the root relationships do not name exactly one part
    that is in the package."""
    def parse(info):
        scan = _XmlScan(info.filename, _STORED_CALL_RE, formulas=False)
        return scan if _parse_member(zf, info, scan, _STORED_CALL_RE, reach) else None

    rels: dict[str, list[dict]] = {}
    for info in infos:
        low = info.filename.casefold()
        if not low.endswith(".rels"):
            continue
        scan = parse(info)
        if scan is None:
            continue
        d, base = posixpath.split(info.filename)
        if posixpath.basename(d).casefold() != "_rels":
            reach.doubt(f"the part {info.filename} is a relationships part outside a _rels folder")
            continue
        source = posixpath.join(posixpath.dirname(d), base[:-len(".rels")])
        ids: set[str] = set()
        for r in scan.rels:
            rid, target, mode = r.get("Id"), r.get("Target"), r.get("TargetMode", "Internal")
            if not rid or rid in ids or not r.get("Type") or target is None:
                reach.doubt(f"the part {info.filename} has a relationship without a unique "
                            "Id, a type or a target")
                continue
            ids.add(rid)
            if mode == "External":
                continue
            if mode != "Internal":
                reach.doubt(f"the part {info.filename} has a relationship of the unknown "
                            f"target mode {mode}")
                continue
            part = _resolve(source, target)
            hit = by_norm.get(_norm(part))
            if hit is None:
                reach.doubt(f"the relationship {rid} in {info.filename} points at {target}, "
                            "which is not in the package")
                continue
            r["_part"] = hit
        rels[_norm(source)] = scan.rels
    root = [r for r in rels.get("", []) if r.get("Type", "").endswith("/officeDocument")]
    if len(root) != 1 or "_part" not in root[0]:
        reach.doubt("the package's root relationships do not name exactly one workbook part "
                    "that it holds")
        return None, {}
    wb_part = root[0]["_part"]
    return wb_part, rels


def _sheets(zf, wb_info, rels: dict, reach: _Reach) -> dict[str, str]:
    scan = _XmlScan(wb_info.filename, _STORED_CALL_RE, kind="workbook", formulas=False)
    if not _parse_member(zf, wb_info, scan, _STORED_CALL_RE, reach):
        return {}
    by_id = {r.get("Id"): r for r in rels.get(_norm(wb_info.filename), [])}
    out: dict[str, str] = {}
    for s in scan.sheets:
        r = by_id.get(s.get("rid"))
        if r is None or "_part" not in r:
            reach.doubt(f"the sheet {s.get('name', '?')} names a relationship that cannot be "
                        "resolved")
            continue
        out.setdefault(r["_part"], s.get("name", r["_part"]))
    return out


def _package_reach(zf: zipfile.ZipFile, path: str) -> tuple[str, dict | None]:
    """file_reach for a ZIP package (already through xmlsafe.check_package)."""
    reach = _Reach()
    infos = []
    for info in zf.infolist():
        if not info.is_dir():
            infos.append(info)
        elif info.file_size:
            reach.doubt(f"the member {info.filename} is named like a folder and holds data")
    by_norm: dict[str, str] = {}
    for info in infos:
        key = _norm(info.filename)
        if key in by_norm:
            reach.doubt(f"the package holds {by_norm[key]} and {info.filename}, which name the "
                        "same part")
        by_norm.setdefault(key, info.filename)
    is_opc = "[content_types].xml" in by_norm or "_rels/.rels" in by_norm
    total = sum(i.file_size for i in infos)
    if total > _MAX_READ:
        reach.doubt(f"the file holds {total // (1024 * 1024):,} MB once unpacked, more than "
                    f"this check reads in full ({_MAX_READ // (1024 * 1024):,} MB)")
        return (READ_PACKAGE if is_opc else READ_TEXT), reach.result()
    if not is_opc:
        return READ_TEXT, _zip_text_reach(zf, infos, reach)
    info_of = {i.filename: i for i in infos}
    ct_info = info_of.get(_CT_PART)
    types, ct_scan = _content_types(zf, ct_info, reach)
    wb_part, rels = _structure(zf, infos, by_norm, reach)
    sheets: dict[str, str] = {}
    if wb_part is not None:
        wb_type = (types.of(wb_part) or "").split(";", 1)[0].strip().lower()
        if wb_type == _XLSB_MAIN:
            # a binary workbook (an .xlsb): Excel looks with calculation off
            return READ_IN_EXCEL, None
        sheets = _sheets(zf, info_of[wb_part], rels, reach)
    # Excel loads a part through the relationship that reaches it, whatever its content
    # type says (measured: a sheet part typed image/png, or of a made-up type, opens and
    # fetches). A part of a binary or unknown type that a relationship for XML content
    # reaches is therefore read as XML as well, and the file counts as unread.
    loaded: dict[str, list[str]] = {}
    for rs in rels.values():
        for r in rs:
            if "_part" in r and r.get("Type", "").rsplit("/", 1)[-1].casefold() not in _BINARY_RELS:
                loaded.setdefault(r["_part"], []).append(r.get("Type", ""))
    order = [info_of[p] for p in sheets] + ([info_of[wb_part]] if wb_part else [])
    order += [i for i in infos if i not in order]
    for info in order:
        if info is ct_info:
            if ct_scan is not None:
                reach.take(ct_scan)
            continue
        ct = types.of(info.filename)
        how = _type_kind(ct)
        if how == "unknown":
            reach.doubt(f"the part {info.filename} has no content type in {_CT_PART}"
                        if not ct else
                        f"the part {info.filename} has the content type {ct}, which this "
                        "check does not know")
        as_xml = how != "xml" and info.filename in loaded
        if as_xml:
            role = ("the workbook" if info.filename == wb_part else "a sheet"
                    if info.filename in sheets else
                    "a part of the type " + loaded[info.filename][0].rsplit("/", 1)[-1])
            reach.doubt(f"the part {info.filename} is loaded as {role}, and its content type "
                        f"({ct or 'none'}) is not an XML type")
        _read_member(zf, info, reach, xml=how == "xml" or as_xml,
                     kind="sheet" if info.filename in sheets else
                     "workbook" if info.filename == wb_part else "other",
                     title=sheets.get(info.filename, ""), pattern=_STORED_CALL_RE)
    return READ_PACKAGE, reach.result()


def _read_member(zf, info, reach: _Reach, *, xml: bool, kind: str, title: str,
                 pattern: re.Pattern, loose_where: str | None = None) -> None:
    """Read one member: as XML (xml=True) through the parser and the raw-text
    search; otherwise through the raw-text search alone, over every byte
    (each byte one character, so a name in plain bytes is seen whatever the
    encoding around it) and with character references resolved. Every byte
    is read either way, so a damaged member still counts."""
    scan = _XmlScan(info.filename, pattern, kind=kind, title=title)
    if xml:
        _parse_member(zf, info, scan, pattern, reach)
    else:
        member = _Member(zf, info)
        try:
            for name in _text_calls((raw.decode("latin-1") for raw in member.chunks()),
                                    pattern=pattern):
                if name not in scan.loose:
                    scan.loose.append(name)
            for d in member.doubts():
                reach.doubt(f"the part {info.filename}: {d}")
        except _READ_ERRORS as exc:
            reach.doubt(f"the part {info.filename}: {_reason(exc)}")
    reach.take(scan, loose_where)


def _parse_member(zf, info, scan: _XmlScan, pattern: re.Pattern, reach: _Reach) -> bool:
    """_read_xml for a ZIP member, its doubts recorded in reach. False when
    the member could not be read at all."""
    member = _Member(zf, info)
    try:
        doubts = _read_xml(member.chunks(), scan, pattern) + member.doubts()
    except _READ_ERRORS as exc:
        reach.doubt(f"the part {info.filename}: {_reason(exc)}")
        return False
    for d in doubts:
        reach.doubt(f"the part {info.filename}: {d}")
    return True


def _zip_text_reach(zf, infos, reach: _Reach) -> dict | None:
    """An OpenDocument spreadsheet or another ZIP that is not an Excel
    package: every member read as text, any spelling counting. Such a file
    has no content types; its reader takes its parts by fixed names
    (content.xml), so a member is parsed as XML by its name, or by its first
    bytes, and every other member is searched byte for byte."""
    for info in infos:
        try:
            head = _Member(zf, info).head()
        except _READ_ERRORS as exc:
            reach.doubt(f"the part {info.filename}: {_reason(exc)}")
            continue
        _read_member(zf, info, reach,
                     xml=info.filename.casefold().endswith(_XML_SUFFIXES) or _looks_like_xml(head),
                     kind="other", title="", pattern=_TEXT_CALL_RE,
                     loose_where="the formulas in the file")
    for f in reach.formulas:  # one place for the whole file, as before
        if f["where"] != "the formulas in the file":
            reach.add("the formulas in the file", f["functions"])
    reach.formulas[:] = [f for f in reach.formulas if f["where"] == "the formulas in the file"]
    return reach.result()


# ------------------------------------------------------------ text files

def _file_chunks(path: str, size: int = _TEXT_CHUNK):
    with open(path, "rb") as fh:
        while True:
            raw = fh.read(size)
            if not raw:
                return
            yield raw


def _lead(path: str, enc: str | None) -> str:
    """The first characters of a text file after its leading whitespace,
    however much of it there is (decoding errors replaced)."""
    dec = codecs.getincrementaldecoder(enc or "utf-8")(errors="replace")
    text = ""
    for raw in _file_chunks(path, _CHUNK):
        text = (text + dec.decode(raw)).lstrip(" \t\r\n\ufeff")
        if len(text) >= 4096:
            break
    return text[:4096]


def _plain_text_reach(path: str) -> dict | None:
    """A text file as Excel opens it: UTF-8 (with or without its mark) or
    UTF-16 with its mark; text that starts like XML through the parser as
    well. Text with no mark that is not UTF-8 is read byte for byte (Excel
    reads it in the system's code page, where the letters of a function name
    are the same bytes). Anything else counts as unread."""
    reach = _Reach()
    where = "the text of the file"
    try:
        size = os.path.getsize(path)
        if size > _MAX_READ:
            reach.doubt(f"the file holds {size // (1024 * 1024):,} MB, more than this check "
                        f"reads in full ({_MAX_READ // (1024 * 1024):,} MB)")
            return reach.result()
        with open(path, "rb") as fh:
            head = fh.read(4096)
        if head.startswith((b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff")):
            reach.doubt("the file is UTF-32 text, which Excel does not open as text")
            return reach.result()
        if head.startswith((b"\xff\xfe", b"\xfe\xff")):
            enc = "utf-16"
        elif head.startswith(b"\xef\xbb\xbf"):
            enc = "utf-8-sig"
        else:
            enc = None
        start = head.decode(enc or "utf-8", "replace") if enc != "utf-16" else \
            head[2:len(head) & ~1].decode("utf-16-le" if head[:2] == b"\xff\xfe" else "utf-16-be",
                                          "replace")
        if not start.lstrip("\ufeff").strip() and size > len(head):
            start = _lead(path, enc)  # whitespace, as much as it likes, before the first text
        if _XML_START_RE.match(start.lstrip("\ufeff")) or (
                enc is None and _xmlsafe.byte_encoding(head) != "utf-8"):
            scan = _XmlScan(Path(path).name, _TEXT_CALL_RE)
            for d in _read_xml(_file_chunks(path, _CHUNK), scan, _TEXT_CALL_RE):
                reach.doubt(f"the file: {d}")
            reach.take(scan, where)
            reach.formulas[:] = [{"where": where, "functions": [fn for f in reach.formulas
                                                                for fn in f["functions"]]}] \
                if reach.formulas else []
            return reach.result()
        if enc is None:
            dec = codecs.getincrementaldecoder("utf-8")(errors="surrogateescape")
        else:
            dec = codecs.getincrementaldecoder(enc)(errors="strict")

        def texts():
            for raw in _file_chunks(path):
                if enc is None and b"\x00" in raw:
                    raise _Unread("the file holds zero bytes, so it is not text Excel reads "
                                  "as written")
                yield dec.decode(raw)
            yield dec.decode(b"", True)

        reach.add(where, _text_calls(texts(), xml=False))
    except UnicodeDecodeError:
        reach.doubt("the file's text does not decode in the encoding its mark names")
    except _Unread as exc:
        reach.doubt(str(exc))
    except OSError as exc:
        reach.doubt(_reason(exc))
    return reach.result()


def file_reach(path: str) -> tuple[str, dict | None]:
    """(how the file was read, what can reach the network) for any file
    Excel may be asked to open. The answer is None only when the whole file
    was read with certainty and nothing in it can reach the network; a file
    the check cannot read in full carries the reasons in "unread".

    A text file and an OpenDocument package enter their formulas as Excel
    loads them, so Excel would calculate them at once even with calculation
    turned off (measured with a CSV); they are read here as text, and any
    call to a network-capable function counts. An .xlsb, a legacy .xls or an
    encrypted workbook stores its formulas in a form this module does not
    read: READ_IN_EXCEL, and the COM open (com.session) has Excel look with
    calculation turned off."""
    try:
        with open(path, "rb") as fh:
            head = fh.read(8)
    except OSError as exc:
        if not os.path.isfile(path):
            return READ_PACKAGE, None  # the open path reports a missing file
        reach = _Reach()  # a file that is there but cannot be read here
        reach.doubt(_reason(exc))
        return READ_PACKAGE, reach.result()
    if head == _OLE_MAGIC:
        return READ_IN_EXCEL, None
    if head[:2] == b"PK":
        try:
            zf = zipfile.ZipFile(path)
        except (zipfile.BadZipFile, OSError, ValueError, EOFError) as exc:
            if head[:4] != b"PK\x03\x04" and isinstance(exc, zipfile.BadZipFile):
                zf = None  # text that happens to start with "PK"
            else:
                # a package whose ZIP structure cannot be read at all: only Excel can
                # say what it makes of it, so Excel looks with calculation turned off
                # (as for an .xlsb), and a file Excel cannot open is reported as such
                return READ_IN_EXCEL, None
        if zf is not None:
            with zf:
                # a decompression bomb, a part named twice or a part that
                # declares a DTD is refused before any part is read
                _xmlsafe.check_package(zf, doc=path)
                return _package_reach(zf, path)
    return READ_TEXT, _plain_text_reach(path)


def network_reach(path: str) -> dict | None:
    """What in the workbook at path can reach the network when Excel
    calculates it: {"formulas": [{"where", "functions"}], "connections":
    [names]}, plus "unread": [reasons] when the file cannot be read with
    certainty; None when nothing can. A file only Excel can read (an .xlsb,
    a legacy or encrypted workbook) answers None here: the shared COM open
    has Excel look at it with calculation turned off."""
    how, reach = file_reach(path)
    return None if how == READ_IN_EXCEL else reach


def network_refusal(path: str, reach: dict, op: str = "recalculate", *,
                    started: bool = False, looked: bool = False):
    """The plain refusal for a workbook that would reach the network, or
    that the check could not read with certainty, naming the places and the
    functions, and the opt-in. started: Excel was already running when the
    file was judged; looked: Excel opened the file with calculation turned
    off to find these, and closed it."""
    from .errors import FormulaRejected
    name = Path(path).name
    parts = []
    formulas = reach.get("formulas") or []
    if formulas:
        shown = "; ".join(f"{f['where']} ({', '.join(f['functions'])})"
                          for f in formulas[:_SHOW])
        more = f"; and {len(formulas) - _SHOW} more" if len(formulas) > _SHOW else ""
        parts.append(f"formulas that fetch from the network when Excel calculates them: "
                     f"{shown}{more}")
    conns = reach.get("connections") or []
    if conns:
        shown = ", ".join(conns[:_SHOW])
        more = f" and {len(conns) - _SHOW} more" if len(conns) > _SHOW else ""
        parts.append(f"external data connections that Excel can refresh from their "
                     f"sources: {shown}{more}")
    unread = reach.get("unread") or []
    if unread:
        shown = "; ".join(unread[:_SHOW])
        more = f"; and {len(unread) - _SHOW} more" if len(unread) > _SHOW else ""
        parts.append("content this server could not read with certainty, so it cannot rule "
                     f"out a network request when Excel calculates: {shown}{more}")
    if looked:
        stop = ("Excel opened it with calculation turned off to look, and closed "
                "it without calculating")
    elif started:
        stop = "so Excel did not open it"
    else:
        stop = "so Excel was not started"
    exc = FormulaRejected(
        f"{name} has " + "; it also has ".join(parts) + ". This server makes no "
        f"network request you did not ask for, {stop}. The values cached from the "
        "last calculation stay readable either way. Ask the user whether this "
        "workbook may contact the internet. Only if they agree, run the same "
        "call again with allow_network=true, and tell them it was allowed.")
    exc.detail = {"network_reach": reach, "override": "allow_network=true"}
    return exc


def refuse_network_reach(path: str, allow_network: bool, op: str = "recalculate") -> None:
    """Raise the refusal unless allow_network, when the workbook at path can
    reach the network as Excel calculates it."""
    if allow_network:
        return
    reach = network_reach(path)
    if reach:
        raise network_refusal(path, reach, op)


__all__ = ["NETWORK_FUNCTIONS", "READ_IN_EXCEL", "READ_PACKAGE", "READ_TEXT", "calls_in",
           "file_reach", "network_reach", "network_refusal", "refuse_network_reach"]
