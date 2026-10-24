"""Namespace-aware reads of the small Open Packaging parts a check decides on:
a relationships part, and the first element of a part.

The checks that sniff a package (is it Strict Open XML, is a query table's
part really there, which note part belongs to a sheet) used to read these
parts as text with regular expressions, so a valid file spelled differently
(single-quoted attributes, a prefixed namespace, UTF-16, a character
reference) slipped past them. Everything here goes through the hardened
parser in xmlsafe (no DTD, no entities, no network), reads names by their
namespace, and reads attribute values after the XML rules have decoded them.
"""

from __future__ import annotations

import posixpath
from typing import NamedTuple

from . import partnames as _partnames
from . import xmlsafe as _xmlsafe

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
STRICT_MAIN_NS = "http://purl.oclc.org/ooxml/spreadsheetml/main"
#: Both spreadsheet namespaces: the one Excel writes and the Strict one.
MAIN_NAMESPACES = (MAIN_NS, STRICT_MAIN_NS)
REL_TYPE_ROOT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
STRICT_REL_TYPE_ROOT = "http://purl.oclc.org/ooxml/officeDocument/relationships/"

#: How much of a part is read to find its first element.
ROOT_LOOK_BYTES = 1 << 20


class Rel(NamedTuple):
    id: str
    type: str
    target: str
    external: bool


def relationships(data: bytes, *, doc: str = "the workbook",
                  part: str = "a relationships part") -> list[Rel]:
    """The Relationship elements of one relationships part, by namespace
    (a prefixed element counts, a lookalike in another namespace does not).
    Raises what xmlsafe.fromstring raises for a part that is not well-formed
    or declares a DTD."""
    root = _xmlsafe.fromstring(data, doc=doc, part=part)
    out: list[Rel] = []
    for el in root.iter(f"{{{PKG_REL_NS}}}Relationship"):
        mode = (el.get("TargetMode") or "").strip().lower()
        out.append(Rel(el.get("Id") or "", (el.get("Type") or "").strip(),
                       (el.get("Target") or "").strip(), mode == "external"))
    return out


def resolve_part(source_part: str, target: str) -> str:
    """A relationship Target as a package part name (lower-case compare is
    the caller's)."""
    t = _partnames.normalize(target.replace("\\", "/"))
    if t.startswith("/"):
        return posixpath.normpath(t.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), t))


def rels_part_for(part: str) -> str:
    head, base = posixpath.split(part)
    return posixpath.join(head, "_rels", base + ".rels")


class _Found(Exception):
    def __init__(self, tag: str):
        self.tag = tag


def root_tag(data: bytes, *, doc: str = "the workbook",
             part: str = "an XML part") -> tuple[str, str] | None:
    """(namespace, local name) of a part's root element, read from at most
    its first ROOT_LOOK_BYTES bytes in whatever encoding the part declares;
    None when no root start tag is found there."""
    def start(tag, _attrib):
        raise _Found(tag)

    try:
        _xmlsafe.walk(data[:ROOT_LOOK_BYTES], start=start, doc=doc, part=part)
    except _Found as hit:
        tag = hit.tag
        if tag.startswith("{"):
            ns, _sep, local = tag[1:].partition("}")
            return ns, local
        return "", tag
    except _xmlsafe.XML_ERRORS:
        return None
    return None
