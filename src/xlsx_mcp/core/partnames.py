"""core/partnames.py: part names the way Open Packaging Conventions compare them.

A relationship Target is a URI reference. Two spellings of the same part name
are equal when they differ only by a percent code for an unreserved character
(sheet%31.xml and sheet1.xml name the same part), and Excel opens a workbook
that spells its sheet targets that way. Readers that took the Target as the
literal zip entry name found no sheet and reported an empty workbook.

Only the unreserved characters (letters, digits, - . _ ~) are decoded. A code
for anything else (a space as %20, a non-ASCII letter as %E8%A1%A8) is left as
written, because Excel does not read it as the plain character either: a
workbook whose target needs that decoding is not one Excel opens, and the
callers refuse it by name when the part cannot be found.

No imports from the rest of the package, so every module may use it.
"""

from __future__ import annotations

import posixpath
import re

_PERCENT = re.compile(r"%([0-9A-Fa-f]{2})")
_UNRESERVED = frozenset(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~")


def normalize(name: str) -> str:
    """`name` with each percent code for an unreserved character decoded and
    the hex digits of every other code upper-cased (the OPC equivalence form).
    A name with no percent sign comes back unchanged."""
    if "%" not in name:
        return name

    def _one(m: re.Match) -> str:
        ch = chr(int(m.group(1), 16))
        return ch if ch in _UNRESERVED else "%" + m.group(1).upper()

    return _PERCENT.sub(_one, name)


def _resolve(source_part: str, target: str) -> str:
    t = target.replace("\\", "/")
    if t.startswith("/"):
        return posixpath.normpath(t.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), t))


def source_of_rels(rels_part: str) -> str | None:
    """The part a relationships part belongs to: xl/_rels/workbook.xml.rels
    is the relationships of xl/workbook.xml; _rels/.rels belongs to the
    package ("")."""
    head, base = posixpath.split(rels_part)
    if posixpath.basename(head) != "_rels" or not base.endswith(".rels"):
        return None
    return posixpath.join(posixpath.dirname(head), base[:-len(".rels")])


def spelled_target(source_part: str, target: str, names) -> str | None:
    """The text a relationship Target should be read as when, as written, it
    names no part of the package but its OPC-equivalent spelling does. None
    when the target already names a part, or when no spelling does. `names`
    is any container of zip entry names."""
    if "%" not in target:
        return None
    if _resolve(source_part, target) in names:
        return None
    plain = normalize(target)
    if plain != target and _resolve(source_part, plain) in names:
        return plain
    return None


_ATTR = re.compile(rb"""(\bTarget\s*=\s*)(?:"([^"]*)"|'([^']*)')""")


def respell_rels(rels_part: str, data: bytes, names) -> bytes:
    """A relationships part with each internal Target whose literal spelling
    names no part, but whose OPC-equivalent spelling does, rewritten to that
    spelling. The package on disk is not touched; this is the view a reader is
    given. Bytes come back unchanged when nothing needs it."""
    if b"%" not in data or data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return data
    source = source_of_rels(rels_part)
    if source is None:
        return data

    def _one(m: re.Match) -> bytes:
        quoted = m.group(2) if m.group(2) is not None else m.group(3)
        try:
            text = quoted.decode("utf-8")
        except UnicodeDecodeError:
            return m.group(0)
        better = spelled_target(source, text, names)
        if better is None:
            return m.group(0)
        q = b'"' if m.group(2) is not None else b"'"
        return m.group(1) + q + better.encode("utf-8") + q

    return _ATTR.sub(_one, data)
