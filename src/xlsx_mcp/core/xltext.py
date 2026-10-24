"""core/xltext.py: cell text as Excel stores it and as Excel reads it.

OOXML stores characters XML cannot carry, and a carriage return Excel wants
kept, as an escape: _xHHHH_ names one UTF-16 code unit (_x000D_ is a
carriage return, _x0009_ a tab, _xD83D__xDE00_ an emoji), and a literal
underscore that would otherwise start one is written _x005F_. Excel's own
reading, measured on shared strings and inline strings alike: the escape is
an underscore, a LOWER-CASE x, exactly four hexadecimal digits in either
case and an underscore; it is read left to right in one pass (_x005F_x0041_
is the text "_x0041_"); anything else (_X0041_, _x0041, _x00410_, Ax005F_B)
is text as it stands.

openpyxl reads neither form right. Its shared-string reader deletes the
substring "x005F_" wherever it sits, so the eight characters Ax005F_B read
as AB (and a full rewrite then stored AB), and it decodes no other escape,
in shared or inline strings, so a cell Excel shows as two lines read as
"line one_x000D_ line two". Its writer escapes nothing, so text holding
"_x0041_" was stored as an escape Excel shows as "A".

install() fixes both ends once per process, so every load (reads, the
model a write edits, the verify read-back) holds the characters Excel
shows, and every openpyxl save stores them as Excel would:
  - the shared-string table is read without the deletion;
  - a string cell's text (shared, inline, or a formula's text result) is
    decoded by Excel's rule;
  - a string cell's text is escaped on write: an underscore that starts an
    escape becomes _x005F_, and characters XML cannot carry (C0 controls
    other than tab and line feed, and carriage return, as Excel writes it)
    become _xHHHH_.
A lone surrogate escape (one half of a pair) is left as its text: decoded,
it could not be carried in a reply at all.
"""

from __future__ import annotations

import re

#: One escape as Excel reads it.
ESCAPE_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")
#: An escape-shaped run in text being written, other than a surrogate half
#: (left as it stands, so a stored lone-surrogate escape round-trips).
_LITERAL_ESCAPE_RE = re.compile(
    r"_(x(?![dD][89a-fA-F][0-9A-Fa-f]{2}_)[0-9A-Fa-f]{4}_)")
#: Characters a cell's text cannot carry as they are.
_UNSAFE_RE = re.compile("[\x00-\x08\x0b\x0c\x0d\x0e-\x1f\ufffe\uffff]")


def decode(text: str) -> str:
    """The characters Excel shows for stored cell text."""
    if "_x" not in text:
        return text
    matches = list(ESCAPE_RE.finditer(text))
    out: list[str] = []
    pos = 0
    i = 0
    while i < len(matches):
        m = matches[i]
        out.append(text[pos:m.start()])
        unit = int(m.group(1), 16)
        if 0xD800 <= unit <= 0xDBFF and i + 1 < len(matches) \
                and matches[i + 1].start() == m.end():
            low = int(matches[i + 1].group(1), 16)
            if 0xDC00 <= low <= 0xDFFF:
                out.append(chr(0x10000 + ((unit - 0xD800) << 10)
                               + (low - 0xDC00)))
                pos = matches[i + 1].end()
                i += 2
                continue
        # an unpaired surrogate half stays as its text
        out.append(m.group(0) if 0xD800 <= unit <= 0xDFFF else chr(unit))
        pos = m.end()
        i += 1
    out.append(text[pos:])
    return "".join(out)


def encode(text: str) -> str:
    """Cell text as Excel stores it (the inverse of decode)."""
    if "_x" in text:
        text = _LITERAL_ESCAPE_RE.sub(r"_x005F_\1", text)
    if _UNSAFE_RE.search(text):
        text = _UNSAFE_RE.sub(lambda m: "_x%04X_" % ord(m.group(0)), text)
    return text


def needs_encoding(text: str) -> bool:
    return "_x" in text or _UNSAFE_RE.search(text) is not None


_INSTALLED = False


def install() -> None:
    """Patch openpyxl's string reading and writing (idempotent)."""
    global _INSTALLED
    if _INSTALLED:
        return
    _INSTALLED = True
    from openpyxl.cell.text import Text
    from openpyxl.worksheet import _reader as _wsreader
    from openpyxl.worksheet import _writer as _wswriter
    from openpyxl.xml.constants import SHEET_MAIN_NS

    from . import xmlsafe as _xmlsafe

    def read_string_table(xml_source):
        """openpyxl's reader without its "x005F_" deletion; the escapes
        are decoded per cell (parse_cell below), once. The parse is
        xmlsafe's, as for every other part this server parses."""
        strings = []
        tag = "{%s}si" % SHEET_MAIN_NS
        part = getattr(xml_source, "name", None) or "the shared strings"
        for _, node in _xmlsafe.iterparse(xml_source, part=str(part)):
            if node.tag == tag:
                strings.append(Text.from_tree(node).content)
                node.clear()
        return strings

    _xmlsafe.install_string_table(read_string_table)

    parse_cell = _wsreader.WorkSheetParser.parse_cell

    def decoding_parse_cell(self, element):
        cell = parse_cell(self, element)
        value = cell.get("value")
        if isinstance(value, str) and "_x" in value and \
                element.get("t") in ("s", "inlineStr", "str"):
            cell["value"] = decode(value)
        return cell

    _wsreader.WorkSheetParser.parse_cell = decoding_parse_cell

    write_cell = _wswriter.write_cell

    def encoding_write_cell(xf, worksheet, cell, styled=None):
        value = cell._value
        if cell.data_type == "s" and isinstance(value, str) and \
                needs_encoding(value):
            cell._value = encode(value)
            try:
                return write_cell(xf, worksheet, cell, styled)
            finally:
                cell._value = value
        return write_cell(xf, worksheet, cell, styled)

    _wswriter.write_cell = encoding_write_cell


__all__ = ["decode", "encode", "needs_encoding", "install", "ESCAPE_RE"]
