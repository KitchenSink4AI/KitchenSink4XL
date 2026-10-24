"""core/sharedformulas.py: shared formulas whose first cell cannot be found.

A shared formula is stored once, in its first cell (the one with the formula
text and a `ref`), and every other cell that uses it carries only the shared
index (<f t="shared" si="0"/>). When the first cell is missing, or comes AFTER
a cell that uses it, the formula of those cells cannot be rebuilt. Excel
repairs such a file by dropping the formula and keeping the stored value;
openpyxl reads the cell as the formula "=".

readcompat.py makes the read side read what Excel reads (the stored value).
This module finds the cells, for the hazard scan and for the flag that the
read tools attach so a caller knows no formula was recovered.
"""

from __future__ import annotations

import codecs
import os
import re
import threading

from . import surgical as _s

NOTE = ("these cells are stored as part of a shared formula whose first cell is "
        "missing or comes after them, so no formula can be recovered. The value "
        "shown is the value stored in the file; Excel drops the formula in the "
        "same way when it repairs the file")

LABEL = "shared formulas with no usable first cell"

_F_OPEN = re.compile(rb"<(?:[\w.\-]+:)?f\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>")
_SI = re.compile(rb"""\bsi\s*=\s*(?:"([^"]*)"|'([^']*)')""")
_SHARED = re.compile(rb"""\bt\s*=\s*(?:"shared"|'shared')""")


def _plain(data: bytes) -> bool:
    """True when the bytes of the part can be read as written: an
    ASCII-compatible encoding that starts with "<", with no character
    reference, comment, CDATA section or processing instruction (past the XML
    declaration) that could spell an attribute value differently or hide
    markup-looking text. Anything else is not read here; the parsed pass
    decides."""
    from . import xmlsafe as _xmlsafe
    body = data[3:] if data.startswith(codecs.BOM_UTF8) else data
    if _xmlsafe.byte_encoding(data) != "utf-8" or body[:1] != b"<":
        return False
    if b"&#" in data or b"<!" in data:
        return False
    return data.count(b"<?") == (1 if body.startswith(b"<?xml") else 0)


def _may_have(data: bytes) -> bool:
    """False only when the part cannot hold a shared formula."""
    return not _plain(data) or b"shared" in data


def _quick_broken(data: bytes) -> bool:
    """A cheap pass over the bytes: could there be a shared member (no text)
    whose index no earlier element defined? True whenever the bytes cannot be
    trusted to say (see _plain); the parsed pass then decides."""
    if not _plain(data):
        return True
    seen: set[bytes] = set()
    for m in _F_OPEN.finditer(data):
        attrs = m.group(1)
        if not _SHARED.search(attrs):
            continue
        si = _SI.search(attrs)
        key = (si.group(1) if si.group(1) is not None else si.group(2)
               ) if si is not None else None
        selfclose = attrs.rstrip().endswith(b"/")
        has_text = (not selfclose) and data[m.end():m.end() + 1] not in (b"<", b"")
        if has_text:
            if key is not None:
                seen.add(key)
        elif key is None or key not in seen:
            return True
    return False


def broken_members(data: bytes) -> list[str]:
    """The A1 addresses, in sheet order, of the shared-formula members of one
    worksheet part that have no usable first cell. openpyxl's own rule: a
    member is any shared element with no text; the first element with text
    for an index defines it."""
    if not _may_have(data) or not _quick_broken(data):
        return []
    from . import hazard as _hazard
    from . import xmlsafe as _xmlsafe
    tags = _hazard._main_tags("row")
    c_tags = set(_hazard._main_tags("c"))
    f_tags = set(_hazard._main_tags("f"))
    seen: set[str | None] = set()
    out: list[str] = []
    row_no = 0
    for row in _xmlsafe.stream_elements(data, tags, part="a worksheet"):
        parent = row.getparent()
        if parent is None or _hazard._local(parent.tag) != "sheetData":
            continue
        given = (row.get("r") or "").strip()
        row_no = int(given) if given.isdigit() else row_no + 1
        col = 0
        for cell in row:
            if cell.tag not in c_tags:
                continue
            rc = _hazard._cell_rc(cell.get("r") or "")
            if rc is not None:
                row_no, col = rc
            else:
                col += 1
            for child in cell:
                if child.tag not in f_tags or child.get("t") != "shared":
                    continue
                si = child.get("si")
                if child.text is not None:
                    seen.add(si)
                elif si not in seen:
                    out.append(f"{_col_letters(col)}{row_no}")
                break
        _hazard._discard(row)
    return out


def _col_letters(col: int) -> str:
    return _s._col_letters(col)


_CACHE: dict[str, tuple[tuple, dict]] = {}
_LOCK = threading.Lock()


def broken_cells(path: str) -> dict[str, list[str]]:
    """{sheet title: [addresses]} for a workbook file; {} when there are none
    or the file cannot be read (a read must never fail over this flag)."""
    import zipfile
    try:
        st = os.stat(path)
        key = (st.st_mtime_ns, st.st_size)
    except OSError:
        return {}
    ap = os.path.abspath(path)
    with _LOCK:
        hit = _CACHE.get(ap)
        if hit is not None and hit[0] == key:
            return hit[1]
    found: dict[str, list[str]] = {}
    try:
        with zipfile.ZipFile(path) as zf:
            pkg = _s.Package(zf)
            for title, part in pkg.sheet_parts().items():
                cells = broken_members(pkg.read(part))
                if cells:
                    found[title] = cells
    except Exception:  # noqa: BLE001
        return {}
    with _LOCK:
        if len(_CACHE) > 16:
            _CACHE.clear()
        _CACHE[ap] = (key, found)
    return found


__all__ = ["NOTE", "LABEL", "broken_members", "broken_cells"]
