"""The byte-preserving writer reuses a shared string only when it means the
same text.

The shared-string table is appended to, never rewritten, and a write reuses
an existing entry when the text matches. The match used to be made against
the entry decoded with html.unescape() and a blanket .replace("x005F_", ""),
which read two different stored forms as the value being written:

  - "Ax005F_B" read as "AB" (the _xHHHH_ escape, deleted wherever it sat, not
    only as the OOXML _x005F_ escape for a literal underscore);
  - "&#x80;" read as the euro sign, because html.unescape applies HTML5's
    Windows-1252 remap of the C1 numeric references, which XML does not.

So writing "AB" reused the cell pointing at "Ax005F_B", and writing the euro
sign reused the cell storing "&#x80;" (U+0080); Excel then showed the stored
text, not the value asked for. The entries are now compared as the text they
decode to under XML and OOXML rules, so a reuse means the same characters.

The workbooks are built by code: a crafted sharedStrings.xml with the
colliding entries, and cells that point at them.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.core import surgical as _surgical
from xlsx_mcp.ops import cells

FIX = Path(__file__).resolve().parents[1] / "fixtures" / "preserve"

# Three stored entries with distinct, carefully chosen raw <t> bytes:
#   si 0  "Keep"           an ordinary control
#   si 1  "Ax005F_B"       decodes (XML + OOXML) to the 8 literal characters
#   si 2  "&#x80;"         XML numeric ref -> U+0080 (NOT the euro sign)
#   si 3  "_x005F_"        the OOXML escape for a single literal underscore
_SST = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'count="4" uniqueCount="4">'
    "<si><t>Keep</t></si>"
    "<si><t>Ax005F_B</t></si>"
    "<si><t>&#x80;</t></si>"
    "<si><t>_x005F_</t></si>"
    "</sst>"
)

_WB = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>'
)
_WB_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    "</Relationships>"
)
_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    "</Relationships>"
)
_CT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    "</Types>"
)
# A2..A5 point at si 0..3; B-column cells are free targets for the write.
_SHEET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    "<sheetData>"
    '<row r="2"><c r="A2" t="s"><v>0</v></c></row>'
    '<row r="3"><c r="A3" t="s"><v>1</v></c></row>'
    '<row r="4"><c r="A4" t="s"><v>2</v></c></row>'
    '<row r="5"><c r="A5" t="s"><v>3</v></c></row>'
    "</sheetData></worksheet>"
)


def _book(tmp_path) -> str:
    p = str(tmp_path / "sst.xlsx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CT)
        z.writestr("_rels/.rels", _ROOT_RELS)
        z.writestr("xl/workbook.xml", _WB)
        z.writestr("xl/_rels/workbook.xml.rels", _WB_RELS)
        z.writestr("xl/sharedStrings.xml", _SST)
        z.writestr("xl/worksheets/sheet1.xml", _SHEET)
    return p


def _si_texts(path: str) -> list[str]:
    with zipfile.ZipFile(path) as z:
        sst = z.read("xl/sharedStrings.xml").decode("utf-8")
    return re.findall(r"<si><t(?:[^>]*)>(.*?)</t></si>", sst, re.S)


def _cell_si(path: str, coord: str) -> int:
    with zipfile.ZipFile(path) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    m = re.search(rf'<c r="{coord}"[^>]*t="s"[^>]*><v>(\d+)</v></c>', sheet)
    assert m, f"{coord} is not a shared-string cell"
    return int(m.group(1))


# ---------------------------------------------------------- the decoder

def test_the_index_decodes_each_stored_form_as_excel_reads_it(tmp_path):
    sst = _surgical.SharedStrings("xl/sharedStrings.xml", _SST.encode("utf-8"))
    sst._build()
    idx = sst._index
    assert idx.get("Keep") == 0
    # "Ax005F_B" is eight literal characters, not "AB".
    assert idx.get("Ax005F_B") == 1
    assert "AB" not in idx
    # "&#x80;" is U+0080 under XML, not the euro sign.
    assert idx.get("\u0080") == 2
    assert "€" not in idx
    # "_x005F_" is the OOXML escape for one literal underscore.
    assert idx.get("_") == 3


# ------------------------------------------------- reuse on write

def test_writing_AB_does_not_reuse_the_escaped_entry(tmp_path):
    p = _book(tmp_path)
    before = _si_texts(p)
    assert before[1] == "Ax005F_B"
    r = cells.set_cell(p, {"cell": "B3", "sheet": "Data"}, "AB")
    assert r["ok"] and r["verified"]
    si = _cell_si(p, "B3")
    # it must NOT point at the "Ax005F_B" entry
    assert si != 1
    # the entry it points at decodes to "AB"
    texts = _si_texts(p)
    assert texts[si] == "AB"
    # the colliding entry is untouched
    assert texts[1] == "Ax005F_B"


def test_writing_the_euro_sign_does_not_reuse_the_c1_entry(tmp_path):
    p = _book(tmp_path)
    r = cells.set_cell(p, {"cell": "B4", "sheet": "Data"}, "€")
    assert r["ok"] and r["verified"]
    si = _cell_si(p, "B4")
    assert si != 2  # not the "&#x80;" entry
    texts = _si_texts(p)
    assert texts[si] == "€"
    assert texts[2] == "&#x80;"


def test_writing_a_literal_underscore_reuses_the_right_entry(tmp_path):
    p = _book(tmp_path)
    r = cells.set_cell(p, {"cell": "B5", "sheet": "Data"}, "_")
    assert r["ok"] and r["verified"]
    si = _cell_si(p, "B5")
    # "_x005F_" decodes to "_", so this reuse IS correct
    assert si == 3


def test_the_index_matches_the_eight_literal_characters_not_AB(tmp_path):
    # The index key for the escaped entry is the eight characters it decodes
    # to, so index_of is asked for the right string -- the half of the fix
    # the write path can prove without openpyxl's reader in the way.
    sst = _surgical.SharedStrings("xl/sharedStrings.xml", _SST.encode("utf-8"))
    assert sst.index_of("Ax005F_B") == 1       # the matching entry
    assert sst.index_of("AB") == 4             # a NEW entry, never entry 1


def test_writing_the_eight_literal_characters_is_never_silently_AB(tmp_path):
    # The reader now reads <t>Ax005F_B</t> as the eight characters Excel
    # shows (core.xltext), so the write is stored, reads back as written,
    # and reuses the entry that holds exactly those characters. (Before,
    # the reader deleted "x005F_" and the verify refused the write.)
    p = _book(tmp_path)
    out = cells.set_cell(p, {"cell": "B3", "sheet": "Data"}, "Ax005F_B")
    assert out["verified"] is True
    assert _cell_si(p, "B3") == 1
    assert _si_texts(p)[1] == "Ax005F_B"
    got = cells.get_cells(p, ["B3", "A3"], sheet="Data")["cells"]
    assert [row[1] for row in got] == ["Ax005F_B", "Ax005F_B"]
