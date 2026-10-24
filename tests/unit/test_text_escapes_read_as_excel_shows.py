"""Cell text holding _xHHHH_ escapes reads as Excel shows it.

OOXML writes a character XML cannot carry (and the carriage return Excel
keeps) as _xHHHH_, and a literal underscore that would start one as _x005F_.
Reads returned the stored form: "line one_x000D_ line two" for a cell Excel
shows on two lines, and the shared-string reader deleted the substring
"x005F_" wherever it sat, so the eight characters Ax005F_B read as AB (and a
full rewrite stored AB). Writing such text was refused outright.

Every expected value below is Excel's own reading of the same stored text
(shared and inline strings read alike), measured in Excel. All text here is
made up.
"""

from __future__ import annotations

import re
import zipfile

import pytest

from xlsx_mcp.core import xltext
from xlsx_mcp.ops import cells, structure

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

#: stored text -> what Excel shows (measured)
EXCEL_READS = [
    ("line one_x000D_\nline two", "line one\r\nline two"),
    ("tab_x0009_here", "tab\there"),
    ("_x0041_", "A"),
    ("_x005F_x0041_", "_x0041_"),
    ("Ax005F_B", "Ax005F_B"),
    ("_x000d_low", "\rlow"),
    ("_X0041_up", "_X0041_up"),
    ("_x00410_", "_x00410_"),
    ("_x004G_", "_x004G_"),
    ("_xD83D__xDE00_", "\U0001f600"),
    ("soh_x0001_end", "soh\x01end"),
    ("__x0041__", "_A_"),
    ("_x0041", "_x0041"),
    ("x0041_", "x0041_"),
    ("_x005f_x0041_", "_x0041_"),
    ("a_x005F_b", "a_b"),
    ("_x0020_sp", " sp"),
    ("_x00e9_", "é"),
    ("_x00E9_", "é"),
]


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _book(path, stored: list[str], extra_cells: str = "") -> str:
    """Column A: each text as a shared string; column B: the same text as an
    inline string."""
    rows = "".join(
        '<row r="%d"><c r="A%d" t="s"><v>%d</v></c><c r="B%d" t="inlineStr">'
        '<is><t xml:space="preserve">%s</t></is></c></row>'
        % (i, i, i - 1, i, _esc(s)) for i, s in enumerate(stored, start=1))
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="%s"><sheetData>%s%s</sheetData></worksheet>'
             % (NS, rows, extra_cells))
    sst = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<sst xmlns="%s" count="%d" uniqueCount="%d">%s</sst>'
           % (NS, len(stored), len(stored), "".join(
               '<si><t xml:space="preserve">%s</t></si>' % _esc(s)
               for s in stored)))
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    parts = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types"><Default Extension="rels" ContentType="'
            'application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="'
            'application/vnd.openxmlformats-officedocument.spreadsheetml.'
            'worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.sharedStrings+xml"/></Types>',
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships"><Relationship Id="rId1" Type="%s/'
            'officeDocument" Target="xl/workbook.xml"/></Relationships>' % rel,
        "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="%s" xmlns:r="%s"><sheets><sheet name="Data" '
            'sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/>'
            '</workbook>' % (NS, rel),
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships"><Relationship Id="rId1" Type="%s/worksheet" '
            'Target="worksheets/sheet1.xml"/><Relationship Id="rId2" '
            'Type="%s/sharedStrings" Target="sharedStrings.xml"/>'
            '</Relationships>' % (rel, rel),
        "xl/worksheets/sheet1.xml": sheet,
        "xl/sharedStrings.xml": sst,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return str(path)


def test_shared_and_inline_strings_read_as_excel_shows_them(tmp_path):
    stored = [s for s, _ in EXCEL_READS]
    p = _book(tmp_path / "esc.xlsx", stored)
    refs = [f"{c}{i}" for i in range(1, len(stored) + 1) for c in "AB"]
    got = cells.get_cells(p, refs, sheet="Data")["cells"]
    values = {row[0].split("!")[-1]: row[1] for row in got}
    for i, (raw, shown) in enumerate(EXCEL_READS, start=1):
        assert values[f"A{i}"] == shown, (raw, values[f"A{i}"])
        assert values[f"B{i}"] == shown, (raw, values[f"B{i}"])


def test_read_range_reads_the_same_characters(tmp_path):
    p = _book(tmp_path / "esc.xlsx", ["tab_x0009_here", "Ax005F_B"])
    out = cells.read_range(p, "A1:B2", sheet="Data")
    assert out["values"] == [["tab\there", "tab\there"],
                             ["Ax005F_B", "Ax005F_B"]]


def test_a_formulas_text_result_reads_as_excel_shows_it(tmp_path):
    p = _book(tmp_path / "f.xlsx", ["x"], extra_cells=(
        '<row r="5"><c r="C5" t="str"><f>"a"&amp;CHAR(9)&amp;"b"</f>'
        '<v>a_x0009_b</v></c></row>'))
    got = cells.get_cells(p, ["C5"], sheet="Data")["cells"]
    assert got[0][1] == "a\tb"


def _stored_texts(path) -> list[str]:
    """Every cell's stored text, in sheet order, decoded as Excel reads it."""
    with zipfile.ZipFile(path) as z:
        x = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        names = z.namelist()
        sst = z.read("xl/sharedStrings.xml").decode("utf-8") \
            if "xl/sharedStrings.xml" in names else ""
    import html
    si = [html.unescape(t) for t in re.findall(
        r"<si><t[^>]*>(.*?)</t></si>", sst, re.S)]
    out = []
    for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', x, re.S):
        attrs, body = m.group(2), m.group(3)
        if 't="s"' in attrs:
            out.append(xltext.decode(si[int(re.search(
                r"<v>(\d+)</v>", body).group(1))]))
        elif 'inlineStr' in attrs:
            out.append(xltext.decode(html.unescape(re.search(
                r"<t[^>]*>(.*?)</t>", body, re.S).group(1))))
    return out


def test_a_full_rewrite_stores_the_same_characters(tmp_path):
    """A rewrite of the whole workbook (here a column insert) used to store
    AB for Ax005F_B; every cell must still show what it showed."""
    stored = [s for s, _ in EXCEL_READS]
    p = _book(tmp_path / "esc.xlsx", stored)
    before = _stored_texts(p)
    out = structure.modify_grid_structure(p, "insert_cols", "D", sheet="Data")
    assert out["route"] == "openpyxl"
    assert _stored_texts(p) == before


@pytest.mark.parametrize("route", ["raw", "rewrite"])
def test_text_that_looks_like_an_escape_is_written_as_typed(tmp_path, route):
    p = _book(tmp_path / "w.xlsx", ["x"])
    text = "_x0041_ and Ax005F_B"
    cells.set_cell(p, {"cell": "C1"}, text, sheet="Data")
    if route == "rewrite":
        structure.modify_grid_structure(p, "insert_cols", "Z", sheet="Data")
    got = cells.get_cells(p, ["C1"], sheet="Data")["cells"][0][1]
    assert got == text
    assert text in _stored_texts(p)


def test_decode_and_encode_are_inverse():
    for _raw, shown in EXCEL_READS:
        assert xltext.decode(xltext.encode(shown)) == shown
    assert xltext.encode("_x0041_") == "_x005F_x0041_"
    assert xltext.encode("a\rb") == "a_x000D_b"
    assert xltext.encode("Ax005F_B") == "Ax005F_B"


def test_a_lone_surrogate_escape_stays_text_and_round_trips(tmp_path):
    """Half a pair cannot be carried in a reply, so it reads as its text, and
    a rewrite stores it as it was."""
    p = _book(tmp_path / "lone.xlsx", ["a_xD83D_b"])
    assert cells.get_cells(p, ["A1"], sheet="Data")["cells"][0][1] == \
        "a_xD83D_b"
    structure.modify_grid_structure(p, "insert_cols", "D", sheet="Data")
    with zipfile.ZipFile(p) as z:
        body = z.read("xl/worksheets/sheet1.xml").decode("utf-8") + \
            (z.read("xl/sharedStrings.xml").decode("utf-8")
             if "xl/sharedStrings.xml" in z.namelist() else "")
    assert "a_xD83D_b" in body and "_x005F_xD83D_" not in body
