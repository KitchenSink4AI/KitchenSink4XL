"""A formula that returns an array is written as a dynamic array.

Written as a plain <f>, =UNIQUE(A1:A5) is a legacy formula: Excel opens it as
=@UNIQUE(A1:A5), one value, no spill, no warning. Both save routes did that.
Now the formula cell is stored as Excel 365 stores it (cm pointing at the
XLDAPR entry of xl/metadata.xml, <f t="array" ref=anchor>). Proven in a hidden
Excel on both routes and on an Excel-made workbook: the formula reads back
without @ and spills.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.core import dynarray
from xlsx_mcp.core.errors import UnsupportedStructure
from xlsx_mcp.ops import cells, formulas


def _book(tmp_path, name="w.xlsx"):
    p = str(tmp_path / name)
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, v in enumerate(["a", "b", "a", "c", "b"], 1):
        ws.cell(i, 1, v)
    wb.save(p)
    return p


def _cell_xml(p, coord):
    data = zipfile.ZipFile(p).read("xl/worksheets/sheet1.xml")
    m = re.search(rb'<c r="' + coord.encode() + rb'"[^>]*>.*?</c>', data)
    return m.group(0) if m else None


@pytest.mark.parametrize("route_off", [False, True], ids=["raw", "rewrite"])
def test_spilling_formula_is_stored_as_a_dynamic_array(tmp_path, monkeypatch, route_off):
    if route_off:
        monkeypatch.setenv("KS4XL_SURGICAL_SAVE", "off")
    p = _book(tmp_path)
    out = formulas.set_formula(p, "C1", "=UNIQUE(A1:A5)")
    assert out["route"] == ("openpyxl" if route_off else "raw_ooxml")
    assert out["changed"]["dynamic_array_formulas"] == 1
    xml = _cell_xml(p, "C1")
    assert b'cm="1"' in xml and b'<f t="array" ref="C1">_xlfn.UNIQUE(A1:A5)</f>' in xml
    z = zipfile.ZipFile(p)
    meta = z.read("xl/metadata.xml")
    assert b'name="XLDAPR"' in meta and b'fDynamic="1"' in meta
    assert b"sheetMetadata" in z.read("xl/_rels/workbook.xml.rels")
    assert b"/xl/metadata.xml" in z.read("[Content_Types].xml")
    assert cells.read_range(p, "C1", values="formula")["values"] == [["=UNIQUE(A1:A5)"]]


def test_a_second_write_reuses_the_entry(tmp_path):
    p = _book(tmp_path)
    formulas.set_formula(p, "C1", "=UNIQUE(A1:A5)")
    formulas.set_formula(p, "E1", "=SEQUENCE(3)")
    z = zipfile.ZipFile(p)
    assert z.namelist().count("xl/metadata.xml") == 1
    assert z.read("xl/_rels/workbook.xml.rels").count(b"sheetMetadata") == 1
    assert b'cm="1"' in _cell_xml(p, "E1") and b'cm="1"' in _cell_xml(p, "C1")


def test_an_ordinary_formula_is_left_alone(tmp_path):
    p = _book(tmp_path)
    out = formulas.set_formula(p, "C1", '=COUNTIF(A1:A5,"UNIQUE(x)")')
    assert "dynamic_array_formulas" not in out["changed"]
    assert b"cm=" not in _cell_xml(p, "C1")
    assert "xl/metadata.xml" not in zipfile.ZipFile(p).namelist()


@pytest.mark.parametrize("text,spills", [
    ("=UNIQUE(A1:A5)", True), ("=_xlfn.SORT(A1:A5)", True), ("=SUM(FILTER(A:A,B:B>1))", True),
    ("=_xlfn._xlws.FILTER(A1:A3,B1:B3)", True), ("=SUM(A1:A5)", False),
    ('="SEQUENCE(3)"', False), ("=MyUnique(A1)", False)])
def test_spilling_detection(text, spills):
    assert dynarray.is_spilling(text) is spills


def test_metadata_without_a_dynamic_array_entry_refuses_unchanged(tmp_path):
    p = _book(tmp_path)
    src = p + ".src"
    Path(p).rename(src)
    rich = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b'<metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/>'
            b'</metadataTypes><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>')
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in zin.infolist():
            b = zin.read(i.filename)
            if i.filename == "xl/_rels/workbook.xml.rels":
                b = b.replace(b"</Relationships>", b'<Relationship Id="rId99" Type="http://schemas.'
                              b'openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" '
                              b'Target="metadata.xml"/></Relationships>')
            if i.filename == "[Content_Types].xml":
                b = b.replace(b"</Types>", b'<Override PartName="/xl/metadata.xml" ContentType="'
                              + dynarray.CT_METADATA.encode() + b'"/></Types>')
            zout.writestr(i, b)
        zout.writestr("xl/metadata.xml", rich)
    before = Path(p).read_bytes()
    with pytest.raises(UnsupportedStructure, match="dynamic array"):
        formulas.set_formula(p, "C1", "=UNIQUE(A1:A5)")
    assert Path(p).read_bytes() == before


@pytest.mark.parametrize("order", ["formula_first", "insert_first"])
def test_formula_and_row_insert_in_one_save(tmp_path, order):
    from xlsx_mcp.core import refs
    from xlsx_mcp.core.package import WorkbookPackage
    p = _book(tmp_path)
    pkg = WorkbookPackage.open(p)
    edit = refs.RefEdit(sheet="Sheet", kind=refs.INSERT_ROWS, index=1, count=1)
    if order == "formula_first":
        pkg.set_formula("Sheet", "C1", "=UNIQUE(A1:A5)")
        pkg.modify_structure(edit)
        where = "C2"
    else:
        pkg.modify_structure(edit)
        pkg.set_formula("Sheet", "C1", "=UNIQUE(A2:A6)")
        where = "C1"
    pkg.save()
    xml = _cell_xml(p, where)
    assert xml is not None and b'cm="1"' in xml and b't="array"' in xml, xml
