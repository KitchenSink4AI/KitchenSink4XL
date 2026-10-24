"""Workbooks Excel opens clean are read, not refused.

Two shapes Excel opens without a word made every read and every edit fail:

  - workbook protection whose attribute names differ from the schema's in
    letter case (workbookalgorithmName, workbookhashValue ...): Excel reads
    them as the protection they name (the structure is protected, and the
    password unprotects it, measured), while the reader refused the whole
    workbook (UNSUPPORTED_CONTENT, an unexpected keyword argument);
  - a TRUE/FALSE cell stored as <v>true</v> (Excel shows TRUE), which made
    the whole sheet unreadable ("could not read worksheets").

The protection status also said has_password false for a password stored
as a salted hash, which is how every current Excel stores one. All data
here is made up; the hash values are random bytes, not a real password's.
"""

from __future__ import annotations

import base64
import os
import re
import zipfile

import openpyxl

from xlsx_mcp.ops import cells, lifecycle, protection

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _package(path, sheet_cells: str, wb_extra: str = "",
             sheet_extra: str = "") -> str:
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
            'worksheet+xml"/></Types>',
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships"><Relationship Id="rId1" Type="%s/'
            'officeDocument" Target="xl/workbook.xml"/></Relationships>' % REL,
        "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="%s" xmlns:r="%s">%s<sheets><sheet name="Data" '
            'sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/>'
            '</workbook>' % (NS, REL, wb_extra),
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships"><Relationship Id="rId1" Type="%s/worksheet" '
            'Target="worksheets/sheet1.xml"/></Relationships>' % REL,
        "xl/worksheets/sheet1.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="%s"><sheetData>%s</sheetData>%s</worksheet>'
            % (NS, sheet_cells, sheet_extra),
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return str(path)


def _b64(n: int) -> str:
    return base64.b64encode(os.urandom(n)).decode("ascii")


def _protected_book(tmp_path) -> str:
    """Structure protection written in the other letter case."""
    prot = ('<workbookProtection workbookalgorithmName="SHA-512" '
            f'workbookhashValue="{_b64(64)}" workbooksaltValue="{_b64(16)}" '
            'workbookspinCount="100000" lockStructure="1" lockWindows="1"/>')
    return _package(tmp_path / "prot.xlsx",
                    '<row r="1"><c r="A1"><v>1</v></c></row>', wb_extra=prot)


def test_protection_in_another_letter_case_is_read(tmp_path):
    p = _protected_book(tmp_path)
    meta = lifecycle.get_workbook_metadata(p)
    assert [s["name"] for s in meta["sheets"]] == ["Data"]
    status = protection.set_protection(p, "status")
    assert status["workbook"]["structure_locked"] is True
    assert status["workbook"]["has_password"] is True


def test_an_edit_keeps_the_protection(tmp_path):
    p = _protected_book(tmp_path)
    cells.set_cell(p, {"cell": "B1"}, 7, sheet="Data")
    wb = openpyxl.load_workbook(p)
    try:
        assert wb.security.lockStructure
        assert wb.security.workbookHashValue
        assert wb["Data"]["B1"].value == 7
    finally:
        wb.close()


def test_a_sheet_password_stored_as_a_salted_hash_is_reported(tmp_path):
    prot = ('<sheetProtection algorithmName="SHA-512" '
            f'hashValue="{_b64(64)}" saltValue="{_b64(16)}" '
            'spinCount="100000" sheet="1" objects="1" scenarios="1"/>')
    p = _package(tmp_path / "sheet.xlsx",
                 '<row r="1"><c r="A1"><v>1</v></c></row>', sheet_extra=prot)
    status = protection.set_protection(p, "status")
    assert status["sheets"][0]["protected"] is True
    assert status["sheets"][0]["has_password"] is True


def _odd_cells(tmp_path) -> str:
    """The odd but valid cells Excel opens clean: TRUE/FALSE as words, a
    row and cells without r, ISO dates, an explicit number type, and an
    error value in lower case."""
    rows = ('<row r="1"><c r="A1" t="b"><v>true</v></c>'
            '<c r="B1" t="b"><v>false</v></c><c r="C1" t="b"><v>1</v></c>'
            '<c r="D1" t="e"><v>#n/a</v></c></row>'
            '<row><c t="inlineStr"><is><t>no address</t></is></c><c><v>7</v>'
            '</c><c t="n"><v>1E3</v></c></row>'
            '<row r="3"><c r="A3" t="d"><v>2026-02-20T00:00:00</v></c></row>')
    return _package(tmp_path / "odd.xlsx", rows)


def test_true_and_false_as_words_read_as_excel_shows_them(tmp_path):
    p = _odd_cells(tmp_path)
    got = cells.read_range(p, "A1:D2", sheet="Data")
    assert got["values"][0] == [True, False, True, "#N/A"]
    assert got["values"][1][:3] == ["no address", 7, 1000]


def test_a_workbook_with_odd_cells_can_be_edited(tmp_path):
    p = _odd_cells(tmp_path)
    out = cells.set_cell(p, {"cell": "E5"}, "ok", sheet="Data")
    assert out["verified"] is True
    got = cells.read_range(p, "A1:B1", sheet="Data")
    assert got["values"][0] == [True, False]


def test_the_case_matching_leaves_known_names_alone(tmp_path):
    """A workbook protection written with the schema's own names reads as
    before (the tolerant path only renames what does not match)."""
    prot = ('<workbookProtection workbookAlgorithmName="SHA-512" '
            f'workbookHashValue="{_b64(64)}" workbookSaltValue="{_b64(16)}" '
            'workbookSpinCount="100000" lockStructure="1"/>')
    p = _package(tmp_path / "canon.xlsx",
                 '<row r="1"><c r="A1"><v>1</v></c></row>', wb_extra=prot)
    wb = openpyxl.load_workbook(p)
    try:
        assert wb.security.lockStructure
        assert wb.security.workbookAlgorithmName == "SHA-512"
    finally:
        wb.close()
    assert re.search(r"workbookAlgorithmName", zipfile.ZipFile(p).read(
        "xl/workbook.xml").decode())
