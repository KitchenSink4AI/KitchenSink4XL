"""Four ways an edit or a read could answer ok while the workbook was wrong
(every workbook here is made up and built by the test).

A macro-enabled template (.xltm): an edit with allow_loss saved it with the
plain template content type, and Excel refuses a file whose extension and
workbook type disagree. The VBA project and the macro-enabled type are kept
for .xlsm and .xltm alike, and the verify step refuses any package whose
workbook type does not match the file name.

Relationship targets spelled with a percent code for a plain character
(worksheets/sheet%31.xml is worksheets/sheet1.xml): the workbook read as ok
with no sheets. The target is read the way Excel reads it; a listed sheet whose
target names no part is refused by name.

A table that grows: a total written as a plain range over the table's column
(SUM(D2:D13)), a defined name, a conditional-format range or a chart series
over the data was left at its old size when a row was added (or inserted at
the first data row), where Excel extends it.

A workbook that validate calls unopenable (a macro part that is not a macro
project): an edit of it answered verified. It is refused with validate's
reason.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import openpyxl
import pytest
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table

from xlsx_mcp.core import verify
from xlsx_mcp.core import xmlsafe as _xmlsafe
from xlsx_mcp.core.errors import ValidationFailed, XlMcpError
from xlsx_mcp.ops import cells, lifecycle, structure, tables, validation

CFB = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 32
CT = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}
VBA_REL = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"


def _rewrite(src: Path, dst: Path, fn) -> Path:
    """Copy a package, giving fn(name, data) -> data|None each part (None
    drops it); fn(None, None) -> iterable of (name, data) to add."""
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = fn(info.filename, zin.read(info.filename))
            if data is not None:
                zout.writestr(info.filename, data)
        for name, data in fn(None, None) or ():
            zout.writestr(name, data)
    return dst


def _plain_book(path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    for i in range(1, 9):
        ws.cell(i, 1, f"Row {i}")
        ws.cell(i, 2, i * 10)
    ws["D1"] = "=SUM(B1:B8)"
    wb.save(path)
    return path


def _macro_book(tmp_path: Path, ext: str, *, workbook_type: str | None = None,
                vba: bytes | None = CFB, force_vba: bool = False) -> Path:
    """A made-up workbook named *.ext whose workbook part has the content type
    for ext (or workbook_type), carrying a macro project when ext is macro."""
    plain = _plain_book(tmp_path / "plain_src.xlsx")
    ct = workbook_type or CT[ext]
    macro = vba is not None and (ext in ("xlsm", "xltm") or force_vba)

    def fn(name, data):
        if name is None:
            return [("xl/vbaProject.bin", vba)] if macro else []
        if name == "[Content_Types].xml":
            text = data.decode("utf-8")
            text = text.replace(CT["xlsx"], ct)
            if macro:
                text = text.replace(
                    "</Types>",
                    '<Default Extension="bin" ContentType="application/'
                    'vnd.ms-office.vbaProject"/></Types>')
            return text.encode("utf-8")
        if name == "xl/_rels/workbook.xml.rels" and macro:
            return data.replace(
                b"</Relationships>",
                b'<Relationship Id="rId99" Type="' + VBA_REL.encode()
                + b'" Target="vbaProject.bin"/></Relationships>')
        return data

    return _rewrite(plain, tmp_path / f"book.{ext}", fn)


def _workbook_type(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        text = zf.read("[Content_Types].xml").decode("utf-8")
    return re.search(r'PartName="/xl/workbook\.xml" ContentType="([^"]+)"',
                     text).group(1)



@pytest.mark.parametrize("ext", ["xlsx", "xlsm", "xltx", "xltm"])
def test_every_workbook_kind_keeps_its_own_content_type_through_an_edit(tmp_path, ext):
    p = _macro_book(tmp_path, ext)
    assert _workbook_type(p) == CT[ext]
    res = structure.modify_grid_structure(
        str(p), "insert_rows", 3, sheet="Report", allow_loss=True,
        backup=False, verify_com=False)
    assert res["ok"] is True
    assert _workbook_type(p) == CT[ext]
    with zipfile.ZipFile(p) as zf:
        has_vba = "xl/vbaProject.bin" in zf.namelist()
    assert has_vba is (ext in ("xlsm", "xltm"))
    assert verify.package_refusals(str(p)) == []


@pytest.mark.parametrize("op", ["insert_rows", "delete_rows", "insert_cols", "delete_cols"])
def test_a_macro_template_keeps_its_macro_project_without_allow_loss(tmp_path, op):
    p = _macro_book(tmp_path, "xltm")
    res = structure.modify_grid_structure(str(p), op, 3, sheet="Report",
                                          backup=False, verify_com=False)
    assert res["ok"] is True
    assert _workbook_type(p) == CT["xltm"]
    with zipfile.ZipFile(p) as zf:
        assert zf.read("xl/vbaProject.bin") == CFB


def test_a_cell_edit_of_a_macro_template_stays_a_macro_template(tmp_path):
    p = _macro_book(tmp_path, "xltm")
    res = cells.set_cell(str(p), "A2", "changed", sheet="Report", backup=False,
                         verify_com=False)
    assert res["ok"] is True
    assert _workbook_type(p) == CT["xltm"]


@pytest.mark.parametrize("ext,wrong", [
    ("xltm", "xltx"), ("xltm", "xlsx"), ("xlsm", "xlsx"), ("xlsx", "xlsm"),
    ("xltx", "xltm"), ("xlsx", "xltx")])
def test_the_workbook_type_must_match_the_file_name(tmp_path, ext, wrong):
    p = _macro_book(tmp_path, ext, workbook_type=CT[wrong], vba=None)
    why = " ".join(verify.package_refusals(str(p)))
    assert f".{ext}" in why and "workbook part" in why


def test_a_file_whose_name_and_type_disagree_is_not_edited_as_verified(tmp_path):
    p = _macro_book(tmp_path, "xltm", workbook_type=CT["xltx"], vba=None)
    before = p.read_bytes()
    with pytest.raises(ValidationFailed) as exc:
        cells.set_cell(str(p), "A2", "changed", sheet="Report", backup=False,
                       verify_com=False)
    assert ".xltm" in str(exc.value)
    assert p.read_bytes() == before


def test_validate_names_a_workbook_type_that_does_not_match_the_name(tmp_path):
    p = _macro_book(tmp_path, "xltm", workbook_type=CT["xltx"], vba=None)
    res = validation.validate(str(p), checks=["structure"])["results"]["structure"]
    assert res["passed"] is False
    assert ".xltm" in " ".join(res["findings"]["issues"])


def test_a_good_file_of_each_kind_validates(tmp_path):
    for ext in ("xlsx", "xlsm", "xltx", "xltm"):
        d = tmp_path / ext
        d.mkdir()
        p = _macro_book(d, ext)
        res = validation.validate(str(p), checks=["structure"])["results"]["structure"]
        assert res["passed"] is True, (ext, res["findings"]["issues"])



def _respell(src: Path, dst: Path, *, target: str, entry: str | None = None,
             old_entry: str = "xl/worksheets/sheet1.xml") -> Path:
    """The sheet's workbook relationship Target spelled `target`; the zip entry
    renamed `entry` when given."""
    old_rels = "xl/worksheets/_rels/sheet1.xml.rels"

    def fn(name, data):
        if name is None:
            return []
        if name == "xl/_rels/workbook.xml.rels":
            return re.sub(rb'Target="/?(?:xl/)?worksheets/sheet1\.xml"',
                          b'Target="' + target.encode("utf-8") + b'"', data)
        return data

    out = _rewrite(src, dst, fn)
    if entry is not None:
        with zipfile.ZipFile(out) as zin:
            items = [(i.filename, zin.read(i.filename)) for i in zin.infolist()]
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in items:
                if name == old_entry:
                    name = entry
                elif name == old_rels:
                    name = old_rels.replace("sheet1.xml", entry.rsplit("/", 1)[1])
                zout.writestr(name, data)
    return out


def _table_book(path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.append(["Site", "Pallets"])
    for i in range(1, 5):
        ws.append([f"Site {i}", i * 7])
    ws.add_table(Table(displayName="Orders_t", ref="A1:B5"))
    wb.save(path)
    return path


@pytest.mark.parametrize("target", [
    "worksheets/sheet%31.xml", "/xl/worksheets/sheet%31.xml", "worksheets/%73heet1.xml",
    "worksheets/sheet1%2Exml"])
def test_a_target_spelled_with_a_percent_code_for_a_plain_character_reads(tmp_path, target):
    src = _table_book(tmp_path / "src.xlsx")
    p = _respell(src, tmp_path / "enc.xlsx", target=target)
    meta = lifecycle.get_workbook_metadata(str(p))
    assert meta["sheet_count"] == 1
    assert meta["sheets"][0]["name"] == "Orders"
    got = cells.read_range(str(p), "A2:B3", sheet="Orders")
    assert "Site 1" in str(got)


@pytest.mark.parametrize("route", ["cell_edit", "structure"])
def test_an_edit_of_a_workbook_with_encoded_targets_keeps_the_sheet_and_its_table(tmp_path, route):
    src = _table_book(tmp_path / "src.xlsx")
    p = _respell(src, tmp_path / "enc.xlsx", target="worksheets/sheet%31.xml")
    if route == "cell_edit":
        res = cells.set_cell(str(p), "B2", 99, sheet="Orders", backup=False,
                             verify_com=False)
    else:
        res = structure.modify_grid_structure(str(p), "insert_rows", 3, sheet="Orders",
                                              backup=False, verify_com=False)
    assert res["ok"] is True
    wb = _xmlsafe.load_workbook(str(p))
    assert wb.sheetnames == ["Orders"]
    assert "Orders_t" in wb["Orders"].tables
    assert wb["Orders"]["A1"].value == "Site"


def test_a_target_with_a_space_code_is_refused_by_name_not_read_as_empty(tmp_path):
    src = _table_book(tmp_path / "src.xlsx")
    p = _respell(src, tmp_path / "enc.xlsx", target="worksheets/sheet%201.xml",
                 entry="xl/worksheets/sheet 1.xml")
    with pytest.raises(XlMcpError) as exc:
        lifecycle.get_workbook_metadata(str(p))
    assert "Orders" in str(exc.value) and "sheet%201.xml" in str(exc.value)
    before = p.read_bytes()
    with pytest.raises(XlMcpError):
        cells.set_cell(str(p), "B2", 1, sheet="Orders", backup=False, verify_com=False)
    assert p.read_bytes() == before


def test_a_non_ascii_code_that_names_no_part_is_refused_by_name(tmp_path):
    src = _table_book(tmp_path / "src.xlsx")
    p = _respell(src, tmp_path / "enc.xlsx", target="worksheets/%E8%A1%A81.xml",
                 entry="xl/worksheets/表" + "1.xml")
    with pytest.raises(XlMcpError) as exc:
        lifecycle.get_workbook_metadata(str(p))
    assert "Orders" in str(exc.value)


def test_a_sheet_whose_target_names_no_part_is_refused_never_read_as_empty(tmp_path):
    src = _table_book(tmp_path / "src.xlsx")
    p = _respell(src, tmp_path / "gone.xlsx", target="worksheets/missing.xml")
    with pytest.raises(XlMcpError) as exc:
        lifecycle.get_workbook_metadata(str(p))
    assert "Orders" in str(exc.value) and "missing.xml" in str(exc.value)


def test_a_plain_workbook_still_reads_and_a_percent_in_a_real_name_is_kept(tmp_path):
    p = _table_book(tmp_path / "plain.xlsx")
    assert lifecycle.get_workbook_metadata(str(p))["sheet_count"] == 1
    # an entry whose real name carries the code is found by its own spelling
    q = _respell(p, tmp_path / "named.xlsx", target="worksheets/sheet%31.xml",
                 entry="xl/worksheets/sheet%31.xml")
    assert lifecycle.get_workbook_metadata(str(q))["sheet_count"] == 1



def _ledger(path: Path, *, totals: bool = False) -> Path:
    """A made-up ledger: a table A3:D13 (header row 3, ten data rows), a total
    and a name over its Line column, a rule over it, a chart over it."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.append([])
    ws.append([])
    ws.append(["Site", "Pallets", "Price", "Line"])
    for i in range(1, 11):
        r = 3 + i
        ws.append([f"Site {i}", i, 2.5, f"=B{r}*C{r}"])
    last = 13
    if totals:
        ws["A14"] = "Total"
        last = 14
    t = Table(displayName="OrdersT", ref=f"A3:D{last}")
    if totals:
        t.totalsRowCount = 1
        t.totalsRowShown = True
    ws.add_table(t)
    ws["F1"] = "=SUM(D4:D13)"
    ws["F2"] = "=SUM(C4:D13)"
    ws["F3"] = "=SUM(D5:D13)"
    ws["G1"] = "=SUM(D4:D12)"
    ws["G2"] = "=SUM(H4:H13)"
    ws["G3"] = "=SUM(D4:D14)"
    ws["G4"] = "=SUM(D3:D13)"
    ws["G5"] = "=D13"
    ws["G6"] = "=SUM($D$4:$D$13)"
    ws["G7"] = "=SUM(Orders!D4:D13)"
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "=SUM(Orders!D4:D13)"
    ws2["A2"] = "=AVERAGE(Orders!B4:B13)"
    wb.defined_names["LineAmount"] = DefinedName("LineAmount", attr_text="Orders!$D$4:$D$13")
    wb.defined_names["PartialAmount"] = DefinedName("PartialAmount", attr_text="Orders!$D$5:$D$13")
    ws.conditional_formatting.add("D4:D13", CellIsRule(
        operator="greaterThan", formula=["5"], fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add("H4:H13", CellIsRule(
        operator="greaterThan", formula=["5"], fill=PatternFill("solid", bgColor="FFC7CE")))
    ch = BarChart()
    ch.add_data(Reference(ws, min_col=4, min_row=4, max_row=13))
    ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=13))
    ws.add_chart(ch, "J2")
    wb.save(path)
    return path


def _facts(path: Path) -> dict:
    wb = openpyxl.load_workbook(path)
    ws = wb["Orders"]
    return {
        "cells": {a: ws[a].value for a in ("F1", "F2", "F3", "G1", "G2", "G3", "G4", "G5", "G6", "G7")},
        "summary": [wb["Summary"]["A1"].value, wb["Summary"]["A2"].value],
        "names": {n: d.attr_text for n, d in wb.defined_names.items()},
        "cf": sorted(str(cf.sqref) for cf in ws.conditional_formatting),
        "series": [(s.val.numRef.f, s.cat.numRef.f if s.cat and s.cat.numRef else
                    s.cat.strRef.f if s.cat and s.cat.strRef else None)
                   for c in ws._charts for s in c.series],
        "table": ws.tables["OrdersT"].ref,
    }


def test_adding_a_row_extends_what_covers_exactly_the_tables_column(tmp_path):
    p = _ledger(tmp_path / "ledger.xlsx")
    res = tables.manage_table(str(p), "OrdersT", "add_row", values=["Site 11", 11, 2.5],
                              backup=False, verify_com=False)
    assert res["ok"] is True
    f = _facts(p)
    assert f["table"] == "A3:D14"
    c = f["cells"]
    assert c["F1"] == "=SUM(D4:D14)"            # exactly the Line column: extended
    assert c["F2"] == "=SUM(C4:D14)"            # two whole columns of the data: extended
    assert c["F3"] == "=SUM(D5:D13)"            # only part of the data: left alone
    assert c["G1"] == "=SUM(D4:D12)"            # one row short: left alone
    assert c["G2"] == "=SUM(H4:H13)"            # a column outside the table: left alone
    assert c["G3"] == "=SUM(D4:D15)"            # reaches past the data: the ordinary insert rule
    assert c["G4"] == "=SUM(D3:D14)"            # header and all the data: extended
    assert c["G5"] == "=D13"
    assert c["G6"] == "=SUM($D$4:$D$14)"
    assert c["G7"] == "=SUM(Orders!D4:D14)"
    assert f["summary"] == ["=SUM(Orders!D4:D14)", "=AVERAGE(Orders!B4:B14)"]
    assert f["names"]["LineAmount"] == "Orders!$D$4:$D$14"
    assert f["names"]["PartialAmount"] == "Orders!$D$5:$D$13"
    assert f["cf"] == ["D4:D14", "H4:H13"]
    assert f["series"][0][0].replace("'", "").endswith("Orders!$D$4:$D$14")
    assert f["series"][0][1].replace("'", "").endswith("Orders!$A$4:$A$14")


def test_adding_several_rows_extends_by_that_many(tmp_path):
    p = _ledger(tmp_path / "ledger.xlsx")
    tables.manage_table(str(p), "OrdersT", "add_row",
                        values=[["Site 11", 11, 2.5], ["Site 12", 12, 2.5], ["Site 13", 13, 2.5]],
                        backup=False, verify_com=False)
    f = _facts(p)
    assert f["table"] == "A3:D16"
    assert f["cells"]["F1"] == "=SUM(D4:D16)"
    assert f["names"]["LineAmount"] == "Orders!$D$4:$D$16"


def test_adding_a_row_to_a_table_with_a_totals_row_extends_the_data_ranges(tmp_path):
    p = _ledger(tmp_path / "ledger.xlsx", totals=True)
    tables.manage_table(str(p), "OrdersT", "add_row", values=["Site 11", 11, 2.5],
                        backup=False, verify_com=False)
    f = _facts(p)
    assert f["table"] == "A3:D15"
    assert f["cells"]["F1"] == "=SUM(D4:D14)"
    assert f["names"]["LineAmount"] == "Orders!$D$4:$D$14"
    assert f["cf"][0] == "D4:D14"


def test_a_row_inserted_at_the_totals_row_grows_the_table_and_what_covers_its_data(tmp_path):
    p = _ledger(tmp_path / "ledger.xlsx", totals=True)
    structure.modify_grid_structure(str(p), "insert_rows", 14, sheet="Orders",
                                    backup=False, verify_com=False, allow_loss=True)
    f = _facts(p)
    assert f["table"] == "A3:D15"
    assert f["cells"]["F1"] == "=SUM(D4:D14)"
    assert f["names"]["LineAmount"] == "Orders!$D$4:$D$14"


def test_a_row_inserted_below_a_table_with_no_totals_row_leaves_it_alone(tmp_path):
    p = _ledger(tmp_path / "ledger.xlsx")
    structure.modify_grid_structure(str(p), "insert_rows", 14, sheet="Orders",
                                    backup=False, verify_com=False, allow_loss=True)
    f = _facts(p)
    assert f["table"] == "A3:D13"
    assert f["cells"]["F1"] == "=SUM(D4:D13)"


def test_a_range_over_the_data_and_the_totals_row_keeps_its_start_at_the_first_data_row(tmp_path):
    p = _ledger(tmp_path / "ledger.xlsx", totals=True)
    structure.modify_grid_structure(str(p), "insert_rows", 4, sheet="Orders",
                                    backup=False, verify_com=False, allow_loss=True)
    # G3 is SUM(D4:D14): over all the data and the totals row
    assert _facts(p)["cells"]["G3"] == "=SUM(D4:D15)"


def test_a_row_inserted_at_the_first_data_row_extends_instead_of_shifting(tmp_path):
    p = _ledger(tmp_path / "ledger.xlsx")
    structure.modify_grid_structure(str(p), "insert_rows", 4, sheet="Orders",
                                    backup=False, verify_com=False, allow_loss=True)
    f = _facts(p)
    assert f["table"] == "A3:D14"
    assert f["cells"]["F1"] == "=SUM(D4:D14)"
    assert f["cells"]["F3"] == "=SUM(D6:D14)"      # not the whole column: shifts down
    assert f["names"]["LineAmount"] == "Orders!$D$4:$D$14"
    assert f["names"]["PartialAmount"] == "Orders!$D$6:$D$14"
    assert f["cf"][0] == "D4:D14"
    assert f["series"][0][0].replace("'", "").endswith("Orders!$D$4:$D$14")


@pytest.mark.parametrize("at,expect", [(7, "=SUM(D4:D14)"), (13, "=SUM(D4:D14)"),
                                       (14, "=SUM(D4:D13)"), (3, "=SUM(D5:D14)")])
def test_other_inserts_keep_the_ordinary_rule(tmp_path, at, expect):
    p = _ledger(tmp_path / "ledger.xlsx")
    structure.modify_grid_structure(str(p), "insert_rows", at, sheet="Orders",
                                    backup=False, verify_com=False, allow_loss=True)
    assert _facts(p)["cells"]["F1"] == expect


def test_a_pivot_source_over_the_data_that_is_not_extended_is_named(tmp_path):
    from xlsx_mcp.core import refs
    wb = openpyxl.load_workbook(_ledger(tmp_path / "ledger.xlsx"))
    ws = wb["Orders"]
    ws.print_area = "A4:D13"
    edit = refs.RefEdit("Orders", refs.INSERT_ROWS, index=14, count=1,
                        grow=((3, 4, 13, 13, 1, 4),))
    report = refs.rewrite_workbook(wb, edit)
    assert any("print area" in n and "was not extended" in n for n in report.not_extended)



def test_an_edit_of_a_workbook_validate_calls_unopenable_is_refused_with_its_reason(tmp_path):
    p = _macro_book(tmp_path, "xlsm", vba=b"not a macro project at all")
    res = validation.validate(str(p), checks=["structure"])["results"]["structure"]
    assert res["passed"] is False
    reason = " ".join(res["findings"]["issues"])
    assert "vbaProject.bin" in reason
    before = p.read_bytes()
    with pytest.raises(ValidationFailed) as exc:
        cells.set_cell(str(p), "A2", "changed", sheet="Report", backup=False,
                       verify_com=False)
    assert "vbaProject.bin" in str(exc.value) and "not a macro project" in str(exc.value)
    assert p.read_bytes() == before


def test_a_damaged_macro_part_in_a_plain_workbook_is_the_allow_loss_route(tmp_path):
    p = _macro_book(tmp_path, "xlsx", vba=b"not a macro project at all", force_vba=True)
    # the macro part is dropped by the documented allow_loss route on a .xlsx
    res = structure.modify_grid_structure(
        str(p), "insert_rows", 3, sheet="Report", allow_loss=True, backup=False,
        verify_com=False)
    assert res["ok"] is True
    with zipfile.ZipFile(p) as zf:
        assert "xl/vbaProject.bin" not in zf.namelist()
