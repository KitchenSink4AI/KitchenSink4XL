"""The shared reference rewriter (sheet renames, 3D references, tables and
the totals row), every expectation taken from Excel's own result on the same
edit.

The expected values below are what a private hidden Excel produced for the
same workbook and the same edit. The two COM tests at the end re-derive them
end to end with the product's own private hidden Excel, where Excel exists.
"""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path

import openpyxl
import pytest

import family_harness as fh
from xlsx_mcp import server as srv
from xlsx_mcp.com import session as com_session
from xlsx_mcp.core import refs

_COM = com_session.com_available()[0]


def _call(tool, op, **args):
    return fh.start().call(tool, {"op": op, "args": args})


def _ok(tool, op, **args):
    payload, err, _ = _call(tool, op, **args)
    assert not err, (op, payload)
    return payload


# ------------------------------------------------------------- 3D


def _regions(p) -> str:
    srv.create_workbook(str(p), sheets=["North", "South", "West", "Sum"])
    for name in ("North", "South", "West"):
        srv.write_range(str(p), "F1", [[1], [2], [3], [4], [5]], sheet=name)
    srv.write_range(str(p), "A1", [["=SUM(North:West!F4)"], ["=West!F2"],
                                    ["=SUM(North:West!F2:F5)"], ["=North!F4"],
                                    ["=SUM(North:South!F4)"]], sheet="Sum")
    return str(p)


def test_x21_a_3d_reference_is_left_alone_by_an_edit_on_one_sheet(tmp_path):
    """Excel: rows inserted at 2 on North leave =SUM(North:West!F4) and
    =SUM(North:West!F2:F5) as they are, and move =North!F4 to F5."""
    p = _regions(tmp_path / "r.xlsx")
    _ok("xl_structure", "modify_grid_structure", path=p, action="insert_rows", at=2,
        sheet="North")
    ws = openpyxl.load_workbook(p)["Sum"]
    assert [ws[f"A{r}"].value for r in (1, 3, 4, 5)] == [
        "=SUM(North:West!F4)", "=SUM(North:West!F2:F5)", "=North!F5", "=SUM(North:South!F4)"]


# --------------------------------------------------- sheet rename


def test_x20_rename_rewrites_every_reference_as_excel_does(tmp_path):
    p = _regions(tmp_path / "r.xlsx")
    _ok("xl_structure", "manage_name", path=p, action="add", name="WestTot",
        refers_to="=West!$F$2:$F$5")
    _ok("xl_format", "manage_hyperlink", path=p, action="add", sheet="Sum",
        location="B1", target="West!A1", display="go west")
    _ok("xl_structure", "manage_data_validation", path=p, action="add", sheet="Sum",
        location="C1", dv_type="list", values="=West!$F$1:$F$5")
    _ok("xl_format", "manage_conditional_format", path=p, action="add", sheet="Sum",
        location="D1", cf_type="formula", params={"formula": "West!$F$1>0", "fill": "FFC7CE"})
    payload = _ok("xl_structure", "manage_worksheet", path=p, action="rename", sheet="West",
                  new_name="West Coast")
    assert payload["changed"]["worksheet"]["references_rewritten"]["formulas"] == 3
    wb = openpyxl.load_workbook(p)
    ws = wb["Sum"]
    # Excel's own results (oracle_refs.py, after_rename_West)
    assert ws["A1"].value == "=SUM('North:West Coast'!F4)"
    assert ws["A2"].value == "='West Coast'!F2"
    assert ws["A3"].value == "=SUM('North:West Coast'!F2:F5)"
    assert wb.defined_names["WestTot"].attr_text == "'West Coast'!$F$2:$F$5"
    assert ws.data_validations.dataValidation[0].formula1 == "'West Coast'!$F$1:$F$5"
    rules = [r for rs in ws.conditional_formatting._cf_rules.values() for r in rs]
    assert rules[0].formula == ["'West Coast'!$F$1>0"]
    # Excel leaves a hyperlink's location as typed.
    assert ws["B1"].hyperlink.location == "West!A1"


def test_x20_a_rename_to_a_plain_name_is_not_quoted(tmp_path):
    p = _regions(tmp_path / "r.xlsx")
    _ok("xl_structure", "manage_worksheet", path=p, action="rename", sheet="West",
        new_name="Pacific")
    ws = openpyxl.load_workbook(p)["Sum"]
    assert ws["A1"].value == "=SUM(North:Pacific!F4)" and ws["A2"].value == "=Pacific!F2"


def test_x20_delete_takes_excels_result(tmp_path):
    """Excel: deleting South (the end of North:South) leaves =SUM(North!F4);
    deleting North then turns =North!F4 into =#REF! and shrinks the span to
    the one sheet left. (Excel shows =#REF!F4 while the workbook is open; the
    file it saves holds #REF!, and it opens a file holding #REF!F4 only after
    a repair.)"""
    p = _regions(tmp_path / "r.xlsx")
    _ok("xl_structure", "manage_worksheet", path=p, action="rename", sheet="West",
        new_name="West Coast")
    _ok("xl_structure", "manage_worksheet", path=p, action="delete", sheet="South")
    ws = openpyxl.load_workbook(p)["Sum"]
    assert ws["A5"].value == "=SUM(North!F4)"
    assert ws["A1"].value == "=SUM('North:West Coast'!F4)"
    _ok("xl_structure", "manage_worksheet", path=p, action="delete", sheet="North")
    ws = openpyxl.load_workbook(p)["Sum"]
    assert ws["A1"].value == "=SUM('West Coast'!F4)"
    assert ws["A4"].value == "=#REF!"


@pytest.mark.parametrize("text,expected", [
    ('=A1&"West!A1"', '=A1&"West!A1"'),           # inside a string: untouched
    ("=[1]West!A1", "=[1]West!A1"),                # another workbook
    ("='[Book.xlsx]West'!A1", "='[Book.xlsx]West'!A1"),
    ("=Westward!A1", "=Westward!A1"),              # a longer name
    ("=west!A1", "='West Coast'!A1"),              # names compare without case
    ("=SUM(West:Sum!A1)", "=SUM('West Coast:Sum'!A1)"),
])
def test_rename_touches_only_qualifiers_that_name_the_sheet(text, expected):
    assert refs.rename_sheet_in_formula(text, "West", "West Coast") == expected


# ----------------------------------------------- column insert


def _budget(p) -> str:
    p = str(p)
    srv.create_workbook(p, sheets=["Budget"])
    srv.write_range(p, "A1", [["Item", "Budget", "Actual", "Var"],
                              ["a", 10, 12, "=C2-B2"], ["b", 20, 15, "=C3-B3"],
                              ["c", 30, 31, "=C4-B4"]])
    srv.manage_conditional_format(p, "add", location="A2:D4", cf_type="formula",
                                  params={"formula": "$C2>$B2", "fill": "FFC7CE"})
    srv.manage_data_validation(p, "add", location="E2", dv_type="decimal",
                               operator="between", formula1="=$C$2", formula2="=$C$4")
    srv.set_page_layout(p, print_area="A1:D4", print_title_rows="1:1",
                        print_title_cols="A:A")
    srv.manage_chart(p, "create", chart_type="bar", data="C1:C4", categories="A2:A4",
                     at_cell="G2")
    return p


def _chart_refs(p) -> list[str]:
    wb = openpyxl.load_workbook(p)
    out = []
    for ch in wb["Budget"]._charts:
        for s in ch.series:
            out += [s.tx.strRef.f if s.tx and s.tx.strRef else None,
                    s.val.numRef.f if s.val and s.val.numRef else None,
                    (s.cat.numRef or s.cat.strRef).f if s.cat else None]
    return out


def test_x48_x49_x50_a_column_insert_moves_everything_excel_moves(tmp_path):
    p = _budget(tmp_path / "b.xlsx")
    payload = _ok("xl_structure", "modify_grid_structure", path=p, action="insert_cols",
                  at="C")
    rw = payload["changed"]["structure"]["rewrites"]
    assert rw["rule_formulas"] >= 3 and rw["chart_refs"] >= 2 and rw["print_settings"] >= 1
    wb = openpyxl.load_workbook(p)
    ws = wb["Budget"]
    # Excel's own results (oracle_refs.py, col_insert_C)
    assert ws["E2"].value == "=D2-B2"
    cf = [(str(c.sqref), r.formula) for c, rs in ws.conditional_formatting._cf_rules.items()
          for r in rs]
    assert cf == [("A2:E4", ["$D2>$B2"])]
    dv = ws.data_validations.dataValidation[0]
    assert (dv.formula1, dv.formula2) == ("$D$2", "$D$4")
    assert ws.print_area == "'Budget'!$A$1:$E$4"
    assert (ws.print_title_rows, ws.print_title_cols) == ("$1:$1", "$A:$A")
    refs_now = [r for r in _chart_refs(p) if r]
    assert all("$C$" not in r for r in refs_now), refs_now
    assert any("$D$2:$D$4" in r for r in refs_now), refs_now


def test_a_column_delete_shrinks_the_print_area_and_drops_a_deleted_category(tmp_path):
    p = _budget(tmp_path / "b.xlsx")
    _ok("xl_structure", "modify_grid_structure", path=p, action="delete_cols", at="A")
    ws = openpyxl.load_workbook(p)["Budget"]
    assert ws.print_area == "'Budget'!$A$1:$C$4"
    # Excel removed the print-title column (its whole range was deleted).
    assert ws.print_title_cols is None
    cf = [r.formula for rs in ws.conditional_formatting._cf_rules.values() for r in rs]
    assert cf == [["$B2>$A2"]]
    ch = ws._charts[0]
    assert ch.series[0].cat is None  # Excel: =SERIES(Budget!$C$1,,Budget!$C$2:$C$4,1)
    assert ch.series[0].val.numRef.f.endswith("$B$2:$B$4")


# ------------------------------------------- tables


def test_the_inventory_job_replayed(tmp_path):
    """J15 (calls.jsonl seq 374-387) through the surface: the typed
    shorthand is stored in Excel's form, the totals row has its SUBTOTAL
    cells, and the rename follows into the Summary sheet and a defined
    name."""
    p = str(tmp_path / "inventory.xlsx")
    steps = [
        ("xl_workbook", "create_workbook", {"sheets": ["Stock", "Summary"]}),
        ("xl_edit", "write_range", {"location": {"cell": "A1", "sheet": "Stock"}, "data": [
            ["SKU", "Item", "Qty", "Cost"], ["S-01", "Bolt", 100, 0.25],
            ["S-02", "Nut", 250, 0.1], ["S-03", "Washer", 500, 0.05]]}),
        ("xl_structure", "create_table", {"location": {"range": "A1:D4", "sheet": "Stock"},
                                          "name": "Stock_Tbl"}),
        ("xl_structure", "manage_table", {"name": "Stock_Tbl", "action": "add_column",
                                          "column": "Value"}),
        ("xl_edit", "set_formula", {"location": {"range": "E2:E4", "sheet": "Stock"},
                                    "formula": "=[@Qty]*[@Cost]"}),
        ("xl_structure", "manage_table", {"name": "Stock_Tbl", "action": "add_row", "values": [
            ["S-04", "Screw", 300, 0.08, "=[@Qty]*[@Cost]"], ["S-05", "Rivet", 80, 0.3, None]]}),
        ("xl_structure", "manage_table", {"name": "Stock_Tbl", "action": "toggle_totals",
                                          "on": True, "totals": {"Qty": "sum", "Value": "sum"}}),
        ("xl_edit", "write_range", {"location": {"cell": "A1", "sheet": "Summary"}, "data": [
            ["Total stock value", "=SUM(Stock_Tbl[Value])"], ["Item count", "=ROWS(Stock_Tbl)"],
            ["Max qty", "=MAX(Stock_Tbl[Qty])"]]}),
        ("xl_structure", "manage_name", {"action": "add", "name": "Vals",
                                         "refers_to": "=Stock_Tbl[Value]"}),
        ("xl_structure", "manage_table", {"name": "Stock_Tbl", "action": "rename",
                                          "new_name": "Inventory"}),
    ]
    for tool, op, args in steps:
        _ok(tool, op, path=p, **args)
    wb = openpyxl.load_workbook(p)
    st, sm = wb["Stock"], wb["Summary"]
    assert st["E2"].value == "=Inventory[[#This Row],[Qty]]*Inventory[[#This Row],[Cost]]"
    assert st["E5"].value == "=Inventory[[#This Row],[Qty]]*Inventory[[#This Row],[Cost]]"
    # The totals row (row 7): Excel's label and SUBTOTAL(109, T[col]) cells.
    assert st["A7"].value == "Total"
    assert st["C7"].value == "=SUBTOTAL(109,Inventory[Qty])"
    assert st["E7"].value == "=SUBTOTAL(109,Inventory[Value])"
    assert [sm[f"B{r}"].value for r in (1, 2, 3)] == [
        "=SUM(Inventory[Value])", "=ROWS(Inventory)", "=MAX(Inventory[Qty])"]
    assert wb.defined_names["Vals"].attr_text == "Inventory[Value]"


def test_x12_a_totals_row_never_overwrites_data_under_the_table(tmp_path):
    p = str(tmp_path / "t.xlsx")
    srv.create_workbook(p)
    srv.write_range(p, "A1", [["Item", "Qty"], ["a", 1], ["b", 2], [None, "keep me"]])
    payload, err, _ = _call("xl_structure", "create_table", path=p, location="A1:B3",
                            name="T", totals_row=True, totals={"Qty": "sum"})
    assert err and payload["error"]["code"] == "CONFLICT", payload
    assert "insert_rows" in payload["error"]["message"]
    assert openpyxl.load_workbook(p).active["B4"].value == "keep me"


def test_x12_toggling_totals_off_takes_its_cells_with_it(tmp_path):
    p = str(tmp_path / "t.xlsx")
    srv.create_workbook(p)
    srv.write_range(p, "A1", [["Item", "Qty"], ["a", 1], ["b", 2]])
    _ok("xl_structure", "create_table", path=p, location="A1:B3", name="T",
        totals_row=True, totals={"Qty": "count"})
    ws = openpyxl.load_workbook(p).active
    assert ws["B4"].value == "=SUBTOTAL(103,T[Qty])" and ws["A4"].value == "Total"
    _ok("xl_structure", "manage_table", path=p, name="T", action="toggle_totals", on=False)
    ws = openpyxl.load_workbook(p).active
    assert ws["A4"].value is None and ws["B4"].value is None


# ------------------------------------------------- Excel as the judge


def _excel_edit(path: str, action) -> None:
    """Make the same edit in the product's own private hidden Excel."""
    def job(manager):
        w = manager.acquire()
        prior = com_session._assert_hygiene(w.app)
        wb = None
        try:
            wb = com_session.open_workbook(w.app, path)
            action(wb)
            wb.Save()
        finally:
            if wb is not None:
                wb.Close(False)
            com_session._restore_hygiene(w.app, *prior)
            manager.release_to_pool(w)
    com_session.run_com("oracle", job, timeout=120)


@pytest.mark.com
@pytest.mark.skipif(not _COM, reason="needs Excel (private hidden pool)")
@pytest.mark.timeout(400)
def test_excel_agrees_with_the_column_insert(tmp_path):
    ours = _budget(tmp_path / "ours.xlsx")
    theirs = str(tmp_path / "excel.xlsx")
    shutil.copy2(ours, theirs)
    _ok("xl_structure", "modify_grid_structure", path=ours, action="insert_cols", at="C")
    _excel_edit(theirs, lambda wb: wb.Worksheets("Budget").Columns("C").Insert())
    a, b = openpyxl.load_workbook(ours)["Budget"], openpyxl.load_workbook(theirs)["Budget"]
    assert [a[f"E{r}"].value for r in (2, 3, 4)] == [b[f"E{r}"].value for r in (2, 3, 4)]
    fa = [r.formula for rs in a.conditional_formatting._cf_rules.values() for r in rs]
    fb = [r.formula for rs in b.conditional_formatting._cf_rules.values() for r in rs]
    assert fa == fb
    assert a.print_area == b.print_area
    # Excel re-saves chart references without quotes and with $ anchors
    # (chart references are absolute either way); compare the cells named.
    ref = lambda f: f.replace("'", "").replace("$", "") if f else f  # noqa: E731
    assert [ref(x) for x in _chart_refs(ours)] == [ref(x) for x in _chart_refs(theirs)]
    assert srv.com_validate_opens_clean(ours)["opens_clean"] is True


@pytest.mark.com
@pytest.mark.skipif(not _COM, reason="needs Excel (private hidden pool)")
@pytest.mark.timeout(400)
def test_excel_reads_the_inventory_job(tmp_path):
    p = str(tmp_path / "inventory.xlsx")
    srv.create_workbook(p, sheets=["Stock", "Summary"])
    srv.write_range(p, "A1", [["SKU", "Qty", "Cost"], ["a", 100, 0.25], ["b", 250, 0.1]],
                    sheet="Stock")
    srv.create_table(p, {"range": "A1:C3", "sheet": "Stock"}, "Stock_Tbl")
    srv.manage_table(p, "Stock_Tbl", "add_column", column="Value")
    srv.set_formula(p, {"range": "D2:D3", "sheet": "Stock"}, "=[@Qty]*[@Cost]")
    srv.manage_table(p, "Stock_Tbl", "toggle_totals", on=True, totals={"Value": "sum"})
    srv.set_formula(p, {"cell": "A1", "sheet": "Summary"}, "=SUM(Stock_Tbl[Value])")
    srv.manage_table(p, "Stock_Tbl", "rename", new_name="Inventory")
    assert srv.com_validate_opens_clean(p)["opens_clean"] is True
    srv.recalculate(p, engine="com")
    st = srv.read_range(p, {"range": "D2:D4", "sheet": "Stock"})["values"]
    sm = srv.read_range(p, {"cell": "A1", "sheet": "Summary"})["values"]
    assert [r[0] for r in st] == [25, 25, 50] and sm[0][0] == 50


@pytest.mark.parametrize("text,expected", [
    ("=ROWS(Stock_Tbl)", "=ROWS(Inventory)"),
    ("=SUM(Stock_Tbl[x'#])+Stock_Tbl[y'#]", "=SUM(Inventory[x'#])+Inventory[y'#]"),
    ('=INDIRECT("Stock_Tbl[Value]")', '=INDIRECT("Stock_Tbl[Value]")'),
    ("='Stock_Tbl'!A1+Stock_Tbl2[a]", "='Stock_Tbl'!A1+Stock_Tbl2[a]"),
    ("=Stock_Tbl!A1", "=Stock_Tbl!A1"),
    ("=stock_tbl[Value]", "=Inventory[Value]"),
])
def test_x47_a_table_rename_touches_only_the_table_name(text, expected):
    wb = openpyxl.Workbook()
    wb.active["A1"] = text
    refs.rename_table_refs(wb, "Stock_Tbl", "Inventory")
    assert wb.active["A1"].value == expected
