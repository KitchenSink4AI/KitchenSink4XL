"""A sheet-column insert or delete through a table's columns refuses, and says so.

Excel deletes the table column when a sheet column through it is deleted
(formulas that name it become #REF!), drops the table when every one of its
columns goes, and adds a table column when a sheet column is inserted inside
it. modify_grid_structure moved the cells and the table's range but kept the
table's own list of columns, so its save read the shifted header as a
rename: a delete of the column under 'Q2' was refused with "renaming column
'Q2' to 'Total' would leave ... pointing at a column that no longer
exists", which is not what was asked, and with no formula naming any column
the edit went through and saved a table Excel only opens after a repair
(range A1:C3 with four columns listed, 'Q3' twice). Inserting inside a table
was refused as if a header had been emptied.

These edits now refuse before anything changes, naming the table, its
columns and the op that does the edit (manage_table delete_column or
add_column, or to_range for the whole table). Edits beside a table still
move it as Excel does. All content is made up.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest
from openpyxl.worksheet.table import Table

from xlsx_mcp.core.errors import UnsupportedStructure
from xlsx_mcp.ops import structure, tables


def _book(tmp_path: Path, *, formulas: bool, name: str = "sales.xlsx") -> str:
    """Region, Q1, Q2, Total in A1:D4 as table Sales; with formulas, Total adds
    Q1 and Q2 and Summary!B1 sums Q2."""
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    rows = [["Region", "Q1", "Q2", "Total"], ["North", 10, 20, 30], ["South", 30, 40, 70],
            ["East", 50, 60, 110]]
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            ws.cell(r, c, v)
    if formulas:
        for r in (2, 3, 4):
            ws.cell(r, 4, "=Sales[[#This Row],[Q1]]+Sales[[#This Row],[Q2]]")
    ws.add_table(Table(displayName="Sales", ref="A1:D4"))
    summary = wb.create_sheet("Summary")
    summary["A1"] = "Q2 total"
    if formulas:
        summary["B1"] = "=SUM(Sales[Q2])"
    wb.save(p)
    return str(p)


def _refused(p: str, action: str, at, count: int = 1) -> str:
    before = Path(p).read_bytes()
    with pytest.raises(UnsupportedStructure) as ei:
        structure.modify_grid_structure(p, action, at, count, sheet="Data")
    assert Path(p).read_bytes() == before, "a refused edit must leave the file as it was"
    msg = str(ei.value)
    assert "renaming" not in msg and "would hold a number or nothing" not in msg, msg
    assert "Sales (Data!A1:D4)" in msg and "nothing was changed" in msg, msg
    return msg


@pytest.mark.parametrize("formulas", [True, False])
def test_deleting_the_sheet_column_under_a_table_column_says_what_and_how(tmp_path, formulas):
    msg = _refused(_book(tmp_path, formulas=formulas), "delete_cols", "C")
    assert "column 'Q2' of table Sales" in msg, msg
    assert "manage_table" in msg and "action='delete_column'" in msg and "column='Q2'" in msg, msg
    assert "#REF!" in msg, msg


def test_deleting_two_table_columns_names_both(tmp_path):
    msg = _refused(_book(tmp_path, formulas=True), "delete_cols", "B", 2)
    assert "columns 'Q1' and 'Q2' of table Sales" in msg, msg
    assert "column='Q1'" in msg and "column='Q2'" in msg, msg


@pytest.mark.parametrize("at,count", [("A", 4), ("A", 6)])
def test_deleting_every_column_of_a_table_points_at_to_range(tmp_path, at, count):
    msg = _refused(_book(tmp_path, formulas=True), "delete_cols", at, count)
    assert "every column of table Sales" in msg, msg
    assert "action='to_range'" in msg, msg


@pytest.mark.parametrize("at,left,right", [("B", "Region", "Q1"), ("D", "Q2", "Total")])
def test_inserting_a_sheet_column_inside_a_table_points_at_add_column(tmp_path, at, left, right):
    msg = _refused(_book(tmp_path, formulas=True), "insert_cols", at)
    assert f"between its columns '{left}' and '{right}'" in msg, msg
    assert "action='add_column'" in msg, msg
    # a weaker model once converted the table to a range on its own after this refusal: the refusal says
    # that route removes the table and needs the user's say-so
    assert "action='to_range') removes the table, so do that only if the user asks for it" in msg, msg


@pytest.mark.parametrize("action,at,ref", [
    ("insert_cols", "A", "B1:E4"),   # at the table's first column: the table moves right
    ("insert_cols", "E", "A1:D4"),   # just right of it
    ("delete_cols", "F", "A1:D4"),   # right of it
])
def test_the_column_edits_excel_takes_beside_a_table_still_work(tmp_path, action, at, ref):
    p = _book(tmp_path, formulas=True)
    structure.modify_grid_structure(p, action, at, 1, sheet="Data")
    t = openpyxl.load_workbook(p)["Data"].tables["Sales"]
    assert t.ref == ref
    assert [c.name for c in t.tableColumns] == ["Region", "Q1", "Q2", "Total"]


def test_a_column_left_of_a_table_still_deletes(tmp_path):
    p = tmp_path / "left.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for r, row in enumerate([["note", "Item", "Qty"], ["x", "a", 1], ["y", "b", 2]], start=1):
        for c, v in enumerate(row, start=1):
            ws.cell(r, c, v)
    ws.add_table(Table(displayName="Stock", ref="B1:C3"))
    wb.save(p)
    structure.modify_grid_structure(str(p), "delete_cols", "A", 1, sheet="Data")
    t = openpyxl.load_workbook(p)["Data"].tables["Stock"]
    assert t.ref == "A1:B3" and [c.name for c in t.tableColumns] == ["Item", "Qty"]


def test_the_op_the_refusal_names_does_the_edit(tmp_path):
    p = _book(tmp_path, formulas=True)
    tables.manage_table(p, "Sales", "delete_column", column="Q2")
    wb = openpyxl.load_workbook(p)
    t = wb["Data"].tables["Sales"]
    assert t.ref == "A1:C4" and [c.name for c in t.tableColumns] == ["Region", "Q1", "Total"]
    assert wb["Summary"]["B1"].value == "=SUM(#REF!)"


def test_deleting_a_table_column_says_in_words_what_now_reads_ref(tmp_path):
    # The count alone (references_now_ref_error) was missed by a weaker model, which answered only "Done":
    # the answer also says it as a warning, with the instruction to tell the user.
    p = _book(tmp_path, formulas=True)
    out = tables.manage_table(p, "Sales", "delete_column", column="Q2")
    assert out["changed"]["table"]["references_now_ref_error"] == 4, out
    said = [w for w in out["warnings"] if "#REF!" in w]
    assert len(said) == 1, out["warnings"]
    assert said[0].startswith("4 cell formulas that used the deleted column 'Q2' of table Sales now hold "
                              "#REF! in place of that column"), said
    assert "may show an error or a fallback value" in said[0], said   # IFERROR and kin return their fallback
    assert "xl_read audit_formulas lists the cell formulas" in said[0], said
    assert "Tell the user what changed" in said[0], said


def test_a_defined_name_alone_is_named_as_a_name_and_not_sent_to_audit_formulas(tmp_path):
    # audit_formulas lists worksheet formula cells only: a warning about a name must not point there
    p = _book(tmp_path, formulas=False)
    from openpyxl.workbook.defined_name import DefinedName
    wb = openpyxl.load_workbook(p)
    wb.defined_names["Q2Total"] = DefinedName("Q2Total", attr_text="SUM(Sales[Q2])")
    wb.save(p)
    out = tables.manage_table(p, "Sales", "delete_column", column="Q2")
    said = [w for w in out["warnings"] if "#REF!" in w]
    assert len(said) == 1, out["warnings"]
    assert said[0].startswith("1 defined name that used the deleted column 'Q2' of table Sales now holds #REF!"), said
    assert "audit_formulas" not in said[0], said


def test_deleting_a_table_column_nothing_names_adds_no_warning(tmp_path):
    p = _book(tmp_path, formulas=False)
    out = tables.manage_table(p, "Sales", "delete_column", column="Q2")
    assert "references_now_ref_error" not in out["changed"]["table"], out
    assert not [w for w in out["warnings"] if "#REF!" in w], out["warnings"]


def test_converting_a_table_to_a_range_says_so_in_a_warning(tmp_path):
    # a weaker model converted the table on its own after the insert refusal and answered only "Done":
    # the answer says in words that the table is gone and that the user must be told
    p = _book(tmp_path, formulas=True)
    out = tables.manage_table(p, "Sales", "to_range")
    said = [w for w in out["warnings"] if "plain range" in w]
    assert said == ["table Sales is now a plain range (Data!A1:D4): the cells keep their values, the table and "
                    "its name are gone, and 4 formulas that used it now use cell references. Tell the user that "
                    "the table was removed, unless they asked for it"], out["warnings"]


def test_converting_a_table_nothing_names_says_so_without_a_formula_clause(tmp_path):
    p = _book(tmp_path, formulas=False)
    out = tables.manage_table(p, "Sales", "to_range")
    said = [w for w in out["warnings"] if "plain range" in w]
    assert len(said) == 1 and "formula" not in said[0] and said[0].endswith("unless they asked for it"), said
