"""How the reference rewriter reads formula text: the shapes Excel writes that a
naive scan for A1 references gets wrong.

Where a test changes text, the expected value is Excel's own result for the
same edit on a made-up workbook: the edit was made once by the product and
once by a private hidden Excel (Rows.Insert, Columns.Delete, Worksheet.Name,
Worksheet.Delete, Range.Copy), and both files were reopened in Excel and
compared formula by formula, name by name, conditional format by
conditional format; the text the product stores is the text Excel stores.
Where a test expects text to stay as it is (a typed table reference, a name
inside a quoted sheet name), it pins what the reader must leave alone.
"""

from __future__ import annotations

import pytest

from xlsx_mcp.core import refs

R = refs.RefEdit


def tf(formula, edit, sheet="Summ"):
    return refs.transpose_formula(formula, edit, sheet)


# ------------------------------------------- structured references (tables)
# A table's columns are often named like cells (Q1, FY24, FY2024, A1). Inside
# a structured reference they are column names: an insert or a delete moves
# the table, never the names.

STRUCTURED = [
    "=SUM(Sales[Q1])",
    "=SUM(Sales[FY24])",
    "=SUM(Sales[FY2024])",
    "=SUM(Sales[[Q1]:[Q2]])",
    "=Sales[[#Headers],[Q1]]",
    "=Sales[[#This Row],[Q1]]+Sales[[#This Row],[Q2]]",
    "=Sales[@Q1]*2",
    "=[@[Q1]]+[@Q2]",
    "=SUM(Sales[x'#A1])",
    "=SUM(Sales[A1])",
    "=ROWS(Sales[#All])",
    "=SUM(Sales[[#Data],[FY24]])",
    '=T[a"b]&"x"',
]


@pytest.mark.parametrize("formula", STRUCTURED)
@pytest.mark.parametrize("edit", [
    R("Data", "insert_rows", index=1, count=1),
    R("Data", "insert_cols", index=1, count=1),
    R("Data", "delete_rows", index=1, count=1),
    R("Data", "delete_cols", index=2, count=1),
    R("Summ", "insert_rows", index=1, count=2),
    R("Summ", "insert_cols", index=1, count=1),
])
def test_a_structured_reference_names_columns_not_cells(formula, edit):
    assert tf(formula, edit, "Data") == formula
    assert tf(formula, edit, "Summ") == formula
    assert refs.transpose_formula(formula[1:], edit, None) == formula[1:]  # a defined name


def test_the_cells_beside_a_structured_reference_still_move():
    e = R("Data", "insert_rows", index=1, count=1)
    assert tf("=SUM(Sales[Q1])+Data!B2", e) == "=SUM(Sales[Q1])+Data!B3"
    assert tf("=Data!J2+SUM(Sales[Q1])", e) == "=Data!J3+SUM(Sales[Q1])"
    assert tf("=SUM(Sales[Q2])*A20", R("Summ", "insert_rows", index=1, count=2)) \
        == "=SUM(Sales[Q2])*A22"


def test_copying_leaves_a_structured_reference_as_it_is():
    assert refs.offset_formula("=SUM(Sales[Q1])+A1", 1, 1) == "=SUM(Sales[Q1])+B2"
    assert refs.offset_formula("=Sales[@FY24]*2", 3, 0) == "=Sales[@FY24]*2"


# ------------------------------------------- references to other workbooks
# Excel stores [1]Data!A1 for a reference to sheet Data of the first linked
# workbook. A row inserted on THIS workbook's sheet Data never moves it.

LOCAL_INSERT = R("Data", "insert_rows", index=1, count=1)


@pytest.mark.parametrize("formula,expected", [
    ("=[1]Data!A1", "=[1]Data!A1"),
    ("=SUM([1]Data!A1:B3)+Data!A1", "=SUM([1]Data!A1:B3)+Data!A2"),
    ("=[1]Data!$A$2", "=[1]Data!$A$2"),
    ("='[1]My Data'!A1", "='[1]My Data'!A1"),
    ("=[1]!ExtTotal+Data!B2", "=[1]!ExtTotal+Data!B3"),
    ("=[1]Data!A3*Data!A3", "=[1]Data!A3*Data!A4"),
    ("=SUM([1]Data!A:A)+SUM([1]Data!$1:$3)", "=SUM([1]Data!A:A)+SUM([1]Data!$1:$3)"),
    ("=SUM([1]Jan:Mar!A1)", "=SUM([1]Jan:Mar!A1)"),
])
def test_a_reference_to_another_workbook_is_never_moved_by_an_edit_here(formula, expected):
    assert tf(formula, LOCAL_INSERT) == expected


def test_another_workbooks_name_target_is_left_alone():
    # A defined name pointing into another workbook; deleting the row that
    # holds B1 on this workbook's Data must not turn it into $#REF!.
    for kind in ("insert_rows", "delete_rows"):
        assert refs.transpose_formula("[1]Data!$B$1", R("Data", kind, index=1, count=1),
                                      None) == "[1]Data!$B$1"
    assert refs.transpose_formula("[1]Data!$B$1", R("Data", "insert_cols", index=1, count=1),
                                  None) == "[1]Data!$B$1"


def test_a_copied_formula_moves_its_references_to_other_workbooks_too():
    # Excel's Range.Copy moves relative references whichever workbook they
    # point into; absolute ones stay.
    assert refs.offset_formula("=SUM([1]Data!A1:B3)+Data!A1", 1, 1) \
        == "=SUM([1]Data!B2:C4)+Data!B2"
    assert refs.offset_formula("=[1]Data!$A$2", 1, 1) == "=[1]Data!$A$2"
    assert refs.offset_formula("='[1]My Data'!A1", 1, 0) == "='[1]My Data'!A2"


def test_renaming_or_deleting_a_local_sheet_leaves_another_workbooks_alone():
    assert refs.rename_sheet_in_formula("=[1]Data!A1+Data!A1", "Data", "Info") \
        == "=[1]Data!A1+Info!A1"
    assert refs.rename_sheet_in_formula("=[1]Jan:Mar!A1+Mar!A1", "Mar", "Apr") \
        == "=[1]Jan:Mar!A1+Apr!A1"
    assert refs.delete_sheet_in_formula("=[1]Data!A1+Data!A1", "Data", ["Data", "Summ"]) \
        == "=[1]Data!A1+#REF!"


# ------------------------------------------- quoted names with double quotes
# 'North "Q"'!A1: the double quotes belong to the sheet's name, not to a
# string; a string beside it holding the same characters is text.

NQ = 'North "Q"'


@pytest.mark.parametrize("formula,expected", [
    ("='North \"Q\"'!A1", "='North \"Q\"'!A2"),
    ("=SUM('North \"Q\"'!A1:A3)", "=SUM('North \"Q\"'!A2:A4)"),
    ("=\"North \"\"Q\"\"!A1\"&'North \"Q\"'!A2", "=\"North \"\"Q\"\"!A1\"&'North \"Q\"'!A3"),
    ("='North Q'!A2+'North \"Q\"'!A3", "='North Q'!A2+'North \"Q\"'!A4"),
    ("=LEN(\"'North \"\"Q\"\"'!A1\")+'North \"Q\"'!A4",
     "=LEN(\"'North \"\"Q\"\"'!A1\")+'North \"Q\"'!A5"),
])
def test_a_sheet_name_holding_double_quotes_is_read_whole(formula, expected):
    assert tf(formula, R(NQ, "insert_rows", index=1, count=1)) == expected


def test_an_edit_on_the_formulas_own_sheet_leaves_another_sheets_range_alone():
    # The range on 'North "Q"' once lost its second half to the scan: A3
    # was read as a cell of the formula's own sheet and moved.
    assert tf("=SUM('North \"Q\"'!A1:A3)", R("Summ", "insert_rows", index=1, count=1)) \
        == "=SUM('North \"Q\"'!A1:A3)"


def test_renaming_and_deleting_a_sheet_whose_name_holds_double_quotes():
    f = "=\"North \"\"Q\"\"!A1\"&'North \"Q\"'!A2"
    assert refs.rename_sheet_in_formula(f, NQ, 'South "Q"') \
        == "=\"North \"\"Q\"\"!A1\"&'South \"Q\"'!A2"
    assert refs.delete_sheet_in_formula(f, NQ, ["Summ", NQ]) == "=\"North \"\"Q\"\"!A1\"&#REF!"
    assert refs.offset_formula("='North \"Q\"'!A1", 1, 0) == "='North \"Q\"'!A2"


def test_nothing_inside_a_quoted_sheet_name_is_read_as_a_reference():
    e = R("Summ", "insert_rows", index=1, count=1)
    assert tf("='A1 B2 C'!Total+A1", e) == "='A1 B2 C'!Total+A2"
    assert refs.rename_sheet_in_formula("='Data x'!A1+'x Data'!A1", "Data", "Info") \
        == "='Data x'!A1+'x Data'!A1"


# ------------------------------------------- sheet names in any script
# Excel stores a sheet name made of letters, marks and digits WITHOUT quotes,
# in any script: Données!A5, 売上!B2, ยอดขาย!A2, Cost€!A1, Q１!A2.

@pytest.mark.parametrize("sheet,formula,edit_kind,expected", [
    ("Données", "=Données!A5*2", "insert_rows", "=Données!A6*2"),
    ("Données", "=COUNT(Données!A1:B7)", "insert_rows", "=COUNT(Données!A2:B8)"),
    ("Données", "=SUM(Données!A:A)", "insert_rows", "=SUM(Données!A:A)"),
    ("売上", "=SUM(売上!B2:B4)", "insert_rows", "=SUM(売上!B3:B5)"),
    ("売上", "=SUM(売上!2:3)", "insert_rows", "=SUM(売上!3:4)"),
    ("ยอดขาย", "=ยอดขาย!A2+ยอดขาย!B3", "insert_cols", "=ยอดขาย!B2+ยอดขาย!C3"),
    ("Cost€", "=Cost€!A1", "insert_cols", "=Cost€!B1"),
    ("Q１", "=Q１!A2", "insert_rows", "=Q１!A3"),
])
def test_a_sheet_named_in_another_script_is_followed(sheet, formula, edit_kind, expected):
    assert tf(formula, R(sheet, edit_kind, index=1, count=1)) == expected


def test_a_shorter_sheet_name_inside_a_longer_one_is_not_matched():
    # A sheet named es is the tail of Données: an edit on es, or its
    # rename, never touches a reference to Données.
    e = R("es", "insert_rows", index=1, count=3)
    assert tf("=Données!A5+es!A5", e) == "=Données!A5+es!A8"
    assert refs.rename_sheet_in_formula("=Données!A5+es!A5", "es", "Zed") \
        == "=Données!A5+Zed!A5"
    assert refs.delete_sheet_in_formula("=Données!A5+es!A5", "es", ["Données", "es"]) \
        == "=Données!A5+#REF!"


def test_renaming_a_sheet_to_a_name_in_another_script():
    assert refs.rename_sheet_in_formula("=Données!A5*2", "Données", "Résumé") == "=Résumé!A5*2"
    assert refs.rename_sheet_in_formula("=Données!A5*2", "Données", "Résumé 2") \
        == "='Résumé 2'!A5*2"
    assert refs.rename_sheet_in_formula("=Data!A1", "Data", "売上") == "=売上!A1"


@pytest.mark.parametrize("name,quoted", [
    ("Données", False), ("売上", False), ("Straße", False), ("ยอดขาย", False), ("Q_1", False),
    ("Sales.2024", False), ("Q1Data", False), ("North Q", True), ("it's", True), ("A1", True),
    ("SUM2", True), ("XFD1", True), ("R1C1", True), ("R1C1x", True), ("RC", True), ("C1", True),
    ("TRUE", True), ("1st", True), ("a£b", True), ("a b", True),
])
def test_a_renamed_sheet_is_quoted_where_excel_quotes_it(name, quoted):
    assert refs.qualifier([name]).startswith("'") is quoted


# ------------------------------------------- names whose insides look like cells
# αA1, A1α, 売上A1, ÀB12, Q1β, x.A1 and Q１ are defined names Excel accepts;
# no part of one is a cell.

NAMES = ["=αA1+A1", "=A1α*2+B2", "=売上A1+A1", "=ÀB12+B12", "=x.A1+A1", "=Q1β+A1", "=Q１+A1",
         "=a?A1+A1"]


@pytest.mark.parametrize("formula", NAMES)
def test_a_name_whose_inside_looks_like_a_cell_is_left_whole(formula):
    name, rest = formula[1:].split("+")
    moved = tf(formula, R("Summ", "insert_rows", index=1, count=1))
    assert moved.startswith("=" + name + "+")
    assert moved != formula  # the real reference after it moved
    copied = refs.offset_formula(formula, 1, 1)
    assert copied.startswith("=" + name + "+")


def test_names_and_their_neighbours_as_excel_rewrote_them():
    ins = R("Summ", "insert_rows", index=1, count=1)
    col = R("Summ", "insert_cols", index=1, count=1)
    assert tf("=αA1+A1", ins) == "=αA1+A2"
    assert tf("=A1α*2+B2", ins) == "=A1α*2+B3"
    assert tf("=ÀB12+B12", ins) == "=ÀB12+B13"
    assert tf("=Q1β+A1", col) == "=Q1β+B1"
    assert refs.offset_formula("=売上A1+A1", 1, 1) == "=売上A1+B2"


# ------------------------------------------- deleting a sheet
# Excel shows =#REF!F5 while the workbook is open, but the file it SAVES
# holds #REF!, and Excel opens a file holding #REF!F5 only after a repair.

ORDER = ["South", "North", "West", "Sum"]


@pytest.mark.parametrize("formula,expected", [
    ("=North!F5", "=#REF!"),
    ("=SUM(North!F2:F5)", "=SUM(#REF!)"),
    ("=North!$F$2", "=#REF!"),
    ("=SUM(North!F:F)", "=SUM(#REF!)"),
    ("=SUM(North!2:3)", "=SUM(#REF!)"),
    ("=North!F4+West!F2", "=#REF!+West!F2"),
    ("=SUM(North:West!F4)", "=SUM(West!F4)"),
    ("=SUM(South:North!F4)", "=SUM(South!F4)"),
    ("=SUM(South:West!F4)", "=SUM(South:West!F4)"),
    ("=ROWS(North!F2:F5)+ROWS(West!F2:F5)", "=ROWS(#REF!)+ROWS(West!F2:F5)"),
    ("=North!F5:F6", "=#REF!"),
    ("=SUM(North!F2,West!F3)", "=SUM(#REF!,West!F3)"),
    ("=IFERROR(North!F5,0)", "=IFERROR(#REF!,0)"),
    ('=INDIRECT("North!F5")', '=INDIRECT("North!F5")'),
    ("North!$F$2", "#REF!"),
    ("=Northern!A1+[1]North!A1", "=Northern!A1+[1]North!A1"),
])
def test_deleting_a_sheet_stores_what_excel_stores(formula, expected):
    assert refs.delete_sheet_in_formula(formula, "North", ORDER) == expected


def test_a_deleted_sheets_own_name_becomes_an_error():
    # Excel keeps the sheet's name as a workbook name holding #REF! and stores
    # =[0]!LocalRate*2; the sheet's names go with it here, so the reference
    # becomes #REF! itself. Both evaluate to #REF!.
    assert refs.delete_sheet_in_formula("=North!LocalRate*2", "North", ORDER) == "=#REF!*2"


# ------------------------------------------- row numbers are ASCII digits

def test_a_full_width_digit_is_never_a_row_number():
    e = R("Summ", "insert_rows", index=1, count=1)
    assert tf("=Q１!A2", e) == "=Q１!A2"   # sheet Q１, not cell Q1 of Summ
    assert tf("=Q１+A1", e) == "=Q１+A2"     # a name
    assert refs.transpose_ref("A１", e, "Summ") == "A１"


def test_find_references_lists_only_what_an_edit_here_moves():
    found = [h.group(0) for h in refs.find_references(
        "=[1]Data!A1:A3+Data!A3+#REF!A1+Sales[Q1]+'North \"Q\"'!B2+αA1+\"C3\"")]
    assert found == ["Data!A3", "'North \"Q\"'!B2"]
