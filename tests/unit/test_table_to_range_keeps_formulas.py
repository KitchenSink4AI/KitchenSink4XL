"""manage_table to_range keeps every formula that names the table working.

Excel's Convert to Range rewrites each structured reference to the table as
the plain cells it stood for before the table goes, so every value stays.
The expected text below is what Excel's own Convert to Range (ListObject.
Unlist, a private hidden Excel) wrote for the same shapes on made-up tables.
A reference that cannot be written as the same cells refuses the conversion
before anything changes, naming where it stands. No Excel here; every
workbook is made up and built by code.
"""
from __future__ import annotations

import hashlib

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableColumn

from xlsx_mcp.ops import tables

ROWS = [["Region", "Q1", "Q2", "Q3 #", "Total"], ["North", 10, 20, 1, None], ["South", 30, 40, 2, None],
        ["East", 50, 60, 3, None]]
CALC = "Sales[[#This Row],[Q1]]+Sales[[#This Row],[Q2]]"

# cell -> (formula written, Excel's text after Convert to Range)
PLAIN = {
    "Data!H2": ("=SUM(Sales[Q1])", "=SUM(Data!$B$2:$B$4)"),
    "Data!H3": ("=SUM(Sales[[Q1]:[Q2]])", "=SUM(Data!$B$2:$C$4)"),
    "Data!H4": ("=ROWS(Sales[#All])", "=ROWS(Data!$A$1:$E$4)"),
    "Data!H5": ("=COUNTA(Sales[#Headers])", "=COUNTA(Data!$A$1:$E$1)"),
    "Data!H6": ("=Sales[[#Headers],[Q2]]", "=Data!$C$1"),
    "Data!H7": ("=SUM(Sales[])", "=SUM(Data!$A$2:$E$4)"),
    "Data!H14": ("=SUM(Sales)", "=SUM(Sales)"),
    "Data!H8": ("=SUM(Sales[#Data])", "=SUM(Data!$A$2:$E$4)"),
    "Data!H9": ("=COUNTA(Sales[[#Headers],[#Data],[Q1]])", "=COUNTA(Data!$B$1:$B$4)"),
    "Data!H10": ('="Sales[Q1]"', '="Sales[Q1]"'),
    "Data!H11": ("=SUM(Sales[Q3 '#])", "=SUM(Data!$D$2:$D$4)"),
    "Data!H12": ('=SUMIFS(Sales[Q2],Sales[Region],"North")', '=SUMIFS(Data!$C$2:$C$4,Data!$A$2:$A$4,"North")'),
    "Data!H13": ("=SUM(Costs[Amt])+SUM(Sales[Q2])", "=SUM(Costs[Amt])+SUM(Data!$C$2:$C$4)"),
    "Data!G3": ("=SUM(Sales[#This Row])", "=SUM(Data!$A3:$E3)"),
    "Data!E2": ("=" + CALC, "=Data!$B2+Data!$C2"),
    "Data!E4": ("=" + CALC, "=Data!$B4+Data!$C4"),
    "Summary!B1": ("=SUM(Sales[Q2])", "=SUM(Data!$C$2:$C$4)"),
    "Summary!B2": ("=INDEX(Sales[Q2],2)", "=INDEX(Data!$C$2:$C$4,2)"),
    "Summary!B5": ("=SUM(Costs[Amt])", "=SUM(Costs[Amt])"),
}
TOTALS = {
    "Summary!B3": ("=SUM(Sales[#Totals])", "=SUM(Data!$A$5:$E$5)"),
    "Summary!B4": ("=ROWS(Sales[#All])", "=ROWS(Data!$A$1:$E$5)"),
    "Summary!B7": ("=SUM(Sales[[#Data],[#Totals],[Q1]])", "=SUM(Data!$B$2:$B$5)"),
    "Summary!B8": ("=Sales[[#Totals],[Q1]]", "=Data!$B$5"),
    "Data!H2": ("=SUM(Sales[Q1])", "=SUM(Data!$B$2:$B$4)"),
}


def _book(p, cells, *, totals=False, sheet="Data", extra=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for r, row in enumerate(ROWS, 1):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)
    for r in (2, 3, 4):
        ws.cell(r, 5, "=" + CALC)
    last = 4
    if totals:
        ws["A5"] = "Total"
        ws["B5"] = "=SUBTOTAL(109,Sales[Q1])"
        last = 5
    t = Table(displayName="Sales", ref=f"A1:E{last}", totalsRowCount=1 if totals else None)
    t.tableColumns = [TableColumn(id=i + 1, name=h) for i, h in enumerate(ROWS[0])]
    if totals:
        t.tableColumns[0].totalsRowLabel = "Total"
        t.tableColumns[1].totalsRowFunction = "sum"
    ws.add_table(t)
    other = wb.create_sheet("Other Sheet")
    for r, row in enumerate([["Item", "Amt"], ["Pens", 5], ["Ink", 7]], 1):
        for c, v in enumerate(row, 1):
            other.cell(r, c, v)
    other.add_table(Table(displayName="Costs", ref="A1:B3"))
    summ = wb.create_sheet("Summary")
    wb.defined_names["SalesQ1"] = DefinedName("SalesQ1", attr_text="Sales[Q1]")
    for where, (written, _excel) in cells.items():
        sh, a = where.split("!")
        wb[sh.replace("Data", sheet) if sh == "Data" else sh][a] = written.replace("Data!", f"{sheet}!") \
            if False else written
    if extra:
        extra(wb)
    wb.save(p)


def _formulas(p):
    wb = openpyxl.load_workbook(p)
    return wb, {f"{ws.title}!{c.coordinate}": c.value for ws in wb.worksheets for row in ws.iter_rows()
                for c in row if isinstance(c.value, str) and c.value.startswith("=")}


@pytest.mark.parametrize("totals,cells", [(False, PLAIN), (True, TOTALS)], ids=["plain", "totals_row"])
def test_to_range_writes_each_reference_as_excels_convert_to_range_does(tmp_path, totals, cells):
    p = tmp_path / "book.xlsx"
    _book(p, cells, totals=totals)
    out = tables.manage_table(str(p), "Sales", "to_range", verify_com=False)
    assert out["ok"]
    wb, got = _formulas(p)
    assert not wb["Data"].tables
    for where, (_written, excel) in cells.items():
        assert got[where] == excel, where
    assert wb.defined_names["SalesQ1"].attr_text == "Data!$B$2:$B$4"
    left = [f"{k} {v}" for k, v in got.items() if "Sales[" in v and not v.startswith('="')]
    assert left == []


def test_a_sheet_name_that_needs_quotes_is_quoted(tmp_path):
    p = tmp_path / "book.xlsx"
    _book(p, {"Summary!B1": ("=SUM(Sales[Q2])", None)}, sheet="Q Data")
    tables.manage_table(str(p), "Sales", "to_range", verify_com=False)
    _wb, got = _formulas(p)
    assert got["Summary!B1"] == "=SUM('Q Data'!$C$2:$C$4)"
    assert got["Q Data!E3"] == "='Q Data'!$B3+'Q Data'!$C3"


def test_a_this_row_reference_outside_the_table_refuses_and_changes_nothing(tmp_path):
    p = tmp_path / "book.xlsx"
    _book(p, {"Summary!B6": ("=Sales[[#This Row],[Q1]]", None), "Summary!B1": ("=SUM(Sales[Q2])", None)})
    before = hashlib.sha256(p.read_bytes()).hexdigest()
    with pytest.raises(Exception) as ei:
        tables.manage_table(str(p), "Sales", "to_range", verify_com=False)
    assert "Summary!B6" in str(ei.value)
    assert "Nothing was changed" in str(ei.value)
    assert hashlib.sha256(p.read_bytes()).hexdigest() == before


def test_a_pivot_or_rule_naming_the_table_refuses(tmp_path):
    p = tmp_path / "book.xlsx"

    def rule(wb):
        from openpyxl.formatting.rule import FormulaRule
        wb["Data"].conditional_formatting.add("B2:B4", FormulaRule(formula=["B2>AVERAGE(Sales[Q1])"]))
    _book(p, {}, extra=rule)
    before = hashlib.sha256(p.read_bytes()).hexdigest()
    with pytest.raises(Exception) as ei:
        tables.manage_table(str(p), "Sales", "to_range", verify_com=False)
    assert "conditional format" in str(ei.value)
    assert hashlib.sha256(p.read_bytes()).hexdigest() == before


def test_a_dropdown_reading_the_table_through_indirect_text_is_left_as_excel_leaves_it(tmp_path):
    p = tmp_path / "book.xlsx"

    def dv(wb):
        from openpyxl.worksheet.datavalidation import DataValidation
        d = DataValidation(type="list", formula1='INDIRECT("Sales[Region]")')
        d.add("C1")
        wb["Summary"].add_data_validation(d)
    _book(p, {"Summary!B1": ("=SUM(Sales[Q2])", None)}, extra=dv)
    tables.manage_table(str(p), "Sales", "to_range", verify_com=False)
    wb, got = _formulas(p)
    assert got["Summary!B1"] == "=SUM(Data!$C$2:$C$4)"
    assert [d.formula1 for d in wb["Summary"].data_validations.dataValidation] == ['INDIRECT("Sales[Region]")']
