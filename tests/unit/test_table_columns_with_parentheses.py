"""A formula naming a table column such as "Unit price ($)" is written.

The write path refuses whitespace between a function name and its opening
parenthesis, because Excel's file parser cannot open SUM (1,2). It read the
inside of a structured reference as formula text, so the column name
"Unit price ($)" looked like a call to a function named price, and
=SUM(Orders[Unit price ($)]) was refused, though Excel stores and opens it
(and the product's own read of such a cell could not be written back).
The brackets of a structured reference are now read as Excel reads them:
a column name, whatever characters it holds. All data here is made up.
"""

from __future__ import annotations

import openpyxl
import pytest
from openpyxl.worksheet.table import Table

from xlsx_mcp.core.errors import FormulaRejected
from xlsx_mcp.ops import cells, formulas


def _book(tmp_path):
    p = tmp_path / "orders.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Item", "Unit price ($)", "Qty", "Rate (%)", "Cost (net) [eur]"])
    for row in [("Widget", 2.5, 4, 5, 9), ("Gadget", 4.0, 3, 7, 11)]:
        ws.append(list(row))
    t = Table(displayName="Orders", ref="A1:E3")
    ws.add_table(t)
    wb.save(p)
    return str(p)


@pytest.mark.parametrize("formula", [
    "=SUM(Orders[Unit price ($)])",
    "=Orders[[#This Row],[Unit price ($)]]*2",
    "=SUM(Orders[[Unit price ($)]:[Qty]])",
    "=AVERAGE(Orders[Rate (%)])+MAX(Orders[Unit price ($)])",
    "=SUM(Orders[Cost (net) '[eur']])",
    "=LET(x,SUM(Orders[Unit price ($)]),x*2)",
])
def test_a_column_name_with_a_space_and_parentheses_is_written(tmp_path,
                                                               formula):
    p = _book(tmp_path)
    out = formulas.set_formula(p, {"cell": "G2"}, formula, sheet="Data")
    assert out["verified"] is True
    got = cells.get_cells(p, ["G2"], values="formula", sheet="Data")
    stored = got["cells"][0][1]
    assert "Unit price ($)" in stored or "Rate (%)" in stored \
        or "Cost (net)" in stored


@pytest.mark.parametrize("formula", [
    "=SUM (1,2)",
    "=SUM(Orders[Qty])+MAX (1,2)",
    "=SUM(Orders[Unit price ($)]) + ROUND\t(1.5,0)",
])
def test_a_real_space_before_a_parenthesis_is_still_refused(tmp_path,
                                                            formula):
    p = _book(tmp_path)
    with pytest.raises(FormulaRejected) as ei:
        formulas.set_formula(p, {"cell": "G2"}, formula, sheet="Data")
    assert "parenthesis" in str(ei.value)


@pytest.mark.parametrize("formula,stored", [
    # a LET name that is also a column name: the column keeps its name
    ("=LET(Qty,2,SUM(Orders[Qty])*Qty)",
     "=_xlfn.LET(_xlpm.Qty,2,SUM(Orders[Qty])*_xlpm.Qty)"),
    # the comma inside the brackets is not an argument separator, so y is
    # still read as a declared name (left bare, Excel refuses the file)
    ("=LET(x,Orders[[#This Row],[Qty]],y,2,x+y)",
     "=_xlfn.LET(_xlpm.x,Orders[[#This Row],[Qty]],_xlpm.y,2,"
     "_xlpm.x+_xlpm.y)"),
    # a quote inside a column name is not a string
    ('=XLOOKUP(1,Orders[Size "L"],Orders[Qty])',
     '=_xlfn.XLOOKUP(1,Orders[Size "L"],Orders[Qty])'),
    # an optional LAMBDA parameter is still read as one
    ("=LAMBDA(a,[b],a+b)(1)", "=_xlfn.LAMBDA(_xlpm.a,_xlop.b,_xlpm.a+_xlpm.b)(1)"),
])
def test_a_structured_reference_is_read_as_column_names(formula, stored):
    from xlsx_mcp.core import calc
    assert calc.normalize_formula(formula)[0] == stored


def test_set_cell_takes_the_same_formula(tmp_path):
    p = _book(tmp_path)
    out = cells.set_cell(p, {"cell": "G3"},
                         "=Orders[[#This Row],[Unit price ($)]]*Orders"
                         "[[#This Row],[Qty]]", sheet="Data")
    assert out["verified"] is True
