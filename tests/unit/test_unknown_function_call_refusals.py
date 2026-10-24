"""Calls to functions Excel does not
have built in.

The compiled-size floor counted every function call's opening as 3 bytes. A
call to a LAMBDA defined name, a user-defined or add-in function, or any name
Excel does not know compiles to a 5-byte name token plus a 4-byte call token,
so a formula of 1,300 such calls (6,499 characters) saved "ok" and Excel would
not open the workbook. Re-measured in a hidden Excel 16 on a private desktop,
both on entry and on files this server wrote: q(1) 1,260 open / 1,261 not;
f() 1,638 / 1,639;
f(1) 1,260 / 1,261. Calls through a sheet qualifier and calls to a LET
variable or LAMBDA parameter cost more than the 9 counted here and are priced
in test_bound_call_size_figures.py; the 18, 22 and 24 bytes a term first reported
for sheet-scoped and qualified names there were the 8,192-character limit, not
the compiled size.
"""
from __future__ import annotations

import openpyxl
import pytest

from xlsx_mcp.core import limits
from xlsx_mcp.core.errors import XlMcpError


def _f(term, n):
    return "=" + "+".join([term] * n)


@pytest.mark.parametrize("term,last_open", [("q(1)", 1260), ("f()", 1638), ("f(1)", 1260)])
def test_the_measured_boundary_exactly(term, last_open):
    limits.check_formula(_f(term, last_open))
    with pytest.raises(XlMcpError):
        limits.check_formula(_f(term, last_open + 1))


@pytest.mark.parametrize("formula", [_f("q(1)", 1300), _f("q(1)", 1638), _f("f(1)", 1638),
                                     _f("f()", 2048), _f("u(A1)", 1000)],
                         ids=["q1_1300", "q1_1638", "f1_1638", "f0_2048", "udf_ref_1000"])
def test_the_reported_shapes_are_refused(formula):
    with pytest.raises(XlMcpError) as ei:
        limits.check_formula(formula)
    assert "too large for Excel to store" in str(ei.value)


@pytest.mark.parametrize("formula", [
    _f("k", 2730),                                   # names (control)
    _f("T[[a]:[b]]", 744),                           # two-column table references
    _f("ABS(1)", 1170),                              # built-in, fixed arguments
    _f("SUM(1)", 1170),                              # built-in, variable arguments
    _f("NOW()", 1365),                               # built-in, no arguments
    _f("EDATE(1,1)", 744),                           # once an add-in, now built in
    _f("_xlfn.XLOOKUP(1,A1:A2,B1:B2)", 280),         # newer function, prefixed
    _f("XLOOKUP(1,A1:A2,B1:B2)", 282),               # the same, as typed (8,177 stored)
    "=LET(g,LAMBDA(x,x+1)," + "+".join(["g(1)"] * 740) + ")",  # the largest that opens
], ids=["names", "table2col", "abs", "sum", "now", "edate", "xlookup_prefixed",
        "xlookup_typed", "let_bound_lambda"])
def test_what_excel_opens_stays_accepted(formula):
    limits.check_formula(formula)


def test_a_call_is_a_name_call_only_when_excel_has_no_such_function():
    names = {"q(": True, "f(": True, "MyUdf(": True, "_xludf.F(": True, "Sheet1!q(": True,
             "SUM(": False, "sum(": False, "STDEV.S(": False, "_xlfn.XLOOKUP(": False,
             "XLOOKUP(": False, "_xlfn._xlws.FILTER(": False, "_xlpm.g(": False,
             "NETWORKDAYS(": False}
    assert {k: limits._calls_a_name(k) for k in names} == names


def test_the_file_tier_refuses_before_writing(tmp_path):
    from xlsx_mcp.ops import cells, formulas
    p = tmp_path / "n8.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.active["A1"] = 1
    wb.save(p)
    before = p.read_bytes()
    with pytest.raises(XlMcpError):
        cells.set_cell(str(p), {"cell": "F1", "sheet": "S"}, _f("q(1)", 1300))
    with pytest.raises(XlMcpError):
        formulas.set_formula(str(p), {"cell": "F2", "sheet": "S"}, _f("f()", 2048))
    assert p.read_bytes() == before
