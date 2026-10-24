"""What a call costs when its callee is a sheet-qualified
name or a LET variable / LAMBDA parameter.

Measured in a hidden Excel 16 on a private desktop. Every boundary below is on
the FILE side (a workbook this server's own writer produced, opened read-only
in a fresh Excel): the "opens" count opened and the next one did not.

- A call to a name qualified by a sheet (S!q(1), whether q has workbook or
  sheet scope) compiles to a 7-byte qualified-name token plus the 4-byte call:
  11 bytes before its arguments, where an unqualified name costs 9. 1,092 of
  =S!q(1)+... open and 1,093 do not. The floor counted 9, so with a short
  sheet name 1,093 to 1,260 such calls saved and Excel would not open the
  workbook.
- A call to a LET variable or LAMBDA parameter holding a lambda compiles to 10
  bytes. In the file such a name carries the _xlpm. prefix, which the floor
  priced as a built-in (3). With 100 such calls and n references to A1:
  =LET(g,LAMBDA(x,x+1),g(1)+..., A1+...) opens up to n = 1,867 and the same
  with LAMBDA(f, f(1)+..., A1+...)(q) up to n = 1,869. Excel's entry limit
  for these shapes is NOT the file limit (the entry accepted n = 1,999), which
  is why the figures here are file-side.
- A sheet-scoped name costs what a workbook-scoped one does: a one-letter
  sheet-scoped LAMBDA name measures 1,260 / 1,261 as a call (13 bytes a term)
  and 2,730 / 2,731 as an operand (6), exactly the workbook-scoped figures.
  An earlier version reported 18 bytes a term (910) for a sheet-scoped name, 22 (744) for
  S!lam_s and 24 (682) for Sheet2!q: those were the 8,192-CHARACTER limit
  meeting terms of 9, 11 and 12 characters, never the compiled size.

LET's and LAMBDA's own tokens are still priced as built-ins, so for the two
combined shapes the floor accepts up to two references more than Excel does.
"""
from __future__ import annotations

import openpyxl
import pytest

from xlsx_mcp.core import calc, limits
from xlsx_mcp.core.errors import XlMcpError


def _plus(term, n):
    return "+".join([term] * n)


def _f(term, n):
    return "=" + _plus(term, n)


def _let(calls, refs):
    return "=LET(g,LAMBDA(x,x+1)," + _plus("g(1)", calls) + "+" + _plus("A1", refs) + ")"


def _lam(calls, refs):
    return "=LAMBDA(f," + _plus("f(1)", calls) + "+" + _plus("A1", refs) + ")(q)"


def _stored(formula):
    """The text the file holds: modern functions and LET/LAMBDA names prefixed."""
    return calc.normalize_formula(formula)[0]


def _refused(formula, match="too large for Excel to store"):
    with pytest.raises(XlMcpError, match=match):
        limits.check_formula(formula)


# ---------------------------------------------------------------- qualified


@pytest.mark.parametrize("term", ["S!q(1)", "S!s(1)"])
def test_a_sheet_qualified_call_costs_eleven_exactly(term):
    limits.check_formula(_f(term, 1092))
    _refused(_f(term, 1093))


def test_a_qualified_call_is_priced_whatever_qualifies_it():
    base = limits.compiled_size_floor("=q(1)")
    for name in ("S!q(1)", "'My Sheet'!q(1)", "Sheet2!q(1)", "[Book.xlsx]Sheet1!q(1)"):
        assert limits.compiled_size_floor("=" + name) - base == 2, name


# ------------------------------------------------------- sheet scope = same


def test_a_sheet_scoped_name_costs_what_a_workbook_name_does():
    """Measured with one-letter names so the character limit stays out of it.
    The floor takes no scope because no cost depends on one."""
    limits.check_formula(_f("s(1)", 1260))   # a sheet-scoped LAMBDA name, as a call
    _refused(_f("s(1)", 1261))
    limits.check_formula(_f("j", 2730))      # a sheet-scoped name, as an operand
    _refused(_f("j", 2731))


@pytest.mark.parametrize("term,last_open", [("lam_s(1)", 910), ("S!lam_s(1)", 744), ("Sheet2!q(1)", 682)])
def test_the_earlier_figures_for_scoped_and_qualified_names_are_the_character_limit(term, last_open):
    limits.check_formula(_f(term, last_open))
    _refused(_f(term, last_open + 1), match="characters long")


# ----------------------------------------------- LET variable / LAMBDA param


@pytest.mark.parametrize("form", [lambda f: f, _stored], ids=["typed", "stored"])
@pytest.mark.parametrize("build,last_open,first_refused",
                         [(_let, 1867, 1870), (_lam, 1869, 1872)], ids=["let", "lambda"])
def test_a_call_to_a_bound_lambda_costs_ten(form, build, last_open, first_refused):
    limits.check_formula(form(build(100, last_open)))
    _refused(form(build(100, first_refused)))


@pytest.mark.parametrize("build", [_let, _lam], ids=["let", "lambda"])
@pytest.mark.parametrize("refs", [0, 50, 700, 1867])
def test_the_typed_formula_and_the_text_the_file_holds_are_priced_alike(build, refs):
    typed = build(100, refs) if refs else build(100, 1)
    assert limits.compiled_size_floor(typed) == limits.compiled_size_floor(_stored(typed))


def _extra(formula, same_but_unbound):
    """Bytes the first formula costs over one whose call goes to an unknown name."""
    return limits.compiled_size_floor(formula) - limits.compiled_size_floor(same_but_unbound)


def test_a_call_is_bound_only_inside_the_let_or_lambda_that_declares_it():
    assert _extra("=LET(g,LAMBDA(x,x),g(1))", "=LET(g,LAMBDA(x,x),h(1))") == 1
    assert _extra("=LAMBDA(f,f(1))(q)", "=LAMBDA(f,h(1))(q)") == 1
    assert _extra("=LET(G,LAMBDA(x,x),g(1))", "=LET(G,LAMBDA(x,x),h(1))") == 1           # case
    assert _extra("=LET(g,LAMBDA(x,x),LET(u,1,g(1)))", "=LET(g,LAMBDA(x,x),LET(u,1,h(1)))") == 1  # nested
    assert _extra("=LET(g,LAMBDA(x,x),1)+g(1)", "=LET(g,LAMBDA(x,x),1)+h(1)") == 0       # after the LET
    assert _extra("=LET(g,LAMBDA(x,x),g(1))+q(1)", "=LET(g,LAMBDA(x,x),g(1))+h(1)") == 0  # another name
    assert _extra("=LET(SUM,LAMBDA(x,x),SUM(1))", "=LET(SUM,LAMBDA(x,x),MAX(1))") == 0   # a built-in wins


def test_a_workbook_lambda_keeps_its_nine_inside_a_let_that_does_not_declare_it():
    assert _extra("=LET(a,1,q(a))", "=LET(a,1,h(a))") == 0


# -------------------------------------------------------- through the product


def _book(tmp_path):
    p = tmp_path / "n8.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.active["A1"] = 1
    wb.save(p)
    return p


def test_the_file_tier_refuses_a_bound_call_formula_excel_would_not_open(tmp_path):
    from xlsx_mcp.ops import cells
    p = _book(tmp_path)
    before = p.read_bytes()
    with pytest.raises(XlMcpError, match="too large for Excel to store"):
        cells.set_cell(str(p), {"cell": "F1", "sheet": "S"}, _let(100, 1875))
    assert p.read_bytes() == before


def test_the_file_tier_still_writes_what_excel_opens(tmp_path):
    from xlsx_mcp.ops import cells
    p = _book(tmp_path)
    cells.set_cell(str(p), {"cell": "F1", "sheet": "S"}, _let(100, 1867))
    cells.set_cell(str(p), {"cell": "F2", "sheet": "S"}, _f("S!q(1)", 1092))


def test_the_save_gate_prices_the_stored_text_too(tmp_path):
    """A formula that reaches the workbook object without set_formula is
    judged on the text the file will hold when the edit is saved."""
    from xlsx_mcp.core.package import WorkbookPackage
    p = _book(tmp_path)
    before = p.read_bytes()
    pkg = WorkbookPackage.open(str(p))
    pkg.workbook["S"]["F1"] = _stored(_let(100, 1875))
    with pytest.raises(XlMcpError, match="too large for Excel to store"):
        pkg.save()
    assert p.read_bytes() == before
