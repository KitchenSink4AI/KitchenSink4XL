"""The formula size count against Excel for LET, array constants and seven
more variable-argument functions.

MEASURED in a hidden Excel 16 on a private desktop: k copies of the shape
plus n references to A1, written as the text the product stores; Excel's
largest n by bisection (it opens n, not n + 1, and the formula comes back
intact). An earlier count refused 12 reference counts Excel opens for
LET(x,A1,x) (it priced the LET call at 3 bytes and each name at 5;
measured: the call takes 5, a declared name none, each use 6, plus a 4-byte
flag once), and accepted formulas Excel will not open for IMSUM, TOCOL,
WRAPROWS and array constants (a constant takes 15 bytes whatever it holds,
not 1). Every shape here is accepted at Excel's n and refused at n + 1. The
LET and LAMBDA shapes measured to the byte are in
test_let_and_lambda_sizes_to_the_byte.py. No Excel here.
"""
from __future__ import annotations

import pytest

from xlsx_mcp.core import calc, limits

#: (the shape as typed, copies per formula, Excel's largest n of references)
EXACT = [
    ("LET(x,A1,x)", 50, 1928),
    ("LET(x,1,x)", 50, 1953),
    ("LET(x,A1,1)", 50, 1947),
    ("LET(x,A1,y,A1,x+y)", 25, 1944),
    ("LET(x,A1,x+x+x)", 25, 1944),
    ("LET( x , A1 , x )", 25, 1950),
    ("LET(x,A1,x)+NOW()", 25, 1975),
    ("LET(x,{1,2,3},SUM(x))", 25, 1950),
    ('IMSUM("1","2")', 100, 1860),
    ('IMPRODUCT("1","2")', 50, 1954),
    ("TOCOL(A1:B2)", 50, 1935),
    ("TOROW(A1:B2)", 50, 1935),
    ("WRAPROWS(A1:B2,1)", 50, 1916),
    ("WRAPCOLS(A1:B2,1)", 50, 1916),
    ("EXPAND(A1:B2,2)", 50, 1916),
    ("SUM({1})", 100, 1798),
    ("SUM({1,2,3})", 100, 1798),
    ("SUM({1,2;3,4})", 100, 1798),
    ('SUM({"a","b"})', 100, 1798),
    ("SUM({1,2,3,4,5,6,7,8,9,10})", 50, 1923),
    ("{1,2,3}", 100, 1848),
    ("ABS(A1)", 100, 1910),
    # Counted below Excel until the LAMBDA call, the inline call and LET's
    # flag were measured to the byte; exact since.
    ("LET(x,A1:B2,SUM(x))", 25, 1956),
    ("LET(x,A1,y,x+1,y)", 25, 1956),
    ("LET(f,LAMBDA(a,a+1),f(A1))", 25, 1875),
    ("LAMBDA(a,a+1)(A1)", 25, 1910),
    ("LAMBDA(a,b,a+b)(A1,A1)", 25, 1879),
]


def _formula(call, k, n):
    return "=" + "+".join([call] * k + ["A1"] * n)


def _accepted(typed):
    try:
        limits.check_formula(typed)
        stored = calc.normalize_formula(typed)[0]
        limits.check_formula(stored if stored.startswith("=") else "=" + stored)
        return True
    except limits.ExcelWouldRefuse:
        return False


@pytest.mark.parametrize("call,k,n", EXACT, ids=[c for c, _, _ in EXACT])
def test_excels_boundary_is_kept(call, k, n):
    assert _accepted(_formula(call, k, n)), "refused a formula Excel opens"
    assert not _accepted(_formula(call, k, n + 1)), "accepted a formula Excel will not open"


def test_a_let_declaration_is_told_from_a_use():
    # LET(x,A1,x): the call 5, the declared name 0, A1 7, the use 6, and the
    # 4-byte LET flag once.
    assert limits.compiled_size_floor("=LET(x,A1,x)") == 5 + 0 + 7 + 6 + 4
    # The last argument is the calculation, never a declaration, even when bare.
    assert limits.compiled_size_floor("=LET(x,1,y,2,x)") == 5 + 0 + 3 + 0 + 3 + 6 + 4
