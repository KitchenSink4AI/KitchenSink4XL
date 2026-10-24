"""A union of references is priced as Excel compiles it.

Excel compiles a parenthesized union of references, SUM((A1,B2)), inside
one memory token around the whole group: 7 bytes when the group holds only
cell and range references on the formula's own sheet, 3 when it holds
something Excel evaluates (a function call, a reference to another sheet).
The compiled-size floor counted the references and the union operators but
not that token, so =SUM((A1,...,A1)) with 2,047 references saved "ok" and
Excel would not open the workbook (2,046 is Excel's largest).

MEASURED in a hidden Excel 16 on a private desktop: for each call below,
100 calls of it plus n references to A1, with Excel's largest n found by
bisection (Excel opens n, not n + 1), and the direct shape bisected on its
reference count. The product now accepts exactly n and refuses n + 1.
"""
from __future__ import annotations

import pytest

from xlsx_mcp.core import calc, limits

#: (the call as typed, calls per formula, Excel's largest n of references)
MEASURED = [
    ("SUM((A1))", 100, 1885),                       # no union: nothing more
    ("ABS((A1))", 100, 1898),
    ("SUM((A1,A1))", 100, 1698),                    # one group: 7
    ("SUM((A1,A1,A1))", 100, 1598),                 # still one token
    ("SUM((A1,A1,A1,A1,A1))", 100, 1398),
    ("SUM((A1:B2,A1:B2))", 100, 1548),
    ("SUM((A1,A1),(A1,A1))", 100, 1410),            # two groups: 14
    ("SUM(((A1,A1),A1))", 100, 1585),               # a group inside a group: one token
    ("COUNT((A1,A1))", 100, 1698),
    ("INDEX((A1:B2,A1:B2),1,1,2)", 100, 1435),
    ("(A1,A1)", 100, 1748),                         # a group on its own
    ("SUM((A1,A1)+1)", 100, 1648),
    ("SUM((Sheet!A1,Sheet!A1))", 100, 1698),        # another sheet: 3
    ("SUM((A1,INDEX(A1:B2,1,1)))", 100, 1548),      # a function inside: 3
    ("SUM((INDEX(A1:B2,1,1),INDEX(A1:B2,1,1)))", 100, 1348),
]


def _formula(call, k, n):
    return "=" + "+".join([call] * k + ["A1"] * n)


def _accepted(typed):
    """What the server runs: the typed formula (set_formula) and the stored
    text (the save gate)."""
    try:
        limits.check_formula(typed)
        stored = calc.normalize_formula(typed)[0]
        limits.check_formula(stored if stored.startswith("=") else "=" + stored)
        return True
    except limits.ExcelWouldRefuse:
        return False


@pytest.mark.parametrize("call,k,n", MEASURED, ids=[f"{c}x{k}" for c, k, _ in MEASURED])
def test_the_boundary_is_excels_exactly(call, k, n):
    assert _accepted(_formula(call, k, n)), "refused a formula Excel opens"
    assert not _accepted(_formula(call, k, n + 1)), "accepted a formula Excel will not open"


def test_one_long_union_stops_where_excel_does():
    """The boundary shape: 2,046 references in one group open, 2,047 do not."""
    def direct(n):
        return "=SUM((" + ",".join(["A1"] * n) + "))"
    assert _accepted(direct(2046))
    assert not _accepted(direct(2047))


def test_the_union_token_is_counted_once_per_group():
    plain = limits.compiled_size_floor("=SUM((A1))")
    assert limits.compiled_size_floor("=SUM((A1,A1))") == plain + 7 + 1 + limits.UNION_GROUP_BYTES
    assert limits.compiled_size_floor("=SUM(((A1,A1),A1))") == \
        limits.compiled_size_floor("=SUM(((A1,A1),A1))".replace("((A1,A1),A1)", "(A1,A1,A1)")) + 1
    assert limits.compiled_size_floor("=SUM((Sheet!A1,Sheet!A1))") == \
        4 + 9 + 1 + 9 + 1 + limits.UNION_GROUP_EVALUATED_BYTES


def test_what_was_not_measured_adds_nothing():
    """A group holding something not measured (a number, a name, a space)
    adds no union token, so the floor never refuses what Excel opens."""
    assert limits.compiled_size_floor("=SUM((A1,1))") == 4 + 7 + 1 + 3 + 1
    assert limits.compiled_size_floor("=SUM((q,A1))") == 4 + 5 + 1 + 7 + 1
    assert limits.compiled_size_floor("=SUM((A1, A1))") == 4 + 7 + 1 + 4 + 7 + 1
    # A union inside a function call inside another group is left to the
    # outer group's token.
    inner = limits.compiled_size_floor("=SUM((A1,INDEX((A1,A1),1,1,1)))")
    assert inner == limits.compiled_size_floor("=SUM((A1,INDEX((A1),1,1,1)))") + 1 + 7
