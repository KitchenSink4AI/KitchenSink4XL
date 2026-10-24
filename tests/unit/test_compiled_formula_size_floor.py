"""The compiled-size floor counts IF and CHOOSE jump tokens,
whitespace and intersections, and does not count array-constant values.

Formulas under 8,192 characters that Excel will not open (600 IF terms, ten
200-option CHOOSEs, 2,600 spaced terms) used to pass and save "verified";
the whole workbook then failed to open. A SUM over 1,900 array constants,
which Excel opens, was refused. Every boundary below was measured in a
hidden Excel: the "opens" side opened, the "refused" side did not.
"""
from __future__ import annotations

import pytest

from xlsx_mcp.core.errors import XlMcpError
from xlsx_mcp.core.limits import MAX_FORMULA_COMPILED_BYTES, check_formula, compiled_size_floor


def _if(n):
    return "=" + "+".join(["IF(A1,1,2)"] * n)


EXCEL_REFUSES = {
    "if_terms_600": _if(600),
    "choose_10x200": "=" + "+".join(["CHOOSE(1," + ",".join(["1"] * 200) + ")"] * 10),
    "spaces_2600": "=" + " +".join(["1"] * 2600),
    "isect_area_469": "=" + "+".join(["A1:B2 B1:C3"] * 469),
    "isect_cell_713": "=" + "+".join(["A1 A1"] * 713),
}
EXCEL_OPENS = {
    "array_1900": "=SUM({" + ",".join(["1.5"] * 1900) + "})",
    "if_terms_400": _if(400),
    "isect_area_468": "=" + "+".join(["A1:B2 B1:C3"] * 468),
    "isect_cell_712": "=" + "+".join(["A1 A1"] * 712),
    "array_then_if_545": "=SUM({" + ",".join(["1"] * 1000) + "})+" + "+".join(["IF(A1,1,2)"] * 545),
}


@pytest.mark.parametrize("name", sorted(EXCEL_REFUSES))
def test_refused_when_excel_would_not_open_it(name):
    with pytest.raises(XlMcpError, match="too large for Excel to store"):
        check_formula(EXCEL_REFUSES[name])


@pytest.mark.parametrize("name", sorted(EXCEL_OPENS))
def test_accepted_when_excel_opens_it(name):
    check_formula(EXCEL_OPENS[name])


def test_exact_at_the_limit_for_if_choose_and_whitespace():
    """At a floor of exactly 16,383 each form opened in Excel and one more
    "+1" (4 bytes) did not, so the counts are exact, not just a bound."""
    for base in (_if(500)[1:], "+".join(["CHOOSE(1," + ",".join(["1"] * 200) + ")"] * 8),
                 " +".join(["1"] * 1900)):
        body = base + "+1" * ((MAX_FORMULA_COMPILED_BYTES - compiled_size_floor(base)) // 4)
        assert compiled_size_floor(body) == MAX_FORMULA_COMPILED_BYTES
        check_formula("=" + body)
        with pytest.raises(XlMcpError):
            check_formula("=" + body + "+1")


def test_two_argument_if_and_nested_choose_are_counted():
    assert compiled_size_floor("=IF(A1,1)") == 3 + 1 + 7 + 3 + 4 + 4
    # CHOOSE(1,2,3): call 3+1, three numbers 9, jump table 4+2*3, two jumps 8
    assert compiled_size_floor("=CHOOSE(1,2,3)") == 3 + 1 + 9 + 10 + 8
