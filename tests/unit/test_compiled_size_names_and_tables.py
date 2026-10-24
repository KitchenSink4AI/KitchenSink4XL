"""The compiled-size check counts a
defined-name reference and a structured reference at their real size.

compiled_size_floor counted any operand that is neither a cell nor an area
as one byte, so =q+q+... with 4,000 name references (7,999 characters) and
=T[a]+T[a]+... with 1,630 structured references saved "ok" and Excel would
not open the workbook. MEASURED in a hidden Excel 16 (bisection on
files opened read-only): 2,730 name references open and 2,731 do not, so a
name reference is 5 bytes (6n - 1 with the + between them, against the
16,383-byte limit); 1,092 single-column structured references open and 1,093
do not, so one is 14 bytes. A two-column reference (T[[a]:[b]]) measures
larger still (744 open, 745 do not); counting it at 14 keeps the floor a
floor (it can miss, never refuse wrongly).
"""
from __future__ import annotations

import pytest

from xlsx_mcp.core import limits
from xlsx_mcp.core.errors import XlMcpError


def _names(n):
    return "=" + "+".join(["q"] * n)


def _table(n, ref="T[a]"):
    return "=" + "+".join([ref] * n)


def test_name_and_structured_reference_sizes():
    assert limits.compiled_size_floor(_names(10)) == 6 * 10 - 1
    assert limits.compiled_size_floor(_table(10)) == 15 * 10 - 1
    assert limits.compiled_size_floor(_table(10, "T[[#This Row],[a]]")) == 15 * 10 - 1
    assert limits.compiled_size_floor("=Sheet1!q+1") >= 5 + 1 + 3


@pytest.mark.parametrize("formula", [_names(2731), _table(1093), _names(4000), _table(1630)],
                         ids=["names_2731", "table_1093", "names_4000", "table_1630"])
def test_refused_where_excel_would_not_open_the_workbook(formula):
    with pytest.raises(XlMcpError) as ei:
        limits.check_formula(formula)
    assert "16,383" in str(ei.value)


@pytest.mark.parametrize("formula", [_names(2730), _table(1092), _names(2700), _table(1000),
                                     _table(744, "T[[a]:[b]]")],
                         ids=["names_2730", "table_1092", "names_2700", "table_1000", "two_cols_744"])
def test_accepted_where_excel_opens_it(formula):
    limits.check_formula(formula)
