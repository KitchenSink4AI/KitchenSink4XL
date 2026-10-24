"""Each built-in function call is priced as Excel compiles it.

The compiled-size floor priced every built-in call at 3 bytes, the size of a
fixed-argument call. Excel compiles a function whose arguments can vary in
number (an optional or a repeating argument) with a 4-byte call token, and
IFERROR and IFNA with the two extra tokens a two-argument IF has. A formula
of 100 XOR calls and 1,899 to 1,910 references, or 100 IFERROR calls and
1,774 to 1,873 references, saved "ok", and Excel would not open the
workbook. Volatile functions (and SHEET, SHEETS) add one more 4-byte token
per formula.

MEASURED in a hidden Excel 16 on a private desktop: for each call below,
100 calls of it (50 where the text would pass 8,000 characters) plus n
references to A1, written as the text this server stores, with Excel's
largest n found by bisection (Excel opens n, not n + 1). The product now
accepts exactly n and refuses n + 1 for every one of them.
"""
from __future__ import annotations

import pytest

from xlsx_mcp.core import calc, limits

#: (the call as typed, calls per formula, Excel's largest n of references)
MEASURED = [
    ('ABS(A1)', 100, 1910), ('ADDRESS(1,1)', 100, 1910), ('AGGREGATE(9,6,A1:B2)', 50, 1898),
    ('AND(A1)', 100, 1898), ('ARABIC("I")', 100, 1935), ('AVEDEV(A1)', 100, 1898),
    ('AVERAGE(A1)', 100, 1898), ('AVERAGEA(A1)', 100, 1898), ('AVERAGEIF(A1:B2,1)', 100, 1785),
    ('AVERAGEIFS(A1:B2,A1:B2,1)', 100, 1623), ('BASE(1,2)', 100, 1910),
    ('CEILING(A1,1)', 100, 1873), ('CEILING.MATH(A1)', 50, 1973), ('CELL("type")', 100, 1847),
    ('CHAR(65)', 100, 1960), ('CHOOSE(1,A1)', 100, 1710), ('CHOOSECOLS(A1:B2,1)', 50, 1916),
    ('CHOOSEROWS(A1:B2,1)', 50, 1916), ('CLEAN(A1)', 100, 1910), ('CODE(A1)', 100, 1910),
    ('COLUMN(A1)', 100, 1898), ('COLUMNS(A1:B2)', 100, 1835), ('COMBIN(1,1)', 100, 1923),
    ('CONCAT(A1)', 100, 1898), ('CONCATENATE(A1)', 100, 1898), ('CONFIDENCE(1,1,1)', 100, 1885),
    ('CONVERT(1,"m","m")', 100, 1835), ('CORREL(A1:B2,A1:B2)', 100, 1673),
    ('COUNT(A1)', 100, 1898), ('COUNTA(A1)', 100, 1898), ('COUNTBLANK(A1:B2)', 100, 1835),
    ('COUNTIF(A1:B2,1)', 100, 1798), ('COUNTIFS(A1:B2,1)', 100, 1785), ('DATE(1,1,1)', 100, 1885),
    ('DATEDIF(A1,A1,"d")', 100, 1760), ('DATEVALUE("a")', 100, 1935), ('DAYS(A1,A1)', 100, 1823),
    ('DB(1,1,1,1)', 100, 1835), ('DDB(1,1,1,1)', 100, 1835), ('DEC2BIN(1)', 100, 1948),
    ('DECIMAL("1",2)', 100, 1898), ('DEVSQ(A1)', 100, 1898), ('DOLLAR(A1)', 100, 1898),
    ('DROP(A1:B2,1)', 100, 1785), ('DSUM(A1:B2,1,A1:B2)', 100, 1635), ('EDATE(A1,1)', 100, 1873),
    ('ENCODEURL("a")', 100, 1935), ('EOMONTH(A1,0)', 100, 1873), ('ERROR.TYPE(A1)', 100, 1910),
    ('EXACT(A1,A1)', 100, 1823), ('EXP(A1)', 100, 1910), ('FILTER(A1:B2,A1:B2)', 50, 1854),
    ('FIND("a",A1)', 100, 1835), ('FIXED(A1)', 100, 1898), ('FLOOR(A1,1)', 100, 1873),
    ('FLOOR.MATH(A1)', 100, 1898), ('FORECAST(1,A1:B2,A1:B2)', 100, 1635),
    ('FORMULATEXT(A1)', 100, 1910), ('FREQUENCY(A1:B2,A1:B2)', 100, 1673),
    ('FV(1,1,1)', 100, 1873), ('GCD(A1)', 100, 1898), ('GEOMEAN(A1)', 100, 1898),
    ('GETPIVOTDATA("a",A1)', 100, 1835), ('HARMEAN(A1)', 100, 1898),
    ('HLOOKUP(A1,A1:B2,2)', 100, 1698), ('HOUR(A1)', 100, 1910), ('HSTACK(A1)', 100, 1898),
    ('HYPERLINK("a")', 100, 1923), ('IF(A1,1)', 100, 1760), ('IF(A1,1,0)', 100, 1673),
    ('IFERROR(A1,0)', 100, 1773), ('IFNA(A1,0)', 100, 1773), ('IFS(A1,1)', 100, 1860),
    ('INDEX(A1:B2,1)', 100, 1785), ('INDIRECT("A1")', 100, 1897),
    ('INFO("osversion")', 100, 1735), ('INT(A1)', 100, 1910), ('IPMT(1,1,1,1)', 100, 1835),
    ('IRR(A1:B2)', 100, 1823), ('ISBLANK(A1)', 100, 1910), ('ISERROR(A1)', 100, 1910),
    ('ISEVEN(A1)', 100, 1910), ('ISFORMULA(A1)', 100, 1910), ('ISNA(A1)', 100, 1910),
    ('ISNUMBER(A1)', 100, 1910), ('ISOWEEKNUM(A1)', 100, 1910), ('ISREF(A1)', 100, 1910),
    ('KURT(A1)', 100, 1898), ('LARGE(A1:B2,1)', 100, 1798), ('LCM(A1)', 100, 1898),
    ('LEFT(A1)', 100, 1898), ('LEFT(A1,1)', 100, 1860), ('LEN(A1)', 100, 1910),
    ('LINEST(A1:B2)', 100, 1823), ('LOG(A1)', 100, 1898), ('LOOKUP(A1,A1:B2)', 100, 1735),
    ('LOWER(A1)', 100, 1910), ('MATCH(A1,A1:B2)', 100, 1735), ('MAX(A1)', 100, 1898),
    ('MAXA(A1)', 100, 1898), ('MAXIFS(A1:B2,A1:B2,1)', 100, 1623), ('MEDIAN(A1)', 100, 1898),
    ('MID(A1,1,1)', 100, 1835), ('MIN(A1)', 100, 1898), ('MINA(A1)', 100, 1898),
    ('MINIFS(A1:B2,A1:B2,1)', 100, 1623), ('MIRR(A1:B2,1,1)', 100, 1760),
    ('MMULT(A1,A1)', 100, 1823), ('MOD(A1,2)', 100, 1873), ('MODE(A1)', 100, 1898),
    ('MODE.MULT(A1)', 100, 1898), ('MODE.SNGL(A1)', 100, 1898), ('MROUND(A1,1)', 100, 1873),
    ('MUNIT(1)', 100, 1960), ('N(A1)', 100, 1910), ('NA()', 100, 1998),
    ('NETWORKDAYS(A1,A1)', 100, 1810), ('NORM.DIST(1,1,1,TRUE)', 50, 1954),
    ('NOT(A1)', 100, 1910), ('NOW()', 100, 1997), ('NPER(1,1,1)', 100, 1873),
    ('NPV(1,A1)', 100, 1860), ('NUMBERVALUE("1")', 50, 1985), ('OFFSET(A1,1,1)', 100, 1822),
    ('OR(A1)', 100, 1898), ('PERCENTILE(A1:B2,0)', 100, 1798),
    ('PERCENTRANK(A1:B2,1)', 100, 1785), ('PI()', 100, 1998), ('PMT(1,1,1)', 100, 1873),
    ('POWER(A1,2)', 100, 1873), ('PPMT(1,1,1,1)', 100, 1835), ('PRODUCT(A1)', 100, 1898),
    ('PROPER(A1)', 100, 1910), ('PV(1,1,1)', 100, 1873), ('QUARTILE(A1:B2,0)', 100, 1798),
    ('QUOTIENT(A1,1)', 100, 1873), ('RAND()', 100, 1997), ('RANDARRAY()', 100, 1985),
    ('RANDBETWEEN(1,2)', 100, 1922), ('RANK(A1,A1:B2)', 100, 1735),
    ('RANK.AVG(A1,A1:B2)', 100, 1735), ('RANK.EQ(A1,A1:B2)', 100, 1735),
    ('RATE(1,1,1)', 100, 1873), ('REPLACE(A1,1,1,"a")', 100, 1773), ('REPT(A1,2)', 100, 1873),
    ('RIGHT(A1)', 100, 1898), ('ROMAN(1)', 100, 1948), ('ROUND(A1,0)', 100, 1873),
    ('ROUNDUP(A1,0)', 100, 1873), ('ROW(A1)', 100, 1898), ('ROWS(A1:B2)', 100, 1835),
    ('SEARCH("a",A1)', 100, 1835), ('SEQUENCE(1)', 100, 1948), ('SHEET(A1)', 100, 1897),
    ('SHEETS(A1)', 100, 1897), ('SIGN(A1)', 100, 1910), ('SKEW(A1)', 100, 1898),
    ('SLN(1,1,1)', 100, 1885), ('SMALL(A1:B2,1)', 100, 1798), ('SORT(A1:B2)', 100, 1823),
    ('SORTBY(A1:B2,A1:B2)', 100, 1660), ('SQRT(A1)', 100, 1910), ('STDEV(A1)', 100, 1898),
    ('STDEV.P(A1)', 100, 1898), ('STDEV.S(A1)', 100, 1898), ('STDEVA(A1)', 100, 1898),
    ('STDEVP(A1)', 100, 1898), ('STDEVPA(A1)', 100, 1898), ('SUBSTITUTE(A1,"a","b")', 100, 1773),
    ('SUBTOTAL(9,A1:B2)', 100, 1785), ('SUM(A1)', 100, 1898), ('SUM(A1,A1)', 100, 1810),
    ('SUMIF(A1:B2,1)', 100, 1785), ('SUMIFS(A1:B2,A1:B2,1)', 100, 1623),
    ('SUMPRODUCT(A1:B2)', 100, 1823), ('SUMSQ(A1)', 100, 1898),
    ('SUMX2MY2(A1:B2,A1:B2)', 100, 1673), ('SWITCH(A1,1,2)', 100, 1823), ('T(A1)', 100, 1910),
    ('TAKE(A1:B2,1)', 100, 1785), ('TEXT(A1,"0")', 100, 1848), ('TEXTAFTER(A1,"a")', 100, 1835),
    ('TEXTBEFORE(A1,"a")', 50, 1941), ('TEXTJOIN(",",TRUE,A1)', 50, 1929),
    ('TEXTSPLIT(A1,"a")', 100, 1835), ('TIME(1,1,1)', 100, 1885), ('TODAY()', 100, 1997),
    ('TRANSPOSE(A1:B2)', 100, 1835), ('TREND(A1:B2)', 100, 1823), ('TRIM(A1)', 100, 1910),
    ('TRUNC(A1)', 100, 1898), ('TYPE(A1)', 100, 1910), ('UNICHAR(65)', 100, 1960),
    ('UNIQUE(A1:B2)', 100, 1823), ('UPPER(A1)', 100, 1910), ('VALUE(A1)', 100, 1910),
    ('VAR(A1)', 100, 1898), ('VAR.P(A1)', 100, 1898), ('VAR.S(A1)', 100, 1898),
    ('VARA(A1)', 100, 1898), ('VARP(A1)', 100, 1898), ('VARPA(A1)', 100, 1898),
    ('VLOOKUP(A1,A1:B2,2)', 100, 1698), ('VSTACK(A1)', 100, 1898), ('WEEKDAY(A1)', 100, 1898),
    ('WEEKNUM(A1)', 100, 1898), ('WORKDAY(A1,1)', 100, 1860),
    ('XLOOKUP(A1,A1:B2,A1:B2)', 100, 1573), ('XMATCH(A1,A1:B2)', 100, 1735),
    ('XNPV(1,A1:B2,A1:B2)', 100, 1635), ('XOR(A1)', 100, 1898), ('YEAR(A1)', 100, 1910),
    ('YEARFRAC(A1,A1)', 100, 1810),
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


def test_a_variable_argument_call_costs_one_byte_more_than_a_fixed_one():
    assert limits._call_bytes("XOR(") == 4
    assert limits._call_bytes("_xlfn.XOR(") == 4
    assert limits._call_bytes("_xlfn.STDEV.S(") == 4
    assert limits._call_bytes("_xlfn._xlws.FILTER(") == 4
    assert limits._call_bytes("ABS(") == 3
    assert limits._call_bytes("COUNTIF(") == 3
    assert limits._call_bytes("q(") == limits.NAME_CALL_BYTES


def test_iferror_and_ifna_carry_the_if_tokens():
    base = limits.compiled_size_floor("=ABS(A1)")
    assert limits.compiled_size_floor("=IFERROR(A1,0)") == base + 3 + 8
    assert limits.compiled_size_floor("=_xlfn.IFNA(A1,0)") == base + 3 + 8


def test_a_volatile_call_adds_its_token_once_per_formula():
    one = limits.compiled_size_floor("=NOW()")
    assert one == 3 + limits.ONCE_PER_FORMULA_BYTES
    assert limits.compiled_size_floor("=NOW()+NOW()") == 2 * 3 + 1 + limits.ONCE_PER_FORMULA_BYTES
    assert limits.compiled_size_floor("=NOW()+TODAY()+RAND()") == 3 * 3 + 2 + limits.ONCE_PER_FORMULA_BYTES


def test_a_let_name_that_is_a_function_name_still_calls_the_function():
    """A LET that declares a name Excel already has as a function: the call
    is Excel's function (a built-in wins), priced like any other
    variable-argument built-in."""
    assert limits.compiled_size_floor("=LET(SUM,LAMBDA(x,x),SUM(1))") ==         limits.compiled_size_floor("=LET(SUM,LAMBDA(x,x),MAX(1))")
