"""The formula-length refusal speaks in the length the user wrote.

Excel's 8,192-character limit counts a formula as the file stores it, with
the prefixes Excel adds: _xlfn. on newer functions such as LET and LAMBDA,
_xlpm. on every LET or LAMBDA name. A formula written at 5,517 characters is
13,305 as stored, and the refusal said only "13,305 characters long", a
number the user never typed; an assistant could not reconcile the two and
told the user a wrong maximum.

MEASURED in Excel 16 (hidden, on a private desktop):
_xlfn.LET(_xlpm.x,A1,_xlpm.x+...) opens at 8,192 stored characters (2,054
as written) and is refused at 8,193 (2,055 as written); a plain formula opens
at 8,192 and is refused at 8,193. So the count stays on the stored form, and
the refusal now says what was written, the most that fits as written for a
formula like it, and the stored count only as the reason.

No Excel here except the test marked LIVE, which starts the product's hidden
Excel (run it on a private desktop).
"""
from __future__ import annotations

import openpyxl
import pytest

from xlsx_mcp.com import session as com_session
from xlsx_mcp.core import calc, limits
from xlsx_mcp.core.errors import ExcelWouldRefuse

LIVE = pytest.mark.skipif(not com_session.com_available()[0], reason="needs Excel (COM)")

#: The formula an assistant wrote: 324 inline LAMBDA calls inside a LET.
REPORTED = "LET(x,A1," + "+".join(["LAMBDA(a,a+1)(x)"] * 324) + ")"


def _let_stored_at(target: int) -> str:
    """=LET(x,A1,x+x+...) whose stored form is exactly `target` characters:
    21 + 8n stored for n uses, plus a +111... tail inside the LET."""
    n = (target - 21) // 8
    rem = target - (21 + 8 * n)
    if rem == 1:
        n, rem = n - 1, rem + 8
    tail = "" if rem == 0 else "+" + "1" * (rem - 1)
    return "=LET(x,A1," + "+".join(["x"] * n) + tail + ")"


@pytest.fixture()
def book(tmp_path):
    p = tmp_path / "long.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = 1
    wb.save(p)
    return p


def test_the_reported_formula_is_5517_written_and_13305_stored():
    assert len(REPORTED) == 5517
    assert len(calc.normalize_formula(REPORTED)[0]) == 13305


def _assert_written_terms(msg: str, what: str):
    first = msg.split(":")[0]
    assert msg.startswith(f"{what} is 5,517 characters as written"), msg
    assert "13,305" not in first, msg
    assert "about 3,396 characters as written" in msg, msg
    assert "8,192" in msg and "13,305" in msg, msg
    assert "13,305 characters long" not in msg, msg
    assert "_xlfn." in msg and "_xlpm." in msg, msg


@pytest.mark.parametrize("route", ["set_formula", "set_cell"])
def test_the_refusal_says_what_was_written_and_nothing_is_saved(book, route):
    from xlsx_mcp.ops import cells, formulas
    before = book.read_bytes()
    with pytest.raises(ExcelWouldRefuse) as ei:
        if route == "set_formula":
            formulas.set_formula(str(book), "B1", "=" + REPORTED, sheet="Data")
        else:
            cells.set_cell(str(book), "B1", "=" + REPORTED, sheet="Data")
    _assert_written_terms(str(ei.value), "the formula for B1")
    assert book.read_bytes() == before


def test_the_check_on_a_stored_formula_reports_it_as_written():
    stored = calc.normalize_formula(REPORTED)[0]
    with pytest.raises(ExcelWouldRefuse) as ei:
        limits.check_formula("=" + stored, what="the formula in Data!B1")
    _assert_written_terms(str(ei.value), "the formula in Data!B1")


def test_a_plain_formula_keeps_its_plain_words():
    plain = "=" + "+".join(["1111111"] * 1024) + "11"  # 8,193 characters, no prefixes
    with pytest.raises(ExcelWouldRefuse) as ei:
        limits.check_formula(plain, what="the formula for B1")
    assert str(ei.value).startswith(
        "the formula for B1 is 8,193 characters long; Excel's limit is 8,192"), str(ei.value)


def test_newer_functions_alone_name_only_their_prefix():
    body = "+".join(["XLOOKUP(1,A1:A2,A1:A2)"] * 380)  # 8,739 written, more stored
    with pytest.raises(ExcelWouldRefuse) as ei:
        limits.check_formula("=" + body, what="the formula for B1")
    msg = str(ei.value)
    assert msg.startswith(f"the formula for B1 is {len(body):,} characters as written"), msg
    assert "_xlfn." in msg and "_xlpm." not in msg, msg


def test_the_limit_is_the_stored_length_as_measured(book):
    from xlsx_mcp.ops import formulas
    at = _let_stored_at(8192)
    past = _let_stored_at(8193)
    assert len(calc.normalize_formula(at)[0][1:]) == 8192  # the stored body, without '='
    assert len(calc.normalize_formula(past)[0][1:]) == 8193
    out = formulas.set_formula(str(book), "B1", at, sheet="Data")
    assert out["ok"]
    before = book.read_bytes()
    with pytest.raises(ExcelWouldRefuse, match="characters as written"):
        formulas.set_formula(str(book), "B2", past, sheet="Data")
    assert book.read_bytes() == before


def test_a_defined_name_over_the_limit_is_refused_in_written_terms(book):
    from xlsx_mcp.ops import names
    before = book.read_bytes()
    with pytest.raises(ExcelWouldRefuse) as ei:
        names.manage_name(str(book), "add", name="big", refers_to=REPORTED)
    _assert_written_terms(str(ei.value), "the definition of big")
    assert book.read_bytes() == before


@pytest.mark.parametrize("formula", [
    "=" + "+".join(["XLOOKUP(1,A1:A2,B1:B2)"] * 350),
    "=LET(g,LAMBDA(x,x+1)," + "+".join(["g(1)"] * 741) + ")",
    "=LET(g,LAMBDA(x,x+1)," + "+".join(["g(1)"] * 1167) + ")",
], ids=["xlookup_350", "let_g_741", "let_g_1167"])
def test_formulas_excel_will_not_open_as_files_are_refused(formula):
    """MEASURED as files in Excel 16: 350 typed XLOOKUP calls are 10,149
    characters stored and do not open (282 do: 8,177 stored); LET(g,LAMBDA(..))
    with 740 calls of g opens at 8,190 stored characters, 741 (8,201) and 1,167
    (12,887) do not. Excel takes 1,167 when typed, but cannot open the file it
    then saves."""
    with pytest.raises(ExcelWouldRefuse, match="characters as written"):
        limits.check_formula(formula)


@LIVE
@pytest.mark.timeout(300)
def test_the_longest_accepted_formula_opens_in_excel(book):
    """The product writes the formula at exactly 8,192 stored characters, and
    real Excel opens the workbook with it intact."""
    from xlsx_mcp.ops import formulas
    at = _let_stored_at(8192)
    formulas.set_formula(str(book), "B1", at, sheet="Data")
    try:
        verdict = com_session.opens_clean(str(book), timeout=120)
        assert verdict["opens_clean"] is True, verdict

        def read(manager):
            w = manager.acquire()
            wb = com_session.open_workbook(w.app, str(book), read_only=True)
            try:
                return str(wb.Worksheets("Data").Range("B1").Formula2)
            finally:
                wb.Close(False)
                manager.release_to_pool(w)
        back = com_session.run_com("read-back", read, timeout=120)
        assert back == at, (len(back), len(at))
    finally:
        com_session.get_executor().shutdown()
