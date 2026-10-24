"""format_cells checks every nested key and honours per-side
borders, and no formatting mistake reaches the caller as a library
error."""

from __future__ import annotations

import openpyxl
import pytest

import family_harness as fh
from xlsx_mcp import server as srv


@pytest.fixture
def book(tmp_path):
    p = str(tmp_path / "f.xlsx")
    srv.create_workbook(p)
    srv.write_range(p, "A1", [["a", "b"], [1, 2]])
    return p


def _call(op, **args):
    return fh.start().call("xl_format", {"op": op, "args": args})


@pytest.mark.parametrize("args,needle", [
    ({"border": {"bogus": 1}}, "border has no key 'bogus'"),
    ({"border": {"sides": ["tpo"]}}, "did you mean 'top'"),
    ({"font": {"sise": 14}}, "did you mean 'size'"),
    ({"alignment": {"spin": True}}, "alignment has no key 'spin'"),
    ({"border": {"top": {"style": "thick"}, "sides": ["top"]}}, "not both"),
    ({"border": {"top": {"weight": "thick"}}}, "border.top has no key 'weight'"),
    ({"fill": "red"}, "did not match its schema"),
])
def test_x01_an_unknown_or_misspelt_key_refuses_and_names_the_valid_ones(args, needle, book):
    before = open(book, "rb").read()
    payload, err, _ = _call("format_cells", path=book, location="A1", **args)
    assert err and payload["error"]["code"] == "BAD_PARAMS", payload
    assert needle in payload["error"]["message"], payload["error"]["message"]
    assert open(book, "rb").read() == before


def test_x01_per_side_borders_are_honoured(book):
    payload, err, _ = _call("format_cells", path=book, location="A1", border={
        "top": {"style": "thick", "color": "FF0000"}, "bottom": {"style": "double"}})
    assert not err, payload
    b = openpyxl.load_workbook(book).active["A1"].border
    assert b.top.style == "thick" and b.top.color.rgb == "FFFF0000"
    assert b.bottom.style == "double"
    assert b.left.style is None and b.right.style is None


@pytest.mark.parametrize("args", [{"fill": {"colour": "00FF00"}},
                                  {"font": {"colour": "FF0000"}},
                                  {"font": {"fontSize": 14}}])
def test_x02_common_spellings_are_mapped_and_said_so(args, book):
    payload, err, _ = _call("format_cells", path=book, location="A1", **args)
    assert not err, payload
    assert "normalized_from" in payload, payload
    c = openpyxl.load_workbook(book).active["A1"]
    if "fill" in args:
        assert c.fill.fgColor.rgb == "FF00FF00"
    elif "colour" in args["font"]:
        assert c.font.color.rgb == "FFFF0000"
    else:
        assert c.font.size == 14


def test_define_takes_the_same_check(book):
    payload, err, _ = fh.start().call("xl_format", {"op": "apply_style", "args": {
        "path": book, "style": "Mine", "define": {"font": {"bolld": True}}}})
    assert err and "did you mean 'bold'" in payload["error"]["message"], payload
