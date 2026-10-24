"""Excel's own look at a workbook this server cannot read (an .xlsb, a legacy
.xls, an encrypted file) must treat a required read that fails as unread, and
refuse unless allow_network. A fake COM object stands in for Excel; no Excel
runs. Everything here is made up."""
from __future__ import annotations

import types

import pytest

from xlsx_mcp.com import session
from xlsx_mcp.core import netreach
from xlsx_mcp.core.errors import FormulaRejected

BOOM = RuntimeError("made-up COM read failure")


def _obj(**attrs):
    """An object whose attributes are values, or raise when the value is an
    exception instance."""
    class O:
        pass
    for k, v in attrs.items():
        if isinstance(v, Exception):
            def getter(self, v=v):
                raise v
            setattr(O, k, property(getter))
        else:
            setattr(O, k, v)
    return O()


class _Used:
    def Find(self, **kw):  # noqa: N802
        return None


def _rules(*rules):
    fcs = types.SimpleNamespace(Count=len(rules), Item=lambda i: rules[i - 1])
    return fcs


def _sheet(*rules):
    return types.SimpleNamespace(Name="Rates", UsedRange=_Used(),
                                 Cells=types.SimpleNamespace(FormatConditions=_rules(*rules)))


def _book(sheets=(), names=(), connections=None):
    conns = connections if connections is not None else _obj(Count=0)
    return types.SimpleNamespace(Worksheets=list(sheets), Names=list(names), Connections=conns)


def _scan(book):
    return session._reach_of_open_workbook(book, netreach)


def _unread(book):
    reach = _scan(book)
    assert reach and reach.get("unread"), reach
    return reach["unread"]


# --- failed required reads are unread -------------------------------------

def test_a_name_whose_target_cannot_be_read_is_unread():
    book = _book(names=[_obj(Name="Rates", RefersTo=BOOM)])
    assert "Rates" in _unread(book)[0]


def test_names_that_cannot_be_listed_are_unread():
    class Bad:
        def __iter__(self):
            raise BOOM
    book = types.SimpleNamespace(Worksheets=[], Names=Bad(), Connections=_obj(Count=0))
    assert "defined names" in _unread(book)[0]


def test_a_connection_count_that_cannot_be_read_is_unread():
    assert "data connections" in _unread(_book(connections=_obj(Count=BOOM)))[0]


def test_a_connection_whose_name_cannot_be_read_still_counts_as_a_connection():
    class Conns:
        Count = 1

        def Item(self, i):  # noqa: N802
            return _obj(Name=BOOM)
    reach = _scan(_book(connections=Conns()))
    assert reach and reach["connections"], reach


def test_an_expression_rule_whose_formula_cannot_be_read_is_unread():
    rule = _obj(Type=2, Formula1=BOOM, Formula2=BOOM)
    unread = _unread(_book(sheets=[_sheet(rule)]))
    assert any("Formula1" in u for u in unread)
    assert not any("Formula2" in u for u in unread), unread


def test_a_between_rule_needs_both_formulas_and_a_greater_than_rule_only_one():
    between = _obj(Type=1, Operator=1, Formula1="=1", Formula2=BOOM)
    assert any("Formula2" in u for u in _unread(_book(sheets=[_sheet(between)])))
    greater = _obj(Type=1, Operator=5, Formula1="=1", Formula2=BOOM)
    assert _scan(_book(sheets=[_sheet(greater)])) is None


def test_a_rule_whose_type_cannot_be_read_is_unread():
    rule = _obj(Type=BOOM, Formula1=BOOM, Formula2=BOOM)
    assert _unread(_book(sheets=[_sheet(rule)]))


def test_a_rule_of_a_type_this_does_not_know_is_read_in_full():
    rule = _obj(Type=99, Formula1="=1", Formula2=BOOM)
    assert any("Formula2" in u for u in _unread(_book(sheets=[_sheet(rule)])))


@pytest.mark.parametrize("kind", [3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 16, 17])
def test_rule_types_that_hold_no_formula_are_not_unread_when_the_property_is_absent(kind):
    rule = _obj(Type=kind, Formula1=BOOM, Formula2=BOOM)
    assert _scan(_book(sheets=[_sheet(rule)])) is None


# --- controls: clean and positive stay what they were ----------------------

def test_a_clean_workbook_is_still_clean():
    rule = _obj(Type=2, Formula1="=A1>1", Formula2=BOOM)
    book = _book(sheets=[_sheet(rule)], names=[_obj(Name="Rate", RefersTo="=Sheet1!$A$1")])
    assert _scan(book) is None


def test_a_name_that_calls_a_network_function_is_still_found():
    book = _book(names=[_obj(Name="Rates", RefersTo='=WEBSERVICE("http://127.0.0.1:9/made-up")')])
    reach = _scan(book)
    assert reach["formulas"][0]["functions"] == ["WEBSERVICE"]
    assert not reach.get("unread")


def test_a_rule_that_calls_a_network_function_is_still_found():
    rule = _obj(Type=2, Formula1='=LEN(WEBSERVICE("http://127.0.0.1:9/made-up"))>0', Formula2=BOOM)
    reach = _scan(_book(sheets=[_sheet(rule)]))
    assert reach["formulas"][0]["functions"] == ["WEBSERVICE"]


# --- through the gate -------------------------------------------------------

@pytest.mark.parametrize("book", [
    _book(names=[_obj(Name="Rates", RefersTo=BOOM)]),
    _book(connections=_obj(Count=BOOM)),
    _book(sheets=[_sheet(_obj(Type=2, Formula1=BOOM, Formula2=BOOM))]),
], ids=["name", "connections", "rule"])
def test_the_gate_refuses_a_failed_read_unless_the_user_allowed_the_network(monkeypatch, book):
    monkeypatch.setattr(netreach, "file_reach", lambda p: (netreach.READ_IN_EXCEL, None))
    monkeypatch.setattr(session, "_reach_in_excel", lambda *a: _scan(book))
    with pytest.raises(FormulaRejected) as exc:
        session.network_gate(None, "made-up.xls", "", False, "export")
    assert "could not read" in str(exc.value)
    # allowed: the open goes ahead and the result is marked
    assert session.network_gate(None, "made-up.xls", "", True, "export") is True
