"""Two ways a save or a read could answer ok while the workbook said something
else (every workbook here is made up and built by the test).

A custom document property can be linked to a cell through a defined name
(docProps/custom.xml, linkTarget="Name"). Excel refreshes that cell from the
property when the workbook opens. An edit of the cell used to leave the
property as it was, answered "verified", and Excel put the old value back. A
save that changes such a cell now writes the new value into the linked
property, in the property's own type, on every save route; a value the
property cannot hold (text into a number property, a formula, an error, an
empty cell into a number) is refused naming the property, or, with allow_loss,
saved with a warning that names it and leaves the property as it was.

A shared formula whose first cell is missing (or comes after the cell that
uses it) cannot be recovered: Excel drops the formula and keeps the stored
value. Every read showed the formula "=" for such a cell while the hazard scan
said the workbook was clean, and set_formula "=" stored the text "=".
"""
from __future__ import annotations

import datetime
import re
import zipfile
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from xlsx_mcp.core import hazard as _hazard
from xlsx_mcp.core.errors import HazardRefused, XlMcpError
from xlsx_mcp.ops import cells, formulas, inspectors, lifecycle, view

CUSTOM_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CUSTOM_REL = ("http://schemas.openxmlformats.org/officeDocument/2006/"
              "relationships/custom-properties")
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

PLAIN_PROPS = (
    ("Text", "", "lpwstr", "Made-up text"),
    ("Real", "", "r8", "3.14159"),
    ("Flag", "", "bool", "false"),
)


@pytest.fixture(params=["raw", "rewrite"])
def route(request, monkeypatch):
    """Both save routes: the byte-preserving cell edit and the full rewrite."""
    monkeypatch.setenv("KS4XL_SURGICAL_SAVE", "on" if request.param == "raw" else "off")
    return request.param


_ZONE = {}


def _linkprops():
    from xlsx_mcp.core import linkprops
    return linkprops


@pytest.fixture(autouse=True)
def _utc_zone(monkeypatch):
    """The date tests below do not depend on the zone of the machine."""
    mod = _linkprops()
    _ZONE.setdefault("real", mod._utc_offset)
    monkeypatch.setattr(mod, "_utc_offset", lambda when: datetime.timedelta(0))


def _props_xml(linked: list[tuple[str, str, str, str]]) -> bytes:
    rows = []
    pid = 2
    for name, target, vt, text in list(PLAIN_PROPS) + list(linked):
        link = f' linkTarget="{target}"' if target else ""
        rows.append(f'<property fmtid="{FMTID}" pid="{pid}" name="{name}"{link}>'
                    f"<vt:{vt}>{text}</vt:{vt}></property>")
        pid += 1
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/'
            f'2006/custom-properties" xmlns:vt="{VT_NS}">'
            + "".join(rows) + "</Properties>").encode("utf-8")


def _with_custom(src: Path, dst: Path, linked) -> Path:
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    f'<Override PartName="/docProps/custom.xml" '
                    f'ContentType="{CUSTOM_CT}"/></Types>'.encode())
            elif info.filename == "_rels/.rels":
                data = data.replace(
                    b"</Relationships>",
                    f'<Relationship Id="rId90" Type="{CUSTOM_REL}" '
                    f'Target="docProps/custom.xml"/></Relationships>'.encode())
            zout.writestr(info.filename, data)
        zout.writestr("docProps/custom.xml", _props_xml(linked))
    return dst


def _linked_book(tmp_path: Path, linked, *, value=0.075, fmt=None, date1904=False) -> Path:
    """A made-up workbook: Data!B1 holds `value`, the defined name LinkedRate
    names it, and the custom properties in `linked` point at that name."""
    wb = openpyxl.Workbook()
    if date1904:
        from openpyxl.utils.datetime import CALENDAR_MAC_1904
        wb.epoch = CALENDAR_MAC_1904
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Rate"
    ws["B1"] = value
    if fmt:
        ws["B1"].number_format = fmt
    ws["A2"] = "Total"
    ws["B2"] = "=B1*2"
    ws["A3"] = "Other"
    ws["B3"] = 7
    wb.defined_names["LinkedRate"] = DefinedName("LinkedRate", attr_text="Data!$B$1")
    plain = tmp_path / "plain.xlsx"
    wb.save(plain)
    return _with_custom(plain, tmp_path / "linked.xlsx", linked)


def _custom(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        return zf.read("docProps/custom.xml").decode("utf-8")


def _prop(path: Path, name: str) -> tuple[str, str]:
    m = re.search(rf'name="{re.escape(name)}"[^>]*><vt:(\w+)>([^<]*)</vt:', _custom(path))
    assert m, f"property {name} not found in {_custom(path)}"
    return m.group(1), m.group(2)


RATE = ("Linked rate", "LinkedRate", "r8", "0.075")
TEXT = ("Linked text", "LinkedRate", "lpwstr", "0.075")



def test_editing_a_linked_cell_writes_the_new_value_into_the_property(tmp_path, route):
    p = _linked_book(tmp_path, [RATE, TEXT])
    before = _custom(p)
    res = cells.set_cell(str(p), "B1", 1.075, sheet="Data", backup=False)
    assert res["ok"] is True and res["verified"] is True
    assert res["route"] == ("raw_ooxml" if route == "raw" else "openpyxl")
    assert _prop(p, "Linked rate") == ("r8", "1.075")
    assert _prop(p, "Linked text") == ("lpwstr", "1.075")
    after = _custom(p)
    # every other property is byte-for-byte as it was
    for name in ("Text", "Real", "Flag"):
        keep = re.search(rf'<property [^>]*name="{name}".*?</property>', before).group(0)
        assert keep in after
    assert res["changed"]["linked_properties"] == [
        {"property": "Linked rate", "cell": "Data!B1", "value": "1.075"},
        {"property": "Linked text", "cell": "Data!B1", "value": "1.075"}]
    wb = openpyxl.load_workbook(p)
    assert wb["Data"]["B1"].value == 1.075


def test_an_edit_of_another_cell_leaves_custom_xml_untouched(tmp_path, route):
    p = _linked_book(tmp_path, [RATE, TEXT])
    before = _custom(p)
    res = cells.set_cell(str(p), "B3", 8, sheet="Data", backup=False)
    assert res["ok"] is True
    assert _custom(p) == before
    assert "linked_properties" not in res["changed"]


def test_a_whole_number_goes_into_a_real_property_without_a_decimal_point(tmp_path, route):
    p = _linked_book(tmp_path, [RATE, TEXT])
    cells.set_cell(str(p), "B1", 5, sheet="Data", backup=False)
    assert _prop(p, "Linked rate") == ("r8", "5")
    assert _prop(p, "Linked text") == ("lpwstr", "5")


def test_text_goes_into_a_text_property(tmp_path, route):
    p = _linked_book(tmp_path, [TEXT], value="draft")
    cells.set_cell(str(p), "B1", "final & <signed>", sheet="Data", backup=False)
    assert _prop(p, "Linked text") == ("lpwstr", "final &amp; &lt;signed&gt;")


def test_a_boolean_goes_into_a_boolean_property(tmp_path, route):
    p = _linked_book(tmp_path, [("Linked flag", "LinkedRate", "bool", "false")],
                     value=False)
    cells.set_cell(str(p), "B1", True, sheet="Data", backup=False)
    assert _prop(p, "Linked flag") == ("bool", "true")


def test_a_date_goes_into_a_date_property(tmp_path, route):
    p = _linked_book(tmp_path, [("Linked day", "LinkedRate", "date", "2026-05-17T00:00:00Z")],
                     value=46159, fmt="yyyy-mm-dd")
    day = datetime.date(2031, 3, 4).isoformat()
    cells.set_cell(str(p), "B1", {"date": day}, sheet="Data", backup=False)
    assert _prop(p, "Linked day") == ("date", day + "T00:00:00Z")


def test_an_integer_property_takes_a_whole_number_only(tmp_path, route):
    p = _linked_book(tmp_path, [("Linked count", "LinkedRate", "i4", "3")], value=3)
    cells.set_cell(str(p), "B1", 9, sheet="Data", backup=False)
    assert _prop(p, "Linked count") == ("i4", "9")
    before = _custom(p)
    with pytest.raises(HazardRefused) as exc:
        cells.set_cell(str(p), "B1", 9.5, sheet="Data", backup=False)
    assert "Linked count" in str(exc.value)
    assert _custom(p) == before


@pytest.mark.parametrize("value", ["not a number", "=B3+1", None])
def test_a_value_the_property_cannot_hold_is_refused_by_name(tmp_path, route, value):
    p = _linked_book(tmp_path, [RATE])
    snapshot = p.read_bytes()
    with pytest.raises(HazardRefused) as exc:
        cells.set_cell(str(p), "B1", value, sheet="Data", backup=False)
    text = str(exc.value)
    assert "Linked rate" in text and "Data!B1" in text
    assert "allow_loss" in text
    assert p.read_bytes() == snapshot  # nothing was written


def test_allow_loss_saves_and_names_the_property_it_left_alone(tmp_path, route):
    p = _linked_book(tmp_path, [RATE, TEXT])
    res = cells.set_cell(str(p), "B1", "not a number", sheet="Data",
                         allow_loss=True, backup=False)
    assert res["ok"] is True
    assert _prop(p, "Linked rate") == ("r8", "0.075")  # left as it was
    assert _prop(p, "Linked text") == ("lpwstr", "not a number")  # text fits text
    assert any("Linked rate" in w and "Data!B1" in w for w in res["warnings"])


def test_a_cleared_cell_empties_a_text_property_and_is_refused_for_a_number(tmp_path, route):
    p = _linked_book(tmp_path, [TEXT])
    cells.set_cell(str(p), "B1", None, sheet="Data", backup=False)
    assert _prop(p, "Linked text") == ("lpwstr", "")


def test_a_row_inserted_above_the_linked_cell_moves_it_and_changes_no_property(tmp_path):
    p = _linked_book(tmp_path, [RATE, TEXT])
    before = _custom(p)
    from xlsx_mcp.ops import structure
    res = structure.modify_grid_structure(str(p), "insert_rows", 1, sheet="Data",
                                          backup=False, verify_com=False)
    assert res["ok"] is True
    wb = openpyxl.load_workbook(p)
    assert wb["Data"]["B2"].value == 0.075
    assert "Data!$B$2" in wb.defined_names["LinkedRate"].attr_text
    assert _custom(p) == before
    assert "linked_properties" not in res["changed"]


def test_a_structural_edit_that_also_changes_the_cell_is_not_a_property_change(tmp_path):
    """An insert alone never rewrites a property; a later cell edit at the
    cell's new address does."""
    p = _linked_book(tmp_path, [RATE])
    from xlsx_mcp.ops import structure
    structure.modify_grid_structure(str(p), "insert_rows", 1, sheet="Data",
                                    backup=False, verify_com=False)
    cells.set_cell(str(p), "B2", 2.5, sheet="Data", backup=False)
    assert _prop(p, "Linked rate") == ("r8", "2.5")


def test_a_workbook_without_linked_properties_saves_as_before(tmp_path, route):
    p = _linked_book(tmp_path, [])
    before = _custom(p)
    res = cells.set_cell(str(p), "B1", 1.5, sheet="Data", backup=False)
    assert res["ok"] is True and _custom(p) == before
    assert "linked_properties" not in res["changed"]



def _sheet_book(tmp_path: Path, name: str, member_cells: str) -> Path:
    """A made-up one-sheet workbook whose column B holds the numbers 2, 4, 6, 8,
    10 beside A1:A5 = 1..5, with sheetData rows replaced by member_cells."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Shared"
    for i in range(1, 6):
        ws.cell(i, 1, i)
    src = tmp_path / f"{name}_src.xlsx"
    wb.save(src)
    out = tmp_path / f"{name}.xlsx"
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                rows = "".join(
                    f'<row r="{i}"><c r="A{i}"><v>{i}</v></c>{member_cells_for(i, member_cells)}</row>'
                    for i in range(1, 6))
                text = re.sub(r"<sheetData>.*</sheetData>",
                              lambda m: f"<sheetData>{rows}</sheetData>", text, flags=re.S)
                text = re.sub(r"<dimension ref=\"[^\"]*\"/>", '<dimension ref="A1:B5"/>', text)
                data = text.encode("utf-8")
            zout.writestr(info.filename, data)
    return out


def member_cells_for(row: int, kind: str) -> str:
    v = row * 2
    if kind == "nomaster":      # x74a: every cell is a member, no master anywhere
        return (f'<c r="B{row}"><f t="shared" si="0"/><v>{v}</v></c>'
                if row >= 2 else "")
    if kind == "before":        # x74b: B1 comes before its master B2
        if row == 1:
            return f'<c r="B1"><f t="shared" si="0"/><v>{v}</v></c>'
        if row == 2:
            return f'<c r="B2"><f t="shared" ref="B1:B5" si="0">A2*2</f><v>{v}</v></c>'
        return f'<c r="B{row}"><f t="shared" si="0"/><v>{v}</v></c>'
    raise ValueError(kind)


@pytest.fixture
def nomaster(tmp_path):
    return _sheet_book(tmp_path, "x74a", "nomaster")


@pytest.fixture
def before_master(tmp_path):
    return _sheet_book(tmp_path, "x74b", "before")


def test_a_member_with_no_master_reads_as_its_stored_value_and_is_flagged(nomaster):
    r = cells.read_range(str(nomaster), "A1:B5", values="formula")
    assert [row[1] for row in r["values"]] == [None, 4, 6, 8, 10]
    flag = r["broken_shared_formulas"]
    assert flag["cells"] == ["B2", "B3", "B4", "B5"]
    assert "shared formula" in flag["note"] and "no formula" in flag["note"]
    c = cells.read_range(str(nomaster), "B2:B3", values="cached")
    assert c["values"] == [[4], [6]]
    assert c["broken_shared_formulas"]["cells"] == ["B2", "B3"]


def test_a_range_without_a_broken_cell_carries_no_flag(nomaster):
    r = cells.read_range(str(nomaster), "A1:A5", values="formula")
    assert "broken_shared_formulas" not in r


def test_get_cells_reads_the_value_and_flags(nomaster):
    r = cells.get_cells(str(nomaster), ["B2", "B3", "A1"], values="formula")
    assert [row[1] for row in r["cells"]] == [4, 6, 1]
    assert r["broken_shared_formulas"]["cells"] == ["B2", "B3"]


def test_a_member_before_its_master_is_broken_but_the_rest_still_read_as_formulas(before_master):
    r = cells.read_range(str(before_master), "B1:B5", values="formula")
    assert [row[0] for row in r["values"]] == [2, "=A2*2", "=A3*2", "=A4*2", "=A5*2"]
    assert r["broken_shared_formulas"]["cells"] == ["B1"]


def test_get_grid_view_lists_no_formula_for_them_and_names_the_hazard(nomaster):
    v = view.get_grid_view(str(nomaster), values="formula")
    assert "=" not in v["formula_cells"].values()
    assert v["broken_shared_formulas"]["cells"] == ["B2", "B3", "B4", "B5"]
    assert any("shared formula" in h for h in v["hazards"])
    assert "|4|" in v["table"].replace(" ", "")


def test_audit_formulas_does_not_report_the_formula_equals_and_counts_the_cells(nomaster):
    a = formulas.audit_formulas(str(nomaster))
    assert a["formulas"]["count"] == 0
    assert a["broken_shared_formulas"]["count"] == 4
    assert a["broken_shared_formulas"]["items"][0] == {"sheet": "Shared", "cell": "B2"}


def test_the_hazard_scan_lists_them_and_the_workbook_is_not_clean(nomaster, before_master):
    for p in (nomaster, before_master):
        rep = _hazard.scan_path(str(p))
        assert not rep.clean
        assert [h.key for h in rep.hazards] == [_hazard.BROKEN_SHARED_KEY]
        assert rep.hazards[0].severity == _hazard.SEV_DEGRADES


def test_an_ordinary_shared_formula_is_not_a_hazard(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    for i in range(1, 6):
        ws.cell(i, 1, i)
    p = tmp_path / "plain.xlsx"
    wb.save(p)
    with zipfile.ZipFile(p) as zin, zipfile.ZipFile(tmp_path / "s.xlsx", "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                text = data.decode()
                rows = "".join(
                    f'<row r="{i}"><c r="A{i}"><v>{i}</v></c>'
                    + (f'<c r="B{i}"><f t="shared" ref="B1:B5" si="0">A1*2</f><v>{i * 2}</v></c>'
                       if i == 1 else
                       f'<c r="B{i}"><f t="shared" si="0"/><v>{i * 2}</v></c>')
                    + "</row>" for i in range(1, 6))
                text = re.sub(r"<sheetData>.*</sheetData>",
                              lambda m: f"<sheetData>{rows}</sheetData>", text, flags=re.S)
                data = text.encode()
            zout.writestr(info.filename, data)
    s = tmp_path / "s.xlsx"
    assert _hazard.scan_path(str(s)).clean
    r = cells.read_range(str(s), "B1:B3", values="formula")
    assert [row[0] for row in r["values"]] == ["=A1*2", "=A2*2", "=A3*2"]
    assert "broken_shared_formulas" not in r


def test_set_formula_with_nothing_after_the_equals_sign_is_refused_by_name(nomaster):
    snapshot = nomaster.read_bytes()
    for text in ("=", "= ", "=   ", " = "):
        with pytest.raises(XlMcpError) as exc:
            formulas.set_formula(str(nomaster), "B2", text, sheet="Shared", backup=False)
        assert "empty formula" in str(exc.value)
        assert nomaster.read_bytes() == snapshot


def test_a_literal_equals_sign_through_set_cell_is_still_stored_as_text(tmp_path):
    wb = openpyxl.Workbook()
    wb.save(tmp_path / "t.xlsx")
    res = cells.set_cell(str(tmp_path / "t.xlsx"), "A1", "=", backup=False)
    assert res["ok"] is True
    assert openpyxl.load_workbook(tmp_path / "t.xlsx").active["A1"].value == "="


def test_an_edit_elsewhere_keeps_the_stored_values_on_both_routes(nomaster, route):
    res = cells.set_cell(str(nomaster), "A1", 10, sheet="Shared", backup=False)
    assert res["ok"] is True
    got = openpyxl.load_workbook(nomaster, data_only=True).active
    assert [got.cell(i, 2).value for i in range(2, 6)] == [4, 6, 8, 10]
    if route == "rewrite":
        assert any("shared formula" in w for w in res["warnings"])


# ------------------------------------------- listed sheets with no usable id

def _two_sheets(tmp_path: Path, count: int) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "First"
    ws["A1"] = 1
    if count == 2:
        wb.create_sheet("Second")["A1"] = 2
    p = tmp_path / f"sheets{count}.xlsx"
    wb.save(p)
    return p


def _strip_sheet_id(src: Path, dst: Path, which: str, how: str) -> Path:
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/workbook.xml":
                text = data.decode("utf-8")
                m = re.search(rf'<sheet [^>]*name="{which}"[^>]*/>', text)
                el = m.group(0)
                if how == "missing":
                    new = re.sub(r'\s+r:id="[^"]*"', "", el)
                elif how == "empty":
                    new = re.sub(r'r:id="[^"]*"', 'r:id=""', el)
                else:  # unknown
                    new = re.sub(r'r:id="[^"]*"', 'r:id="rId99"', el)
                data = text.replace(el, new).encode("utf-8")
            zout.writestr(info.filename, data)
    return dst


@pytest.mark.parametrize("how", ["missing", "empty", "unknown"])
@pytest.mark.parametrize("count,which", [(1, "First"), (2, "First"), (2, "Second")])
def test_a_listed_sheet_without_a_usable_relationship_id_is_refused_by_name(
        tmp_path, count, which, how):
    p = _strip_sheet_id(_two_sheets(tmp_path, count), tmp_path / "bad.xlsx", which, how)
    before = p.read_bytes()
    with pytest.raises(XlMcpError) as exc:
        cells.read_range(str(p), "A1", sheet="First")
    assert repr(which) in str(exc.value)
    with pytest.raises(XlMcpError) as exc:
        lifecycle.get_workbook_metadata(str(p))
    assert repr(which) in str(exc.value)
    with pytest.raises(XlMcpError):
        cells.set_cell(str(p), "A1", 5, sheet="First", backup=False)
    assert p.read_bytes() == before


@pytest.mark.parametrize("count", [1, 2])
def test_sheets_with_usable_ids_still_read_and_edit(tmp_path, count):
    p = _two_sheets(tmp_path, count)
    assert lifecycle.get_workbook_metadata(str(p))["sheet_count"] == count
    assert cells.set_cell(str(p), "A1", 5, sheet="First", backup=False)["ok"] is True


# ------------------------------------- the same files in other XML spellings

def _recode(path: Path, member: str, fn) -> None:
    tmp = path.with_name("recoded.xlsx")
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            zout.writestr(info, fn(data) if info.filename == member else data)
    tmp.replace(path)


def _utf16(data: bytes, order: str) -> bytes:
    text = data.decode("utf-8").replace('encoding="UTF-8"', 'encoding="UTF-16"')
    bom = b"\xff\xfe" if order == "le" else b"\xfe\xff"
    return bom + text.encode("utf-16-" + order)


def _props(path: Path) -> dict:
    from lxml import etree
    with zipfile.ZipFile(path) as zf:
        root = etree.fromstring(zf.read("docProps/custom.xml"))
    return {p.get("name"): (etree.QName(p[0]).localname, p[0].text)
            for p in root if isinstance(p.tag, str)}


def _spelled(tmp_path: Path, spelling: str) -> Path:
    p = _linked_book(tmp_path, [RATE, TEXT])
    if spelling in ("custom-utf16le", "custom-utf16be"):
        _recode(p, "docProps/custom.xml", lambda d: _utf16(d, spelling[-2:]))
    elif spelling == "custom-utf16-prefixed":
        _recode(p, "docProps/custom.xml", lambda d: _utf16(
            d.replace(b"<Properties ", b"<p:Properties ")
            .replace(b'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
                     b'custom-properties"',
                     b'xmlns:p="http://schemas.openxmlformats.org/officeDocument/2006/'
                     b'custom-properties"')
            .replace(b"</Properties>", b"</p:Properties>")
            .replace(b"<property ", b"<p:property ")
            .replace(b"</property>", b"</p:property>"), "le"))
    elif spelling == "custom-cdata-value":
        _recode(p, "docProps/custom.xml", lambda d: d.replace(
            b"<vt:lpwstr>0.075</vt:lpwstr>", b"<vt:lpwstr><![CDATA[0.075]]></vt:lpwstr>"))
    elif spelling == "custom-comment":
        _recode(p, "docProps/custom.xml", lambda d: d.replace(
            b"<Properties ", b"<!-- kept -->\n<Properties ", 1).replace(
            b"</Properties>", b"<!-- <property name=\"x\" linkTarget=\"Nope\"/> -->"
                              b"</Properties>"))
    elif spelling == "name-cdata":
        _recode(p, "xl/workbook.xml", lambda d: d.replace(
            b"Data!$B$1", b"<![CDATA[Data!$B$1]]>"))
    elif spelling == "workbook-utf16le":
        _recode(p, "xl/workbook.xml", lambda d: _utf16(d, "le"))
    elif spelling == "workbook-utf16be":
        _recode(p, "xl/workbook.xml", lambda d: _utf16(d, "be"))
    else:
        assert spelling == "plain"
    return p


SPELLINGS = ["plain", "custom-utf16le", "custom-utf16be", "custom-utf16-prefixed",
             "custom-cdata-value", "custom-comment", "name-cdata", "workbook-utf16le",
             "workbook-utf16be"]


@pytest.mark.parametrize("spelling", SPELLINGS)
def test_a_linked_property_is_found_however_the_xml_is_spelled(tmp_path, route, spelling):
    p = _spelled(tmp_path, spelling)
    res = cells.set_cell(str(p), "B1", 1.075, sheet="Data", backup=False)
    assert res["ok"] is True and res["verified"] is True
    assert _props(p)["Linked rate"] == ("r8", "1.075")
    assert _props(p)["Linked text"] == ("lpwstr", "1.075")
    assert len(res["changed"]["linked_properties"]) == 2


@pytest.mark.parametrize("spelling", SPELLINGS)
def test_a_value_the_property_cannot_hold_is_refused_in_every_spelling(
        tmp_path, route, spelling):
    p = _spelled(tmp_path, spelling)
    snapshot = p.read_bytes()
    with pytest.raises(HazardRefused) as exc:
        cells.set_cell(str(p), "B1", "not a number", sheet="Data", backup=False)
    assert "Linked rate" in str(exc.value) and "Data!B1" in str(exc.value)
    assert p.read_bytes() == snapshot


@pytest.mark.parametrize("spelling,order", [("custom-utf16le", b"\xff\xfe"),
                                            ("custom-utf16be", b"\xfe\xff")])
def test_a_utf16_property_part_stays_utf16_with_its_byte_order_mark(
        tmp_path, route, spelling, order):
    p = _spelled(tmp_path, spelling)
    cells.set_cell(str(p), "B1", 2.5, sheet="Data", backup=False)
    with zipfile.ZipFile(p) as zf:
        data = zf.read("docProps/custom.xml")
    assert data.startswith(order)
    assert "UTF-16" in data.decode("utf-16")
    assert _props(p)["Linked rate"] == ("r8", "2.5")


def test_the_other_properties_and_comments_survive_an_edit(tmp_path, route):
    p = _spelled(tmp_path, "custom-comment")
    cells.set_cell(str(p), "B1", 3.5, sheet="Data", backup=False)
    with zipfile.ZipFile(p) as zf:
        text = zf.read("docProps/custom.xml").decode("utf-8")
    assert "<!-- kept -->" in text and "Nope" in text
    assert text.startswith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    props = _props(p)
    assert props["Text"] == ("lpwstr", "Made-up text") and props["Linked rate"][1] == "3.5"


def test_a_comment_that_looks_like_a_link_is_not_a_link(tmp_path, route):
    p = _linked_book(tmp_path, [])
    _recode(p, "docProps/custom.xml", lambda d: d.replace(
        b"</Properties>", b'<!-- <property name="Ghost" linkTarget="LinkedRate">'
                          b"<vt:r8>1</vt:r8></property> --></Properties>"))
    before = _custom(p)
    res = cells.set_cell(str(p), "B1", "text", sheet="Data", backup=False)
    assert res["ok"] is True and "linked_properties" not in res["changed"]
    assert _custom(p) == before


# ------------------------------------------------------ dates before March 1900

@pytest.mark.parametrize("tag", ["date", "filetime"])
@pytest.mark.parametrize("serial,expected", [
    (1, "1900-01-01T00:00:00Z"), (2, "1900-01-02T00:00:00Z"),
    (31, "1900-01-31T00:00:00Z"), (32, "1900-02-01T00:00:00Z"),
    (59, "1900-02-28T00:00:00Z"), (59.5, "1900-02-28T12:00:00Z"),
    (61, "1900-03-01T00:00:00Z"), (46159, "2026-05-17T00:00:00Z")])
def test_a_serial_becomes_the_date_excel_shows_for_it(tmp_path, route, tag, serial, expected):
    p = _linked_book(tmp_path, [("Linked day", "LinkedRate", tag, "2020-05-17T00:00:00Z")],
                     value=45000, fmt="yyyy-mm-dd")
    cells.set_cell(str(p), "B1", serial, sheet="Data", backup=False)
    assert _prop(p, "Linked day") == (tag, expected)


@pytest.mark.parametrize("tag", ["date", "filetime"])
@pytest.mark.parametrize("serial,expected", [
    (0, "1904-01-01T00:00:00Z"), (1, "1904-01-02T00:00:00Z"),
    (59, "1904-02-29T00:00:00Z"), (61, "1904-03-02T00:00:00Z")])
def test_the_1904_system_counts_from_1904(tmp_path, route, tag, serial, expected):
    p = _linked_book(tmp_path, [("Linked day", "LinkedRate", tag, "2020-05-17T00:00:00Z")],
                     value=45000, fmt="yyyy-mm-dd", date1904=True)
    cells.set_cell(str(p), "B1", serial, sheet="Data", backup=False)
    assert _prop(p, "Linked day") == (tag, expected)


@pytest.mark.parametrize("serial", [60, 60.5, 0, 0.5, -1])
def test_a_serial_that_is_no_date_in_the_1900_system_is_refused_by_name(
        tmp_path, route, serial):
    p = _linked_book(tmp_path, [("Linked day", "LinkedRate", "date", "2020-05-17T00:00:00Z")],
                     value=45000, fmt="yyyy-mm-dd")
    snapshot = p.read_bytes()
    with pytest.raises(HazardRefused) as exc:
        cells.set_cell(str(p), "B1", serial, sheet="Data", backup=False)
    assert "Linked day" in str(exc.value) and "Data!B1" in str(exc.value)
    assert p.read_bytes() == snapshot


# ------------------------------------------------ formula text, spaces and all

def _formula_book(tmp_path: Path, formula: str) -> Path:
    return _linked_book(tmp_path, [("Linked phrase", "LinkedRate", "lpwstr", "ab")],
                        value=formula)


def test_a_space_inside_a_string_literal_is_a_changed_formula(tmp_path, route):
    p = _formula_book(tmp_path, '="ab"')
    snapshot = p.read_bytes()
    with pytest.raises(HazardRefused) as exc:
        formulas.set_formula(str(p), "B1", '="a b"', sheet="Data", backup=False)
    assert "Linked phrase" in str(exc.value) and "Data!B1" in str(exc.value)
    assert p.read_bytes() == snapshot


def test_a_different_formula_is_still_refused(tmp_path, route):
    p = _formula_book(tmp_path, '="ab"')
    with pytest.raises(HazardRefused):
        formulas.set_formula(str(p), "B1", '="cd"', sheet="Data", backup=False)


def test_the_same_formula_written_again_is_not_a_change(tmp_path, route):
    p = _formula_book(tmp_path, '="ab"')
    res = formulas.set_formula(str(p), "B1", '="ab"', sheet="Data", backup=False)
    assert res["ok"] is True and "linked_properties" not in res["changed"]
    assert _prop(p, "Linked phrase") == ("lpwstr", "ab")


def test_an_edit_elsewhere_leaves_a_linked_formula_and_its_property_alone(tmp_path, route):
    p = _formula_book(tmp_path, '="a b"')
    before = _custom(p)
    res = cells.set_cell(str(p), "B3", 9, sheet="Data", backup=False)
    assert res["ok"] is True and _custom(p) == before
    assert "linked_properties" not in res["changed"]


def test_a_formula_with_extra_spaces_between_tokens_is_still_a_change(tmp_path, route):
    p = _formula_book(tmp_path, "=A1&A3")
    with pytest.raises(HazardRefused):
        formulas.set_formula(str(p), "B1", "=A1 & A3", sheet="Data", backup=False)


# -------------------------------------------- a value outside the type's range

@pytest.mark.parametrize("tag,value", [("r4", 1e40), ("r4", -1e40), ("r4", 1e-50),
                                       ("cy", 1e16), ("decimal", 1e30)])
def test_a_number_outside_the_property_types_range_is_refused_by_name(
        tmp_path, route, tag, value):
    p = _linked_book(tmp_path, [("Linked amount", "LinkedRate", tag, "1")], value=1)
    snapshot = p.read_bytes()
    with pytest.raises(HazardRefused) as exc:
        cells.set_cell(str(p), "B1", value, sheet="Data", backup=False)
    assert "Linked amount" in str(exc.value)
    assert tag in str(exc.value)
    assert p.read_bytes() == snapshot


def test_allow_loss_leaves_an_out_of_range_property_and_names_it(tmp_path, route):
    p = _linked_book(tmp_path, [("Linked amount", "LinkedRate", "r4", "1")], value=1)
    res = cells.set_cell(str(p), "B1", 1e40, sheet="Data", allow_loss=True, backup=False)
    assert res["ok"] is True
    assert _prop(p, "Linked amount") == ("r4", "1")
    assert any("Linked amount" in w and "Data!B1" in w for w in res["warnings"])


# The cy and decimal limits are exact numbers that a 64-bit float cannot hold:
# the nearest float to each limit lies just past it, so a value equal to that
# float is outside the type although it equals the "rounded" limit.
CY_ABOVE = 922337203685477.6      # as a float: 922337203685477.625 (limit ...477.5807)
CY_BELOW = -922337203685477.6     # as a float: -922337203685477.625 (limit ...477.5808)
DEC_ABOVE = 7.922816251426434e28  # as a float: 2**96 (limit 2**96 - 1)
DEC_BELOW = -7.922816251426434e28
DEC_LAST_INSIDE = 7.922816251426433e28  # the float before 2**96


@pytest.mark.parametrize("tag,value", [("cy", CY_ABOVE), ("cy", CY_BELOW),
                                       ("decimal", DEC_ABOVE), ("decimal", DEC_BELOW),
                                       ("i8", 2.0**63),
                                       ("i4", 2.0**31), ("ui4", 2.0**32),
                                       ("i4", -2.0**31 - 1), ("ui8", -1.0)])
def test_a_value_just_past_an_exact_limit_is_refused_by_name(tmp_path, route, tag, value):
    p = _linked_book(tmp_path, [("Linked amount", "LinkedRate", tag, "1")], value=1)
    snapshot = p.read_bytes()
    with pytest.raises(HazardRefused) as exc:
        cells.set_cell(str(p), "B1", value, sheet="Data", backup=False)
    assert "Linked amount" in str(exc.value)
    assert tag in str(exc.value)
    assert p.read_bytes() == snapshot


@pytest.mark.parametrize("tag,value", [("cy", CY_ABOVE), ("cy", CY_BELOW),
                                       ("decimal", DEC_ABOVE), ("decimal", DEC_BELOW),
                                       ("i8", 2.0**63), ("ui8", -1.0)])
def test_allow_loss_leaves_the_property_when_the_value_is_just_past_a_limit(
        tmp_path, route, tag, value):
    p = _linked_book(tmp_path, [("Linked amount", "LinkedRate", tag, "1")], value=1)
    res = cells.set_cell(str(p), "B1", value, sheet="Data", allow_loss=True, backup=False)
    assert res["ok"] is True
    assert _prop(p, "Linked amount") == (tag, "1")
    assert any("Linked amount" in w and "Data!B1" in w and "left as it was" in w
               for w in res["warnings"])


@pytest.mark.parametrize("tag,value", [
    ("cy", 922337203685477.5), ("cy", -922337203685477.5),
    ("decimal", DEC_LAST_INSIDE), ("decimal", -DEC_LAST_INSIDE),
    ("i8", -2.0**63), ("i8", 2.0**62), ("ui8", 2.0**63), ("ui4", 4294967295.0),
    ("i4", -2147483648.0)])
def test_a_value_at_the_edge_of_an_exact_limit_is_written(tmp_path, route, tag, value):
    p = _linked_book(tmp_path, [("Linked amount", "LinkedRate", tag, "1")], value=2)
    cells.set_cell(str(p), "B1", value, sheet="Data", backup=False)
    got = _prop(p, "Linked amount")
    assert got[0] == tag and float(got[1]) == value


@pytest.mark.parametrize("tag,value", [
    ("r4", 3.0e38), ("r4", 0.0), ("r4", 1.5), ("r8", 1e300), ("cy", 12345.5),
    ("i4", 2147483647)])
def test_a_number_inside_the_range_is_written(tmp_path, route, tag, value):
    p = _linked_book(tmp_path, [("Linked amount", "LinkedRate", tag, "1")], value=2)
    cells.set_cell(str(p), "B1", value, sheet="Data", backup=False)
    got = _prop(p, "Linked amount")
    assert got[0] == tag and float(got[1]) == value


# ------------------------------------ shared formulas in other XML spellings

def _spelled_sheet(tmp_path: Path, spelling: str) -> Path:
    p = _sheet_book(tmp_path, "sp_" + spelling, "nomaster")
    member = "xl/worksheets/sheet1.xml"
    if spelling == "utf16le":
        _recode(p, member, lambda d: _utf16(d, "le"))
    elif spelling == "utf16be":
        _recode(p, member, lambda d: _utf16(d, "be"))
    elif spelling == "charref":
        _recode(p, member, lambda d: d.replace(b't="shared"', b't="shar&#101;d"'))
    elif spelling == "charref-si":
        _recode(p, member, lambda d: d.replace(b'si="0"', b'si="&#48;"'))
    elif spelling == "comment-master":
        _recode(p, member, lambda d: d.replace(
            b"<sheetData>",
            b'<sheetData><!-- <f t="shared" ref="B1:B5" si="0">A1*2</f> -->'))
    elif spelling == "pi-master":
        _recode(p, member, lambda d: d.replace(
            b"<sheetData>", b'<sheetData><?x <f t="shared" si="0">A1*2</f> ?>'))
    else:
        assert spelling == "plain"
    return p


SHEET_SPELLINGS = ["plain", "utf16le", "utf16be", "charref", "charref-si",
                   "comment-master", "pi-master"]


@pytest.mark.parametrize("spelling", SHEET_SPELLINGS)
def test_a_member_with_no_master_is_flagged_however_the_xml_is_spelled(tmp_path, spelling):
    p = _spelled_sheet(tmp_path, spelling)
    r = cells.read_range(str(p), "B2:B5", sheet="Shared", values="formula")
    assert [row[0] for row in r["values"]] == [4, 6, 8, 10]
    assert r["broken_shared_formulas"]["cells"] == ["B2", "B3", "B4", "B5"]
    rep = _hazard.scan_path(str(p))
    assert not rep.clean
    assert [h.key for h in rep.hazards] == [_hazard.BROKEN_SHARED_KEY]


@pytest.mark.parametrize("spelling", SHEET_SPELLINGS)
def test_an_edit_elsewhere_names_the_dropped_shared_formulas_in_every_spelling(
        tmp_path, route, spelling):
    p = _spelled_sheet(tmp_path, spelling)
    res = cells.set_cell(str(p), "A1", 10, sheet="Shared", backup=False)
    assert res["ok"] is True
    got = openpyxl.load_workbook(p, data_only=True)["Shared"]
    assert [got.cell(i, 2).value for i in range(2, 6)] == [4, 6, 8, 10]
    if route == "rewrite":
        assert any("shared formula" in w for w in res["warnings"])


# ------------------------------- a date property is a UTC instant, as in Excel

@pytest.mark.parametrize("serial,expected", [
    (1, "1899-12-31T15:00:00Z"), (2, "1900-01-01T15:00:00Z"),
    (59, "1900-02-27T15:00:00Z"), (61, "1900-02-28T15:00:00Z"),
    (46159, "2026-05-16T15:00:00Z")])
def test_a_date_property_holds_the_utc_instant_excel_writes_for_the_cell(
        tmp_path, route, monkeypatch, serial, expected):
    """The five values are what Excel itself wrote for a date property linked to
    a cell holding these serials, on a computer set to UTC+9."""
    monkeypatch.setattr(_linkprops(), "_utc_offset", lambda when: datetime.timedelta(hours=9))
    p = _linked_book(tmp_path, [("Linked day", "LinkedRate", "filetime", "2020-05-17T00:00:00Z")],
                     value=45000, fmt="yyyy-mm-dd")
    cells.set_cell(str(p), "B1", serial, sheet="Data", backup=False)
    assert _prop(p, "Linked day") == ("filetime", expected)


def test_a_zone_west_of_utc_moves_the_instant_the_other_way(tmp_path, route, monkeypatch):
    monkeypatch.setattr(_linkprops(), "_utc_offset", lambda when: datetime.timedelta(hours=-8))
    p = _linked_book(tmp_path, [("Linked day", "LinkedRate", "date", "2020-05-17T00:00:00Z")],
                     value=45000, fmt="yyyy-mm-dd")
    cells.set_cell(str(p), "B1", 46159.5, sheet="Data", backup=False)
    assert _prop(p, "Linked day") == ("date", "2026-05-17T20:00:00Z")


def test_the_offset_is_the_computers_zone_in_force_now():
    for when in (datetime.datetime(1900, 1, 1), datetime.datetime(2026, 7, 15, 12),
                 datetime.datetime(2024, 2, 29)):
        off = _ZONE["real"](when)
        assert datetime.timedelta(hours=-14) <= off <= datetime.timedelta(hours=14)
        now = datetime.datetime.now().year
        same = when.replace(year=now, day=28 if (when.month, when.day) == (2, 29) else when.day)
        assert off == same.astimezone().utcoffset()
