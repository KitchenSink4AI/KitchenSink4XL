"""Four ways an edit could leave a workbook wrong while the answer said ok
(every workbook here is made up and built by the test).

A legacy query table: a rewrite with allow_loss dropped the query-table part
and its relationships but kept tableType="queryTable" and the column ids on
the table, so the table pointed at a part that no longer existed. Excel could
not open the file and the save said verified. The rewrite now turns the table
into an ordinary table (its cells stay) and says so; the verify step refuses
any package whose query table lacks its part.

Content under a merged range: a rewrite silently dropped what is stored in a
cell of the range other than the top-left one (a value, a note or a
hyperlink), and moved a hyperlink from such a cell to the top-left one. It is
now a named loss: the edit is refused unless allow_loss accepts it. The
top-left cell itself moves with its column, and a hyperlink moves with its
cell on every row or column insert and delete.

Strict Open XML: a Strict workbook read as ok with no sheets. Every read and
edit now refuses it by name.
"""
from __future__ import annotations

import re
import zipfile

import openpyxl
import pytest

from xlsx_mcp.core import hazard, verify

REL_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
REL_QT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"
CT_QT = "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"

QT_PART = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'name="ExternalData_1" connectionId="1" autoFormatId="16" applyNumberFormats="0" '
    b'applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" '
    b'applyAlignmentFormats="0" applyWidthHeightFormats="0"><queryTableRefresh nextId="4">'
    b'<queryTableFields count="3"><queryTableField id="1" name="Site" tableColumnId="1"/>'
    b'<queryTableField id="2" name="Pallets" tableColumnId="2"/>'
    b'<queryTableField id="3" name="Shipped" tableColumnId="3"/></queryTableFields>'
    b'</queryTableRefresh></queryTable>')
CONNECTIONS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b'<connection id="1" name="Orders feed" type="5" refreshedVersion="8" background="1" '
    b'saveData="1"><dbPr connection="Provider=SampleProvider;Data Source=made_up.accdb" '
    b'command="SELECT 1" commandType="2"/></connection></connections>')


def _query_table_workbook(tmp_path, name="qt.xlsx", *, keep_part=True):
    """A made-up workbook: one table on Feeds!A3:C6 declared as a query table."""
    base = str(tmp_path / ("base_" + name))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Feeds"
    ws["A1"] = "Orders feed"
    ws.append([])
    ws.append(["Site", "Pallets", "Shipped"])
    ws.append(["North", 12, 10])
    ws.append(["South", 7, 7])
    ws.append(["East", 3, 1])
    from openpyxl.worksheet.table import Table
    ws.add_table(Table(displayName="Orders_feed", ref="A3:C6"))
    wb.save(base)
    out = str(tmp_path / name)
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/tables/table1.xml":
                text = data.decode("utf-8")
                text = re.sub(r'<table ', '<table tableType="queryTable" ', text, count=1)
                for i in (1, 2, 3):
                    text = text.replace(f'<tableColumn id="{i}" ',
                                        f'<tableColumn id="{i}" uniqueName="{i}" queryTableFieldId="{i}" ')
                data = text.encode("utf-8")
            if info.filename == "[Content_Types].xml" and keep_part:
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/queryTables/queryTable1.xml" ContentType="'
                    + CT_QT.encode() + b'"/>'
                    b'<Override PartName="/xl/connections.xml" ContentType="application/vnd.'
                    b'openxmlformats-officedocument.spreadsheetml.connections+xml"/></Types>')
            if info.filename == "xl/_rels/workbook.xml.rels" and keep_part:
                data = data.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/'
                    b'officeDocument/2006/relationships/connections" Target="connections.xml"/>'
                    b"</Relationships>")
            zout.writestr(info, data)
        if keep_part:
            zout.writestr("xl/queryTables/queryTable1.xml", QT_PART)
            zout.writestr("xl/connections.xml", CONNECTIONS)
            zout.writestr(
                "xl/tables/_rels/table1.xml.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                f'relationships"><Relationship Id="rId1" Type="{REL_QT}" '
                'Target="../queryTables/queryTable1.xml"/></Relationships>')
    return out


def _table_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("xl/tables/table1.xml").decode("utf-8")


# ------------------------------------------------------------ query tables

def test_the_fixture_is_a_clean_query_table_workbook(tmp_path):
    p = _query_table_workbook(tmp_path)
    assert verify.package_refusals(p) == []
    assert 'tableType="queryTable"' in _table_xml(p)


def test_verify_refuses_a_query_table_without_its_part(tmp_path):
    p = _query_table_workbook(tmp_path, keep_part=False)
    why = verify.package_refusals(p)
    assert why and "query table" in why[0].lower()
    assert "Orders_feed" in why[0]


def test_a_rewrite_with_allow_loss_leaves_an_ordinary_table_and_says_so(tmp_path):
    from xlsx_mcp.ops import lifecycle
    p = _query_table_workbook(tmp_path)
    out = lifecycle.manage_worksheet(p, "rename", sheet="Feeds",
                                     new_name="Feeds2", allow_loss=True)
    assert out.get("ok") is not False
    table = _table_xml(p)
    assert "queryTable" not in table
    assert "queryTableFieldId" not in table and "uniqueName" not in table
    with zipfile.ZipFile(p) as z:
        assert not any("querytable" in n.lower() for n in z.namelist())
    assert verify.package_refusals(p) == []
    wb = openpyxl.load_workbook(p)
    assert wb["Feeds2"]["A4"].value == "North"
    assert "Orders_feed" in wb["Feeds2"].tables
    said = " ".join(out.get("warnings", []))
    assert "query table" in said.lower() and "ordinary table" in said.lower()


def test_a_rewrite_without_allow_loss_still_refuses_a_query_table(tmp_path):
    from xlsx_mcp.ops import lifecycle
    p = _query_table_workbook(tmp_path)
    with pytest.raises(Exception) as exc:
        lifecycle.manage_worksheet(p, "rename", sheet="Feeds", new_name="Feeds2")
    assert "query tables" in str(exc.value).lower()
    assert 'tableType="queryTable"' in _table_xml(p)


# ------------------------------------------------- content under a merge

def _merge_workbook(tmp_path, *, hidden_value=True, hyperlink_on=None,
                    note_on=None, name="m.xlsx"):
    """A1:B1 merged. B1 (not the top-left cell) may hold a value, a link or a note."""
    base = str(tmp_path / ("base_" + name))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Shown"
    ws["A2"] = 1
    ws["B2"] = 2
    ws["D1"] = 5
    wb.save(base)
    out = str(tmp_path / name)
    sheet_cells = ""
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                if hidden_value:
                    text = text.replace(
                        '<c r="D1"', '<c r="B1" t="inlineStr"><is><t>Hidden under the merge</t></is></c><c r="D1"', 1)
                extra = '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'
                if hyperlink_on:
                    extra += (f'<hyperlinks><hyperlink ref="{hyperlink_on}" '
                              'location="Data!A2" display="x"/></hyperlinks>')
                text = text.replace("</sheetData>", "</sheetData>" + extra, 1)
                data = text.encode("utf-8")
            zout.writestr(info, data)
    return out


def _insert_cols(p, **kw):
    from xlsx_mcp.ops import structure
    return structure.modify_grid_structure(p, "insert_cols", 3, count=1, sheet="Data", **kw)


def test_hidden_value_under_a_merge_is_a_named_loss(tmp_path):
    p = _merge_workbook(tmp_path)
    rep = hazard.scan_path(p)
    assert hazard.MERGED_HIDDEN_KEY in rep.lossy_keys
    with pytest.raises(Exception) as exc:
        _insert_cols(p)
    msg = str(exc.value).lower()
    assert "merged" in msg and "allow_loss" in msg
    # the file is untouched
    with zipfile.ZipFile(p) as z:
        assert b"Hidden under the merge" in z.read("xl/worksheets/sheet1.xml")


def test_allow_loss_accepts_it_and_says_what_goes(tmp_path):
    p = _merge_workbook(tmp_path)
    out = _insert_cols(p, allow_loss=True)
    assert out.get("ok") is not False
    said = " ".join(out.get("warnings", [])).lower()
    assert "merged" in said


def test_hidden_hyperlink_under_a_merge_is_a_named_loss(tmp_path):
    p = _merge_workbook(tmp_path, hidden_value=False, hyperlink_on="B1")
    assert hazard.MERGED_HIDDEN_KEY in hazard.scan_path(p).lossy_keys
    with pytest.raises(Exception):
        _insert_cols(p)


def test_hidden_note_under_a_merge_is_a_named_loss(tmp_path):
    p = _merge_workbook(tmp_path, hidden_value=False)
    # add a legacy note on B1 through the file, the way Excel stores it
    out = str(tmp_path / "note.xlsx")
    with zipfile.ZipFile(p) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(
                    b"</worksheet>",
                    b'<legacyDrawing r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/'
                    b'officeDocument/2006/relationships"/></worksheet>')
            if info.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/comments1.xml" ContentType="application/vnd.'
                    b'openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>')
            zout.writestr(info, data)
        zout.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/comments" Target="../comments1.xml"/></Relationships>')
        zout.writestr(
            "xl/comments1.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<authors><author>A</author></authors><commentList><comment ref="B1" authorId="0">'
            '<text><r><t>made-up note</t></r></text></comment></commentList></comments>')
    assert hazard.MERGED_HIDDEN_KEY in hazard.scan_path(out).lossy_keys


@pytest.mark.parametrize("kw", [
    {"hidden_value": False},                                   # nothing under the merge
    {"hidden_value": False, "hyperlink_on": "A1"},             # link on the top-left cell
])
def test_a_merge_with_nothing_hidden_is_no_hazard(tmp_path, kw):
    p = _merge_workbook(tmp_path, **kw)
    assert hazard.MERGED_HIDDEN_KEY not in hazard.scan_path(p).lossy_keys
    out = _insert_cols(p)
    assert out.get("ok") is not False


def test_a_hyperlink_on_the_top_left_cell_stays_put(tmp_path):
    p = _merge_workbook(tmp_path, hidden_value=False, hyperlink_on="A1")
    _insert_cols(p)
    wb = openpyxl.load_workbook(p)
    assert wb["Data"]["A1"].hyperlink is not None
    assert wb["Data"]["A1"].hyperlink.location == "Data!A2"


# ------------------------------------------------------- Strict Open XML

STRICT_MAIN = "http://purl.oclc.org/ooxml/spreadsheetml/main"
STRICT_REL = "http://purl.oclc.org/ooxml/officeDocument/relationships"
TRANS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
TRANS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _strict_workbook(tmp_path, name="strict.xlsx"):
    base = str(tmp_path / ("base_" + name))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Strict"
    ws["A1"] = "Total"
    ws["B1"] = 41
    wb.save(base)
    out = str(tmp_path / name)
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.endswith((".xml", ".rels")) and info.filename != "[Content_Types].xml":
                data = (data.replace(TRANS_MAIN.encode(), STRICT_MAIN.encode())
                        .replace(TRANS_REL.encode(), STRICT_REL.encode()))
            zout.writestr(info, data)
    return out


@pytest.fixture()
def strict_path(tmp_path):
    return _strict_workbook(tmp_path)


def test_the_strict_sniff(strict_path, tmp_path):
    assert hazard.is_strict_package(strict_path) is True
    plain = str(tmp_path / "plain.xlsx")
    openpyxl.Workbook().save(plain)
    assert hazard.is_strict_package(plain) is False


def test_every_read_and_edit_refuses_strict_by_name(strict_path):
    from xlsx_mcp.core.errors import UnsupportedStructure
    from xlsx_mcp.core.package import WorkbookPackage
    from xlsx_mcp.ops import cells, gridio, lifecycle
    calls = [
        lambda: WorkbookPackage.open(strict_path),
        lambda: gridio.open_wb(strict_path),
        lambda: cells.read_range(strict_path, "A1:B1", sheet="Strict"),
        lambda: cells.set_cell(strict_path, "A1", "x", sheet="Strict"),
        lambda: lifecycle.manage_worksheet(strict_path, "rename", sheet="Strict",
                                           new_name="Other"),
    ]
    for call in calls:
        with pytest.raises(UnsupportedStructure) as exc:
            call()
        assert "Strict Open XML" in str(exc.value)
        assert "Nothing was changed" in str(exc.value)


def test_validate_reports_strict_instead_of_no_visible_sheet(strict_path):
    from xlsx_mcp.ops import validation
    res = validation._check_structure(strict_path)
    text = str(res)
    assert "Strict Open XML" in text
    assert "no visible sheet" not in text


# ------------------------------------------------- links and merged anchors

def _links_workbook(tmp_path):
    p = str(tmp_path / "links.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Links"
    for r in range(1, 6):
        ws.cell(r, 1, f"row {r}")
        c = ws.cell(r, 2, f"site {r}")
        c.hyperlink = f"https://example.invalid/page{r}"
    wb.save(p)
    return p


def _link_map(p):
    out = {}
    ws = openpyxl.load_workbook(p)["Links"]
    for row in ws.iter_rows():
        for c in row:
            if c.hyperlink is not None:
                out[c.coordinate] = (c.value, c.hyperlink.target)
    return out


@pytest.mark.parametrize("action, at, shift", [
    ("insert_rows", 3, (1, 0)), ("delete_rows", 2, (-1, 0)),
    ("insert_cols", 2, (0, 1)), ("insert_cols", 1, (0, 1))])
def test_a_hyperlink_follows_its_cell_when_rows_or_columns_move(tmp_path, action, at, shift):
    from xlsx_mcp.ops import structure
    p = _links_workbook(tmp_path)
    before = _link_map(p)
    structure.modify_grid_structure(p, action, at, count=1, sheet="Links")
    after = _link_map(p)
    ws = openpyxl.load_workbook(p)["Links"]
    # every surviving link sits on the cell that carries its text
    for coord, (value, target) in after.items():
        assert target.endswith("page" + str(value).split()[-1]), (coord, value, target)
    # and every link whose cell survived is still on that cell's text
    texts = {v: t for v, t in before.values()}
    for row in ws.iter_rows():
        for c in row:
            if c.value in texts and str(c.value).startswith("site"):
                assert c.hyperlink is not None and c.hyperlink.target == texts[c.value], c.coordinate


def _anchor_workbook(tmp_path):
    """D3:E4 merged; D3 holds text and a link. Rows 5-6 hold ordinary content."""
    p = str(tmp_path / "anchor.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "head"
    ws["D3"] = "Merged link"
    ws["D3"].hyperlink = "https://example.invalid/merged"
    ws.merge_cells("D3:E4")
    ws["D5"] = "below one"
    ws["E5"] = "below two"
    ws["D6"] = "below three"
    wb.save(p)
    return p


def test_the_top_left_cell_of_a_merge_moves_with_its_column(tmp_path):
    from xlsx_mcp.ops import structure
    p = _anchor_workbook(tmp_path)
    structure.modify_grid_structure(p, "insert_cols", 2, count=1, sheet="Data")
    ws = openpyxl.load_workbook(p)["Data"]
    assert [str(r) for r in ws.merged_cells.ranges] == ["E3:F4"]
    assert ws["E3"].value == "Merged link"
    assert ws["E3"].hyperlink is not None
    assert ws["E3"].hyperlink.target == "https://example.invalid/merged"
    assert ws["E5"].value == "below one" and ws["F5"].value == "below two"


def test_deleting_a_merged_range_does_not_eat_the_rows_that_move_up(tmp_path):
    from xlsx_mcp.ops import structure
    p = _anchor_workbook(tmp_path)
    structure.modify_grid_structure(p, "delete_rows", 3, count=2, sheet="Data")
    ws = openpyxl.load_workbook(p)["Data"]
    assert list(ws.merged_cells.ranges) == []
    assert ws["D3"].value == "below one"
    assert ws["E3"].value == "below two"
    assert ws["D4"].value == "below three"
