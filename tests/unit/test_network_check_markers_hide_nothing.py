"""The network check cannot be hidden behind markup that looks like a comment.

The check reads a file's XML as the parser hands it to Excel (comments and processing instructions gone, CDATA
sections unwrapped, character references resolved). A text that merely LOOKS like the start of a comment, such as
"<!--" inside a CDATA section, a processing instruction, an attribute value or a CSV cell, must not join a later
"-->" and delete everything between them, a network call included. So a file counts as calling a network function
when the call is found in the raw text OR in any normalized reading of it: normalizing only adds places to look.
Measured live in Excel 16 on made-up files with a listener on 127.0.0.1: an .xlsx, a CSV and an .ods built this
way fetched when the check missed them.

These tests build the files by code and read them with the file check alone (no Excel).
"""
from __future__ import annotations

import re
import zipfile

import openpyxl
import pytest

from xlsx_mcp.core import netreach
from xlsx_mcp.core.netreach import file_reach

URL = "http://127.0.0.1:9/made-up"
CALL = f'_xlfn.WEBSERVICE("{URL}")'


def _xlsx(tmp_path, cells, name="book.xlsx"):
    p = tmp_path / "base.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.active["A1"] = 1
    wb.save(p)
    with zipfile.ZipFile(p) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    out = tmp_path / name
    with zipfile.ZipFile(out, "w") as z:
        for i, d in items:
            if i.filename == "xl/worksheets/sheet1.xml":
                d = re.sub(r"<sheetData>.*</sheetData>|<sheetData\s*/>",
                           lambda m: "<sheetData>" + cells + "</sheetData>", d.decode(), flags=re.S).encode()
            z.writestr(i, d)
    return str(out)


def _ods(tmp_path, inner, name="rates.ods"):
    content = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
        'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">'
        '<office:body><office:spreadsheet>' + inner + '</office:spreadsheet></office:body></office:document-content>')
    p = tmp_path / name
    with zipfile.ZipFile(p, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("content.xml", content)
    return str(p)


ODS_TABLE = ('<table:table table:name="Rates"><table:table-row>'
             f'<table:table-cell table:formula="of:=COM.MICROSOFT.WEBSERVICE(&quot;{URL}&quot;)" '
             'office:value-type="string"><text:p></text:p></table:table-cell>'
             '</table:table-row></table:table>')


def _sees(reach, fn="WEBSERVICE"):
    return bool(reach) and any(fn in f["functions"] for f in reach["formulas"])


def _between(left, right, call=CALL):
    """The call in a cell between two cells holding text that looks like the two ends of a comment."""
    return ('<row r="1"><c r="A1" t="inlineStr"><is><t>' + left + '</t></is></c>'
            f'<c r="B1"><f>{call}</f></c>'
            '<c r="C1" t="inlineStr"><is><t>' + right + '</t></is></c></row>')


@pytest.mark.parametrize("cells", [
    pytest.param(_between("<![CDATA[<!--]]>", "<![CDATA[-->]]>"), id="cdata-sections"),
    pytest.param('<row r="1"><c r="A1"><v>1</v></c></row><?x <!--?>'
                 f'<row r="2"><c r="B2"><f>{CALL}</f></c></row><?x -->?>', id="processing-instructions"),
    pytest.param('<row r="1"><c r="A1" s="<!--"><v>1</v></c>'
                 f'<c r="B1"><f>{CALL}</f></c><c r="C1" s="-->"><v>2</v></c></row>', id="attribute-values"),
    pytest.param('<row r="1"><c r="A1" s="<!--"><v>1</v></c>'
                 f'<c r="B1"><f>{CALL}</f></c></row><!-- a real comment -->', id="attribute-value-then-comment"),
    pytest.param(_between("<!-- <!-- -->", "-->"), id="nested-comment-openers"),
    pytest.param(_between("<![CDATA[<![CDATA[]]>", "]]>"), id="nested-cdata-openers"),
    pytest.param('<row r="1"><c r="A1"><v>1</v></c></row><!-- never closed '
                 f'<row r="2"><c r="B2"><f>{CALL}</f></c></row>', id="comment-never-closed"),
    pytest.param(f'<row r="1"><c r="B1"><f>{CALL}</f></c></row><![CDATA[ never closed', id="cdata-never-closed"),
    pytest.param(f'<row r="1"><c r="B1"><f>{CALL}</f></c></row><?x never closed', id="instruction-never-closed"),
    pytest.param('<row r="1"><c r="B1"><f>_xlfn.WEB<!-- x -->SERVICE'
                 f'("{URL}")</f></c></row>', id="split-by-a-comment"),
    pytest.param(f'<row r="1"><c r="B1"><f>_xlfn.WEB<![CDATA[SERVICE]]>("{URL}")</f></c></row>',
                 id="split-by-cdata"),
    pytest.param(f'<row r="1"><c r="B1"><f>_xlfn.WEB<?x y?>SERVICE("{URL}")</f></c></row>',
                 id="split-by-an-instruction"),
    pytest.param(f'<row r="1"><c r="B1"><f>_xlfn.WEB&#83;ERVICE("{URL}")</f></c></row>', id="character-reference"),
])
def test_xlsx_call_is_seen_however_the_markup_hides_it(tmp_path, cells):
    how, reach = file_reach(_xlsx(tmp_path, cells))
    assert how == "package"
    assert _sees(reach), reach


def test_xlsx_with_no_call_between_such_markers_is_not_refused(tmp_path):
    how, reach = file_reach(_xlsx(tmp_path, _between("<![CDATA[<!--]]>", "<![CDATA[-->]]>", call="SUM(1,2)")))
    assert reach is None


def test_csv_cells_that_look_like_comment_markers_hide_nothing(tmp_path):
    p = tmp_path / "rates.csv"
    p.write_text(f'<!--,=WEBSERVICE("{URL}"),-->\n', encoding="utf-8")
    how, reach = file_reach(str(p))
    assert how == "text"
    assert _sees(reach), reach


@pytest.mark.parametrize("text", [
    f'a,<!--\n=WEBSERVICE("{URL}"),-->,b\n',
    f'<!--,<!--,-->\n=WEBSERVICE("{URL}")\n',
    f'<![CDATA[,<?x,-->\n=WEBSERVICE("{URL}")\n',
    f'x,=WEBSERVICE("{URL}")\n<!--',
])
def test_csv_markers_in_any_order_hide_nothing(tmp_path, text):
    p = tmp_path / "rates.csv"
    p.write_text(text, encoding="utf-8")
    how, reach = file_reach(str(p))
    assert _sees(reach), reach


def test_a_csv_is_never_read_as_xml(tmp_path):
    """A name split by something that looks like a comment is only text in a CSV; nothing is stripped."""
    p = tmp_path / "notes.csv"
    p.write_text('<!--,WEB<!-- x -->SERVICE(1),-->\n', encoding="utf-8")
    how, reach = file_reach(str(p))
    assert reach is None
    assert netreach._text_calls(["WEB<!-- x -->SERVICE(1)"], xml=False) == []


def test_text_that_starts_like_xml_is_read_as_xml(tmp_path):
    p = tmp_path / "flat.fods"
    p.write_text('<?xml version="1.0"?><root><a>WEB<!-- x -->SERVICE(1)</a></root>', encoding="utf-8")
    how, reach = file_reach(str(p))
    assert _sees(reach), reach


def test_ods_processing_instructions_holding_comment_markers_hide_nothing(tmp_path):
    how, reach = file_reach(_ods(tmp_path, '<?x <!--?>' + ODS_TABLE + '<?x -->?>'))
    assert how == "text"
    assert _sees(reach), reach


@pytest.mark.parametrize("inner", [
    '<![CDATA[<!--]]>' + ODS_TABLE + '<![CDATA[-->]]>',
    '<!-- <!-- -->' + ODS_TABLE + '-->',
    ODS_TABLE + '<!-- never closed',
])
def test_ods_call_is_seen_however_the_markup_hides_it(tmp_path, inner):
    how, reach = file_reach(_ods(tmp_path, inner))
    assert _sees(reach), reach


def test_ods_name_split_by_a_comment_inside_an_attribute_value_is_refused(tmp_path):
    # "<" inside an attribute value is not well-formed XML: no reading of it can be certain, so the check fails closed
    how, reach = file_reach(_ods(tmp_path, ODS_TABLE.replace("COM.MICROSOFT.WEB", "COM.MICROSOFT.WEB<!-- x -->")))
    assert reach and any("not well-formed" in u for u in reach["unread"]), reach


def test_the_chunked_text_check_agrees_with_the_whole():
    text = 'x<![CDATA[<!--]]>' + 'WEBSERVICE("u")' + '<![CDATA[-->]]>y'
    assert netreach._text_calls([text], xml=True) == ["WEBSERVICE"]
    assert netreach._text_calls([text[:20], text[20:]], xml=True) == ["WEBSERVICE"]


def test_views_only_add_places_to_look():
    """Every reading keeps the text itself first, so a call in the raw text is always found."""
    for text in ["a<!--b", "<![CDATA[x", "<?x y", "&#83;", "<!---->", "plain"]:
        for xml in (True, False):
            assert netreach._views(text, xml)[0] == text
