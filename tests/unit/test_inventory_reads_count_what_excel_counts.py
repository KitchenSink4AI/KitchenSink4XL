"""Inventory reads count what Excel counts.

get_workbook_metadata left chart sheets out of the sheets and the count;
manage_name list left out Excel's built-in names (print area, print titles,
the filter range), which openpyxl folds into sheet settings on load; and
manage_conditional_format list left out the rules stored only in Excel's
extension block, answering 14 where Excel has 16. All content is made up.
"""

from __future__ import annotations

import re
import zipfile

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

from xlsx_mcp.ops import condformat, lifecycle, names

X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"


def _book(tmp_path):
    p = tmp_path / "inventory.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for i, (region, units) in enumerate([("North", 4), ("South", 7),
                                         ("East", 2)], start=2):
        ws.cell(i, 1, region)
        ws.cell(i, 2, units)
    ws["A1"], ws["B1"] = "Region", "Units"
    ws.print_area = "A1:B4"
    ws.print_title_rows = "1:1"
    ws.auto_filter.ref = "A1:B4"
    ws.conditional_formatting.add(
        "B2:B4", CellIsRule(operator="greaterThan", formula=["5"],
                            fill=PatternFill("solid", start_color="FFFF00")))
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=4),
                   titles_from_data=True)
    cs = wb.create_chartsheet("Units chart")
    cs.add_chart(chart)
    wb.save(p)
    return str(p)


def _add_extension_only_rule(p: str) -> None:
    """One rule stored only in the x14 block, as Excel stores a rule that
    refers to another sheet."""
    ext = (f'<extLst><ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}" '
           f'xmlns:x14="{X14_NS}"><x14:conditionalFormattings>'
           f'<x14:conditionalFormatting xmlns:xm="{XM_NS}">'
           '<x14:cfRule type="expression" priority="2" '
           'id="{00000000-0000-0000-0000-00000000AB01}">'
           '<xm:f>Data!$B2&gt;Data!$B$4</xm:f><x14:dxf/></x14:cfRule>'
           '<xm:sqref>A2:A4</xm:sqref></x14:conditionalFormatting>'
           '</x14:conditionalFormattings></ext></extLst>')
    with zipfile.ZipFile(p) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == "xl/worksheets/sheet1.xml":
                x = data.decode("utf-8")
                x = re.sub(r"</worksheet>\s*$", ext + "</worksheet>", x)
                data = x.encode("utf-8")
            zout.writestr(info, data)


def test_metadata_lists_the_chart_sheet_in_tab_order(tmp_path):
    out = lifecycle.get_workbook_metadata(_book(tmp_path))
    assert out["sheet_count"] == 2
    assert [s["name"] for s in out["sheets"]] == ["Data", "Units chart"]
    chart = out["sheets"][1]
    assert "chart" in chart["kind"]
    assert chart["used_range"] is None


def test_list_names_includes_excels_built_in_names(tmp_path):
    out = names.manage_name(_book(tmp_path), "list")
    got = {(n["name"], n["scope"]) for n in out["names"]}
    for builtin in ("_xlnm.Print_Area", "_xlnm.Print_Titles",
                    "_xlnm._FilterDatabase"):
        assert (builtin, "Data") in got, out["names"]
    assert all(n["reserved"] for n in out["names"]
               if n["name"].startswith("_xlnm."))
    assert out["count"] == len(out["names"])


def test_list_cf_counts_rules_stored_only_in_the_extension_block(tmp_path):
    p = _book(tmp_path)
    _add_extension_only_rule(p)
    out = condformat.manage_conditional_format(p, "list")
    assert out["count"] == 1                 # the rules listed in detail
    assert out["total_rules"] == 2           # what Excel counts
    assert out["extension_only_rules"] == {"Data": 1}
    assert "extension" in out["note"]
