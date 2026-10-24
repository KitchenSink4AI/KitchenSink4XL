"""Save-path fidelity: cell edits keep everything else byte-for-byte.

Fixtures in tests/fixtures/preserve were authored and saved by Excel itself
(a private hidden instance, synthesised content, authoring identity scrubbed
by tests/fixtures/build_corpus.scrub_identity, and the printer-settings part
replaced by the made-up printer tests/fixtures/build_printer_settings.py
builds), plus one LibreOffice re-save.
Each carries something the openpyxl rewrite loses: pictures and printer
settings, x14 dropdowns / data bars / sparklines, a table with a slicer,
SharePoint-style customXml and sensitivity-label properties, legacy and
threaded comments, defined names and an external link, an Excel 4.0 macro
sheet, a pivot table, shared formulas, wrapped rows with automatic heights.

The Excel-level proof (each edited file opens in Excel without a repair,
renders identically outside the edited cells, keeps every untouched cached
result) was made against these same fixtures in Excel itself; these
tests pin the same properties at the package level so CI holds them.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import pytest

from xlsx_mcp.core import routes as _routes
from xlsx_mcp.core import surgical as _surgical
from xlsx_mcp.core import verify as _verify
from xlsx_mcp.core.errors import ExcelWouldRefuse, HazardRefused
from xlsx_mcp.core.package import WorkbookPackage
from xlsx_mcp.ops import (annotations, cells, datavalidation, dataio, formulas,
                          inspectors, lifecycle, objects, search, sortfilter,
                          structure, validation, view)
from xlsx_mcp.ops import format as fmt
from xlsx_mcp.surface.groups import GROUPS, READ_ONLY_GROUPS

FIX = Path(__file__).resolve().parents[1] / "fixtures" / "preserve"
CORPUS = Path(__file__).resolve().parents[1] / "fixtures" / "workbooks"


def _copy(tmp_path: Path, name: str, src: Path = FIX) -> str:
    dst = tmp_path / name
    shutil.copy2(src / name, dst)
    return str(dst)


def _parts(path) -> dict:
    with zipfile.ZipFile(path) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist()}


def _changed(before: dict, after: dict) -> tuple[set, set, set]:
    changed = {n for n in before if n in after and before[n] != after[n]}
    return changed, set(after) - set(before), set(before) - set(after)


def _sheet(path, part="xl/worksheets/sheet1.xml") -> str:
    return _parts(path)[part].decode("utf-8")


# ================================================================ the table

def test_every_mutating_op_has_a_route_class():
    # The update ops change the server program or its notice setting and
    # never save a workbook, so they have no save route (core/update_flow.py).
    mutating = {op for g, ops in GROUPS.items() if g not in READ_ONLY_GROUPS
                for op in ops} - {"update_now", "update_notices"}
    assert set(_routes.OP_EDIT_CLASS) == mutating


# ============================================== every cell op takes the splice

#: A representative call per "cell" op, on pictures_printer.xlsx (pictures,
#: printer settings): the parts the rewrite would lose must come through
#: byte-for-byte, and the route must be raw_ooxml.
S = "Report"
CELL_CALLS = {
    "set_cell": lambda p: cells.set_cell(p, {"cell": "A60", "sheet": S}, "x"),
    "set_cells": lambda p: cells.set_cells(p, [{"cell": "A60", "value": 1}], sheet=S),
    "write_range": lambda p: cells.write_range(p, {"cell": "A60", "sheet": S}, [[1, 2], [3, 4]]),
    "clear_range": lambda p: cells.clear_range(p, {"range": "A2:B3", "sheet": S}),
    "copy_range": lambda p: cells.copy_range(p, {"range": "A2:B3", "sheet": S},
                                             {"cell": "H2", "sheet": S}),
    "move_range": lambda p: cells.move_range(p, {"range": "A2:B3", "sheet": S},
                                             {"cell": "H20", "sheet": S}),
    "apply_edits": lambda p: view.apply_edits(
        p, [{"op": "set_value", "location": {"cell": "A60", "sheet": S}, "value": 5}]),
    "replace_cells": lambda p: search.replace_cells(p, "North", "Nord", sheet=S),
    "set_formula": lambda p: formulas.set_formula(p, {"cell": "F2", "sheet": S}, "=C2*2"),
    "import_data": lambda p: dataio.import_data(p, source="a,b\n1,2", fmt="csv",
                                                location={"cell": "H30", "sheet": S}),
    "format_cells": lambda p: fmt.format_cells(p, {"range": "A2:A3", "sheet": S},
                                               number_format="0.00"),
    "apply_style": lambda p: fmt.apply_style(p, "Normal", {"cell": "A2", "sheet": S}),
    "copy_format": lambda p: fmt.copy_format(p, {"cell": "A1", "sheet": S},
                                             {"cell": "A60", "sheet": S}),
    "manage_comment": lambda p: annotations.manage_comment(
        p, "add", {"cell": "A60", "sheet": S}, text="checked"),
    "manage_hyperlink": lambda p: annotations.manage_hyperlink(
        p, "add", {"cell": "A60", "sheet": S}, target="https://example.com/a?b=1&c=2"),
    "sort_range": lambda p: sortfilter.sort_range(
        p, {"range": "A1:E41", "sheet": S}, [{"column": "C", "order": "desc"}]),
}


def test_the_call_table_covers_every_cell_op():
    assert set(CELL_CALLS) == set(_routes.ops_in(_routes.CELL))


@pytest.mark.parametrize("op", sorted(CELL_CALLS))
def test_cell_op_takes_the_byte_preserving_route(tmp_path, op):
    p = _copy(tmp_path, "pictures_printer.xlsx")
    before = _parts(p)
    problems_before = _verify.package_problems(p)
    r = CELL_CALLS[op](p)
    assert r["ok"] and r["verified"]
    assert r["route"] == _surgical.ROUTE_RAW, r.get("route_note")
    after = _parts(p)
    changed, added, removed = _changed(before, after)
    keep = [n for n in before if n.startswith(("xl/media/", "xl/drawings/drawing",
                                               "xl/printerSettings/", "xl/theme/",
                                               "docProps/"))]
    assert keep, "the fixture has pictures and printer settings"
    for n in keep:
        assert after.get(n) == before[n], f"{op} changed {n}"
    allowed = {"xl/worksheets/sheet1.xml", "xl/sharedStrings.xml", "xl/styles.xml",
               "xl/workbook.xml", "[Content_Types].xml",
               "xl/worksheets/_rels/sheet1.xml.rels"}
    assert changed <= allowed, changed - allowed
    assert all(n.startswith(("xl/comments", "xl/drawings/vmlDrawing")) for n in added), added
    assert removed <= {"xl/calcChain.xml"}
    assert _verify.package_problems(p) <= problems_before


# =================================================== what stays as it was

def test_strings_stay_shared_and_rows_keep_their_heights(tmp_path):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    before = _sheet(p)
    rows_before = re.findall(r"<row [^>]*>", before)
    r = cells.set_cell(p, {"cell": "B70", "sheet": "Sheet1"}, "Appended text")
    assert r["route"] == "raw_ooxml"
    after = _sheet(p)
    assert "inlineStr" not in after
    assert 't="s"' in re.search(r'<c r="B70"[^>]*>', after).group(0)
    sst = _parts(p)["xl/sharedStrings.xml"].decode()
    assert "<t>Appended text</t>" in sst
    # every existing row element is byte-identical: no customHeight added
    assert re.findall(r"<row [^>]*>", after)[:len(rows_before)] == rows_before
    assert "customHeight" not in after.replace(before, "")


def test_app_properties_and_custom_xml_are_untouched(tmp_path):
    p = _copy(tmp_path, "customxml_label.xlsx")
    before = _parts(p)
    cells.set_cell(p, {"cell": "A40", "sheet": "Budget"}, 1)
    after = _parts(p)
    for n in before:
        if n.startswith(("customXml/", "docProps/")):
            assert after[n] == before[n], n


def test_full_precision_doubles_round_trip_exactly(tmp_path):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    stored = re.search(r'<c r="G3"[^>]*><v>([^<]+)</v>', _sheet(p)).group(1)
    r = cells.set_cell(p, {"cell": "G5", "sheet": "Sheet1"}, 0.30000000000000004)
    assert r["verified"]
    xml = _sheet(p)
    assert re.search(r'<c r="G3"[^>]*><v>([^<]+)</v>', xml).group(1) == stored
    assert "<v>0.30000000000000004</v>" in xml


# ================================================================ cached values

def _cached(path) -> dict:
    out = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*>(?:(?!</c>).)*?<f[^>]*?(?:/>|>(?:(?!</c>).)*?</f>)<v>([^<]*)</v>',
                         _sheet(path), re.S):
        out[m.group(1)] = m.group(2)
    return out


def test_format_only_edit_leaves_nothing_stale(tmp_path):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    before = _cached(p)
    wb_before = _parts(p)["xl/workbook.xml"]
    r = fmt.format_cells(p, {"range": "C2:C10", "sheet": "Sheet1"}, number_format="0.0")
    assert r["route"] == "raw_ooxml"
    assert "possibly_stale" not in r["cached_values"]
    assert _cached(p) == before
    assert _parts(p)["xl/workbook.xml"] == wb_before   # no recalc flag


def test_value_edit_reports_dependents_and_keeps_their_values(tmp_path):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    before = _cached(p)
    r = cells.set_cell(p, {"cell": "C5", "sheet": "Sheet1"}, 99)
    cv = r["cached_values"]
    assert set(cv["possibly_stale_cells"]) == {"Sheet1!E5", "Sheet1!E52"}
    assert cv["recalculates_on_open"] is True
    assert _cached(p) == before            # every <v> is still there
    assert 'fullCalcOnLoad="1"' in _parts(p)["xl/workbook.xml"].decode()


def test_edit_nobody_refers_to_leaves_nothing_stale(tmp_path):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    wb_before = _parts(p)["xl/workbook.xml"]
    r = cells.set_cell(p, {"cell": "K80", "sheet": "Sheet1"}, 5)
    assert "possibly_stale" not in r["cached_values"]
    assert _parts(p)["xl/workbook.xml"] == wb_before


def test_shared_formula_master_edit_keeps_every_dependent(tmp_path):
    p = _copy(tmp_path, "shared_formulas.xlsx")
    before = _cached(p)
    assert 't="shared"' in _sheet(p)
    r = formulas.set_formula(p, {"cell": "C3", "sheet": "Calc"}, "=B3*3")
    assert r["route"] == "raw_ooxml"
    xml = _sheet(p)
    assert 't="shared"' not in xml          # the group was written out plainly
    assert '<c r="C4"><f>B4*2</f>' in xml and '<c r="C21"><f>B21*2</f>' in xml
    after = _cached(p)
    for cell, v in before.items():
        if cell not in ("C3", "D2"):
            assert after.get(cell) == v, cell
    assert r["cached_values"]["possibly_stale_cells"] == ["Calc!D2"]


# ================================================================ notes, tables

def test_notes_add_edit_delete_on_a_picture_workbook(tmp_path):
    p = _copy(tmp_path, "pictures_printer.xlsx")
    orig = _parts(p)
    loc = {"cell": "B50", "sheet": S}
    for action, kw in [("add", {"text": "first"}), ("edit", {"text": "second"})]:
        r = annotations.manage_comment(p, action, loc, **kw)
        assert r["route"] == "raw_ooxml"
    r = annotations.manage_comment(p, "add", {"cell": "C50", "sheet": S}, text="other")
    listed = annotations.manage_comment(p, "list")
    assert {(c["cell"], c["text"]) for c in listed["comments"]} == {("B50", "second"), ("C50", "other")}
    assert '<legacyDrawing r:id="' in _sheet(p)
    sheet = _sheet(p)
    assert sheet.index("<drawing ") < sheet.index("<legacyDrawing ")
    for action, where in [("delete", "C50"), ("delete", "B50")]:
        r = annotations.manage_comment(p, action, {"cell": where, "sheet": S})
        assert r["route"] == "raw_ooxml"
    after = _parts(p)
    assert annotations.manage_comment(p, "list")["count"] == 0
    for n in orig:
        if n != "[Content_Types].xml":
            assert after[n] == orig[n], n
    assert not any(n.startswith(("xl/comments", "xl/drawings/vml")) for n in after)


def test_hyperlinks_add_and_remove_on_a_picture_workbook(tmp_path):
    p = _copy(tmp_path, "pictures_printer.xlsx")
    rels_before = _parts(p)["xl/worksheets/_rels/sheet1.xml.rels"]
    annotations.manage_hyperlink(p, "add", {"cell": "A60", "sheet": S},
                                 target="https://example.com/a?b=1&c=2", tooltip="Policy")
    annotations.manage_hyperlink(p, "add", {"cell": "A61", "sheet": S},
                                 target="Report!A1", display="Back to top")
    links = {h["cell"]: h["target"] for h in annotations.manage_hyperlink(p, "list")["hyperlinks"]}
    assert links == {"A60": "https://example.com/a?b=1&c=2", "A61": "Report!A1"}
    sheet = _sheet(p)
    # schema order: hyperlinks before pageMargins and the drawing
    assert sheet.index("<hyperlinks>") < sheet.index("<pageMargins") < sheet.index("<drawing ")
    for cell in ("A60", "A61"):
        r = annotations.manage_hyperlink(p, "remove", {"cell": cell, "sheet": S})
        assert r["route"] == "raw_ooxml"
    assert "<hyperlinks" not in _sheet(p)
    assert _parts(p)["xl/worksheets/_rels/sheet1.xml.rels"] == rels_before


def test_threaded_comment_cell_takes_the_rewrite(tmp_path):
    p = _copy(tmp_path, "comments_legacy_threaded.xlsx")
    with pytest.raises(HazardRefused) as exc:
        annotations.manage_comment(p, "edit", {"cell": "B3", "sheet": "Review"}, text="x")
    assert "threaded comment" in str(exc.value)


def test_table_header_rename_writes_the_table_part_only(tmp_path):
    p = _copy(tmp_path, "table_slicer.xlsx")
    before = _parts(p)
    r = cells.set_cell(p, {"cell": "B1", "sheet": "Sales"}, "Product")
    assert r["route"] == "raw_ooxml"
    after = _parts(p)
    table = next(n for n in after if n.startswith("xl/tables/"))
    assert 'name="Product"' in after[table].decode()
    for n in before:
        if n.startswith(("xl/slicer", "xl/drawings/", "xl/slicerCaches/")):
            assert after[n] == before[n], n


# ================================================ the rewrite route carries

def test_external_links_survive_a_row_insert_x39(tmp_path):
    p = _copy(tmp_path, "names_extlink.xlsx")
    before = _parts(p)
    r = structure.modify_grid_structure(p, "insert_rows", {"cell": "A1", "sheet": "Budget"})
    assert r["route"] == "openpyxl" and r["verified"]
    after = _parts(p)
    links = [n for n in before if n.startswith("xl/externalLinks/")]
    assert links
    for n in links:
        assert after[n] == before[n], n
    assert _verify.package_problems(p) == set()
    assert "external link definitions" in r["carried_over"]


def test_libreoffice_custom_properties_survive_a_row_insert_x41(tmp_path):
    # LibreOffice writes an EMPTY docProps/custom.xml on every save; the
    # rewrite used to drop it and the inventory check then refused every
    # structural edit of every LibreOffice file.
    p = _copy(tmp_path, "libreoffice_plain.xlsx")
    before = _parts(p)["docProps/custom.xml"]
    r = structure.modify_grid_structure(p, "insert_rows", {"cell": "A30", "sheet": "Sheet1"})
    assert r["verified"] and r["route"] == "openpyxl"
    assert _parts(p)["docProps/custom.xml"] == before
    assert _verify.package_problems(p) == set()


def test_libreoffice_file_takes_cell_edits_with_its_x14_rules_intact(tmp_path):
    p = _copy(tmp_path, "libreoffice_saved.xlsx")
    before = _parts(p)
    r = fmt.format_cells(p, {"range": "D2:D4", "sheet": "Report"}, number_format="0.0%")
    assert r["route"] == "raw_ooxml"
    after = _parts(p)
    assert after["docProps/custom.xml"] == before["docProps/custom.xml"]
    sheet = next(n for n in before if n.startswith("xl/worksheets/sheet") and b"extLst" in before[n])
    assert after[sheet].count(b"<x14:") == before[sheet].count(b"<x14:")


def test_sensitivity_label_and_custom_xml_survive_a_row_insert(tmp_path):
    p = _copy(tmp_path, "customxml_label.xlsx")
    before = _parts(p)
    r = structure.modify_grid_structure(p, "insert_rows", {"cell": "A5", "sheet": "Budget"})
    assert r["verified"]
    after = _parts(p)
    for n in before:
        if n.startswith("customXml/") or n == "docProps/custom.xml":
            assert after.get(n) == before[n], n
    assert "MSIP_Label_" in after["docProps/custom.xml"].decode()
    assert _verify.package_problems(p) == set()


def test_rewrite_keeps_cached_results_and_automatic_heights(tmp_path):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    n_before = len(_cached(p))
    r = structure.modify_grid_structure(p, "insert_rows", {"cell": "A60", "sheet": "Sheet1"})
    assert r["route"] == "openpyxl"
    assert r["cached_values"]["kept"] >= n_before - 1
    xml = _sheet(p)
    assert "customHeight" not in xml


def test_macro_sheet_survives_cell_edits_and_refuses_the_rewrite_x42(tmp_path):
    p = _copy(tmp_path, "xlm_macrosheet.xlsm")
    before = _parts(p)
    r = cells.set_cell(p, {"cell": "B3", "sheet": "Orders"}, 12)
    assert r["route"] == "raw_ooxml"
    after = _parts(p)
    for n in before:
        if n.startswith("xl/macrosheets/"):
            assert after[n] == before[n]
    with pytest.raises(HazardRefused) as exc:
        structure.modify_grid_structure(p, "insert_rows", {"cell": "A2", "sheet": "Orders"})
    assert "Excel 4.0 macro sheets" in str(exc.value)
    vba = inspectors.inspect_vba(p)
    assert vba["has_macros"] is True and vba["xlm_macro_sheets"] == ["Macro1"]
    meta = lifecycle.get_workbook_metadata(p)
    kinds = {s["name"]: s.get("kind") for s in meta["sheets"]}
    assert kinds["Macro1"] == "Excel 4.0 macro sheet"


def test_vba_project_bytes_survive_a_cell_edit(tmp_path):
    p = _copy(tmp_path, "macro.xlsm", CORPUS)
    before = _parts(p)
    r = cells.set_cell(p, "A30", 1)
    assert r["route"] == "raw_ooxml"
    after = _parts(p)
    assert after["xl/vbaProject.bin"] == before["xl/vbaProject.bin"]
    assert before["xl/_rels/workbook.xml.rels"] == after["xl/_rels/workbook.xml.rels"]
    assert b"vbaProject" in after["[Content_Types].xml"]


# ===================================================== the integrity check

def test_integrity_check_catches_a_dangling_relationship(tmp_path):
    p = _copy(tmp_path, "names_extlink.xlsx")
    assert _verify.package_problems(p) == set()
    broken = tmp_path / "broken.xlsx"
    parts = _parts(p)
    rels = next(n for n in parts if n.startswith("xl/externalLinks/_rels/"))
    parts[rels] = re.sub(rb'<Relationship Id="rId1"[^>]*/>', b"", parts[rels])
    with zipfile.ZipFile(broken, "w") as z:
        for n, d in parts.items():
            z.writestr(n, d)
    ok, problems = _verify.integrity_check(str(broken), set())
    assert not ok and any("rId1" in x for x in problems)
    # a file that arrived broken is not this save's doing
    ok, _ = _verify.integrity_check(str(broken), _verify.package_problems(str(broken)))
    assert ok


# ============================================================ formula limits

@pytest.mark.parametrize("formula, words", [
    # each boundary measured in Excel (see core/limits.py)
    ("=" + "1+" * 4096 + "1", "8,192"),
    ("=" + "ABS(" * 66 + "1" + ")" * 66, "65"),
    ("=SUM(" + ",".join(["1"] * 256) + ")", "255"),
    ("=" + "+".join(["A1"] * 2049), "16,383"),
])
def test_excel_formula_limits_refuse_before_the_disk_x58(tmp_path, formula, words):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    before = Path(p).read_bytes()
    with pytest.raises(ExcelWouldRefuse) as exc:
        cells.set_cell(p, {"cell": "H1", "sheet": "Sheet1"}, formula)
    assert words in str(exc.value)
    assert Path(p).read_bytes() == before


@pytest.mark.parametrize("formula", [
    "=" + "+".join(["A1"] * 2048),               # 16,383 compiled bytes: opens
    "=" + "ABS(" * 65 + "1" + ")" * 65,          # 65 levels: opens
    "=SUM(" + ",".join(["1"] * 255) + ")",       # 255 arguments: opens
    "=" + "+".join(["1"] * 3000),                # constants are 3 bytes each
])
def test_formulas_at_the_measured_boundary_are_accepted(tmp_path, formula):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    r = cells.set_cell(p, {"cell": "H1", "sheet": "Sheet1"}, formula)
    assert r["verified"]


def test_formula_limit_holds_for_a_formula_written_straight_into_the_model(tmp_path):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    pkg = WorkbookPackage.open(p)
    pkg.workbook["Sheet1"]["H1"] = "=" + "B2+" * 3000 + "1"
    with pytest.raises(ExcelWouldRefuse):
        pkg.save()


# =================================================================== readers

def test_image_list_reads_the_drawing_not_the_model_x37(tmp_path, monkeypatch):
    p = _copy(tmp_path, "pictures_printer.xlsx")
    import openpyxl.reader.drawings as _rd
    monkeypatch.setattr(_rd, "find_images", lambda *a, **k: ([], []))
    out = objects.manage_image(p, "list")
    assert out["count"] == 2
    first = out["images"][0]
    assert first["placement"]["type"] == "two-cell"
    assert first["width"] and first["height"] and first["media"].startswith("xl/media/")


def test_dropdown_list_on_another_sheet_is_listed_x38(tmp_path):
    p = _copy(tmp_path, "x14_dv_databar_sparkline.xlsx")
    out = datavalidation.manage_data_validation(p, "list")
    x14 = [v for v in out["validations"] if v.get("stored_as") == "x14"]
    assert x14 and x14[0]["formula1"] == "Lists!$A$2:$A$5"
    assert x14[0]["editable"] is False and "note" in out


def test_grid_view_and_metadata_say_what_the_grid_hides(tmp_path):
    p = _copy(tmp_path, "pictures_printer.xlsx")
    v = view.get_grid_view(p, sheet=S)
    assert v["not_shown"]["images"]["count"] == 2
    meta = lifecycle.get_workbook_metadata(p)
    assert meta["sheets"][0]["objects"]["pictures"] == 2


def test_calc_staleness_never_reports_fresh_without_proof_x45(tmp_path):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    out = validation.validate(p, checks=["calc_staleness"])
    c = out["results"]["calc_staleness"]
    assert c["passed"] is None and c["findings"]["freshness"] == "unverified"
    assert out["passed"] is None and out["unverified"] == ["calc_staleness"]
    cells.set_cell(p, {"cell": "C5", "sheet": "Sheet1"}, 99)
    c = validation.validate(p, checks=["calc_staleness"])["results"]["calc_staleness"]
    assert c["passed"] is False and c["findings"]["freshness"] == "recalculation pending"
    r = cells.read_range(p, {"range": "E5:E5", "sheet": "Sheet1"})
    assert "recalculate" in r["cached_note"]


# ================================================================= diagnose

def test_diagnose_tells_the_truth_about_each_route(tmp_path):
    p = _copy(tmp_path, "pictures_printer.xlsx")
    d = lifecycle.diagnose_workbook(p)["routing"]
    assert d["cell_edits"]["route"] == "raw_ooxml"
    assert "set_cell" in d["cell_edits"]["ops"]
    assert d["structural_edits"]["route"] == "refuse"
    assert any("media" in x or "images" in x for x in d["structural_edits"]["would_lose"])
    p2 = _copy(tmp_path, "customxml_label.xlsx")
    d2 = lifecycle.diagnose_workbook(p2)["routing"]
    assert d2["structural_edits"]["route"] == "openpyxl"
    assert d2["structural_edits"]["carried_over"]


def test_surgical_switch_off_restores_the_rewrite(tmp_path, monkeypatch):
    monkeypatch.setenv("KS4XL_SURGICAL_SAVE", "off")
    p = _copy(tmp_path, "plain_wrap.xlsx")
    r = cells.set_cell(p, {"cell": "A80", "sheet": "Sheet1"}, 1)
    assert r["route"] == "openpyxl"


def test_structural_edit_never_takes_the_splice(tmp_path):
    p = _copy(tmp_path, "plain_wrap.xlsx")
    r = structure.modify_grid_structure(p, "insert_rows", {"cell": "A3", "sheet": "Sheet1"})
    assert r["route"] == "openpyxl"
    assert "inserts or deletes rows" in r["route_note"]
