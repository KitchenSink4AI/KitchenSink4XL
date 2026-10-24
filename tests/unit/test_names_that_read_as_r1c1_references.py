"""A defined name or table name that BEGINS with an R1C1 reference (R1_x, C21_Name,
R2D2, Rc21_Name) is refused by Excel: its own Names.Add answers "The syntax of this
name isn't correct", and a workbook that carries one will not open. The earlier check
refused only a name that was entirely a reference (R1C1, R2, C33), so the tool saved
the file, verified it, and answered ok. Names with no digit after the R or C
(R_x, RC_x) and ordinary names (Q21_Name, Rate_1) are fine."""
from __future__ import annotations

import openpyxl
import pytest

from xlsx_mcp.core.errors import XlMcpError
from xlsx_mcp.ops import names as _names, tables as _tables

BAD = ["R1_x", "R21_Name", "R2D2", "C21_Name", "C1x", "R1C1x", "Rc21_Name", "R1C_x"]
GOOD = ["R_x", "RC_x", "Q21_Name", "Rate_1", "Cost_2"]


def _make(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for r in range(1, 4):
        for c in range(1, 3):
            ws.cell(r, c, f"h{c}" if r == 1 else r * c)
    wb.save(path)
    wb.close()
    return str(path)


@pytest.mark.parametrize("bad", BAD)
def test_a_defined_name_that_starts_with_an_r1c1_reference_is_refused(tmp_path, bad):
    p = _make(tmp_path / "n.xlsx")
    with pytest.raises(XlMcpError):
        _names.manage_name(p, "add", name=bad, refers_to="=Data!$A$1:$A$3")
    _names.manage_name(p, "add", name="Fine_name", refers_to="=Data!$A$1:$A$3")
    with pytest.raises(XlMcpError):
        _names.manage_name(p, "rename", name="Fine_name", new_name=bad)


@pytest.mark.parametrize("bad", BAD)
def test_a_table_name_that_starts_with_an_r1c1_reference_is_refused(tmp_path, bad):
    p = _make(tmp_path / "t.xlsx")
    with pytest.raises(XlMcpError):
        _tables.create_table(p, {"range": "A1:B3"}, bad)
    _tables.create_table(p, {"range": "A1:B3"}, "Fine_table")
    with pytest.raises(XlMcpError):
        _tables.manage_table(p, "Fine_table", "rename", new_name=bad)


@pytest.mark.parametrize("good", GOOD)
def test_names_without_a_reference_prefix_are_accepted(tmp_path, good):
    p = _make(tmp_path / "g.xlsx")
    assert _names.manage_name(p, "add", name=good, refers_to="=Data!$A$1:$A$3")["ok"] is True
    q = _make(tmp_path / "g2.xlsx")
    assert _tables.create_table(q, {"range": "A1:B3"}, good)["ok"] is True


@pytest.mark.parametrize("good", ["r7.name", "R0_x", "C16385_x", "RC0_x", "R1048577_x"])
def test_a_defined_name_excel_accepts_is_accepted_even_if_it_starts_like_a_reference(tmp_path, good):
    """Excel's Names.Add accepts these (the number is not a real row or column,
    or a period follows it). Table names keep the stricter rule."""
    p = _make(tmp_path / "d.xlsx")
    assert _names.manage_name(p, "add", name=good, refers_to="=Data!$A$1:$A$3")["ok"] is True
