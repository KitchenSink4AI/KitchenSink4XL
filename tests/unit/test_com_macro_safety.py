"""The COM tier never runs a workbook's code.

Excel started for automation enables macros by default. Two layers close
that:

  1. Every pooled worker is spawned with AutomationSecurity set to
     ForceDisable, which disables the VBA project and auto-run macros
     (Workbook_Open, Auto_Open) for every workbook the worker opens.
  2. ForceDisable does NOT disable Excel 4.0 (XLM) macro sheets (Microsoft
     documents the exclusion), so a workbook that carries one is refused at
     the one open choke point every COM operation passes, before Excel
     starts. The file-based route reads and edits such a workbook without
     running anything.

The workbooks here are built by code. The XLM one is a minimal OOXML
package with a macrosheet part; it is never opened in Excel in these pure
tests.
"""

from __future__ import annotations

import zipfile

import pytest

from xlsx_mcp.com import instances as _instances
from xlsx_mcp.com import session as _session
from xlsx_mcp.core import hazard as _hazard
from xlsx_mcp.core.errors import UnsupportedStructure

_CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="{wb_ct}"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
{macro_override}</Types>"""

_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

_WB = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Data" sheetId="1" r:id="rId1"/>{macro_sheet}</sheets></workbook>"""

_WB_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
{macro_rel}</Relationships>"""

_SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"""

#: A macro sheet that ADDS; no Auto_Open name. It never runs here: these
#: tests only read the package.
_MACROSHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xm:macrosheet xmlns:xm="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<xm:sheetData><xm:row r="1"><xm:c r="A1"><xm:f>SUM(Data!A1:A1)</xm:f><xm:v>1</xm:v></xm:c></xm:row></xm:sheetData></xm:macrosheet>"""


def _write(path, parts: dict) -> str:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return str(path)


def _plain_xlsx(tmp_path, name="plain.xlsx") -> str:
    parts = {
        "[Content_Types].xml": _CT.format(
            wb_ct="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            macro_override=""),
        "_rels/.rels": _ROOT_RELS,
        "xl/workbook.xml": _WB.format(macro_sheet=""),
        "xl/_rels/workbook.xml.rels": _WB_RELS.format(macro_rel=""),
        "xl/worksheets/sheet1.xml": _SHEET,
    }
    return _write(tmp_path / name, parts)


def _xlm_xlsm(tmp_path, name="xlm.xlsm") -> str:
    parts = {
        "[Content_Types].xml": _CT.format(
            wb_ct="application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            macro_override='<Override PartName="/xl/macrosheets/sheet1.xml" '
            'ContentType="application/vnd.ms-excel.macrosheet+xml"/>'),
        "_rels/.rels": _ROOT_RELS,
        "xl/workbook.xml": _WB.format(
            macro_sheet='<sheet name="Macro1" sheetId="2" r:id="rId2"/>'),
        "xl/_rels/workbook.xml.rels": _WB_RELS.format(
            macro_rel='<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/xlMacrosheet" '
            'Target="macrosheets/sheet1.xml"/>'),
        "xl/worksheets/sheet1.xml": _SHEET,
        "xl/macrosheets/sheet1.xml": _MACROSHEET,
    }
    return _write(tmp_path / name, parts)


# ---------------------------------------------------- detection (pure)

def test_the_synthetic_xlm_workbook_is_detected(tmp_path):
    assert _hazard.has_xlm_macro_sheet(_xlm_xlsm(tmp_path)) is True


def test_a_plain_workbook_is_not_an_xlm_workbook(tmp_path):
    assert _hazard.has_xlm_macro_sheet(_plain_xlsx(tmp_path)) is False


def test_the_committed_xlm_fixture_is_detected():
    from pathlib import Path
    fix = (Path(__file__).resolve().parents[1] / "fixtures" / "preserve"
           / "xlm_macrosheet.xlsm")
    if not fix.is_file():
        pytest.skip("preserve corpus not built")
    assert _hazard.has_xlm_macro_sheet(str(fix)) is True


def test_a_dialog_sheet_is_not_an_xlm_macro_sheet(tmp_path):
    # Excel 5 dialog sheets live under xl/dialogsheets/, not macrosheets/.
    p = _plain_xlsx(tmp_path, "dialog.xlsx")
    with zipfile.ZipFile(p, "a") as z:
        z.writestr("xl/dialogsheets/sheet1.xml", "<x/>")
    assert _hazard.has_xlm_macro_sheet(p) is False


def test_a_non_package_file_is_not_an_xlm_workbook(tmp_path):
    p = tmp_path / "junk.xlsx"
    p.write_bytes(b"not a zip at all")
    assert _hazard.has_xlm_macro_sheet(str(p)) is False


# ----------------------------------------- refusal at the open choke point


class _ExcelThatMustNotBeTouched:
    """Any attribute access means Excel was asked to do something; the XLM
    refusal has to fire before that."""

    def __getattr__(self, name):
        raise AssertionError(
            f"the COM layer reached Excel (.{name}) on an XLM macro-sheet "
            "workbook; it must refuse before opening it")


@pytest.mark.parametrize("read_only", [True, False])
def test_open_workbook_refuses_an_xlm_workbook_before_touching_excel(
        tmp_path, read_only):
    p = _xlm_xlsm(tmp_path)
    with pytest.raises(UnsupportedStructure) as exc:
        _session.open_workbook(_ExcelThatMustNotBeTouched(), p,
                               read_only=read_only)
    assert "XLM" in str(exc.value) and "macro sheet" in str(exc.value)


def test_open_workbook_opens_a_plain_workbook_past_the_guard(tmp_path):
    # The plain workbook passes the XLM guard and only then reaches Excel
    # (the stub's first attribute access proves the guard let it through).
    p = _plain_xlsx(tmp_path)
    with pytest.raises(AssertionError) as exc:
        _session.open_workbook(_ExcelThatMustNotBeTouched(), p,
                               read_only=True)
    assert "reached Excel" in str(exc.value)


def test_the_refusal_names_the_file_route_remedy(tmp_path):
    exc = _hazard.xlm_macro_sheet_refusal(_xlm_xlsm(tmp_path))
    msg = str(exc)
    assert "was not started" in msg and "file-based tools" in msg
    assert exc.hint_tools == ("diagnose_workbook",)


# ------------------------------------------ the worker config (pure)

def test_force_disable_constant_is_three():
    # msoAutomationSecurityForceDisable.
    assert _instances.MSO_AUTOMATION_SECURITY_FORCE_DISABLE == 3


def test_new_app_sets_automation_security_to_force_disable():
    """_new_app spawns Excel, suppresses UI, and force-disables active
    content. With a fake win32 it proves the property is set to ForceDisable
    without starting Excel."""
    import sys
    import types

    seen = {}

    class _FakeApp:
        def __init__(self):
            self._v = {}

        def __setattr__(self, name, value):
            if name == "_v":
                object.__setattr__(self, name, value)
            else:
                self._v[name] = value

    fake = _FakeApp()
    fake_win32 = types.SimpleNamespace(DispatchEx=lambda progid: fake)
    mod = types.ModuleType("win32com.client")
    mod.DispatchEx = fake_win32.DispatchEx
    pkg = types.ModuleType("win32com")
    pkg.client = mod
    old_pkg = sys.modules.get("win32com")
    old_mod = sys.modules.get("win32com.client")
    sys.modules["win32com"] = pkg
    sys.modules["win32com.client"] = mod
    try:
        mgr = _instances.ExcelInstanceManager(journal_path=tmp_journal())
        app = mgr._new_app()
    finally:
        if old_pkg is not None:
            sys.modules["win32com"] = old_pkg
        else:
            sys.modules.pop("win32com", None)
        if old_mod is not None:
            sys.modules["win32com.client"] = old_mod
        else:
            sys.modules.pop("win32com.client", None)
    seen = app._v
    assert seen["Visible"] is False
    assert seen["DisplayAlerts"] is False
    assert seen["EnableEvents"] is False
    assert seen["AutomationSecurity"] == 3


def tmp_journal():
    import tempfile
    import os
    return os.path.join(tempfile.gettempdir(),
                        f"ks4xl-macro-safety-journal-{os.getpid()}.json")
