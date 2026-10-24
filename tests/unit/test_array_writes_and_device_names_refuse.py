"""The two defect classes an earlier fix closed only in part: writes into part
of an array formula, and device-name spellings that hung instead of refusing.
Red before the follow-up fix."""

from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

import openpyxl
import pytest
from openpyxl.worksheet.formula import ArrayFormula

import family_harness as fh


def _array_book(p: Path) -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["B1"], ws["B2"] = 1, 2
    ws["A1"] = ArrayFormula("A1:A2", "=B1:B2*2")
    wb.save(p)
    return str(p)


def _sha(p) -> str:
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


# F-A2 (MAJOR, numbers safety) F-A is refused at WorkbookPackage.save only for
# writes that go through pkg.set_cell / set_formula. Writers that assign
# cell.value directly still write into part of an array: manage_hyperlink
# (display text, or the target when the cell is empty; ops/annotations.py)
# and import_data's formula-injection lint path (text starting with = + - @
# is stored as TEXT; ops/dataio.py). Excel's verdict (private hidden pool):
# at A1 the whole array is silently deleted (A2 goes blank); at A2 the written
# text is silently discarded and the array's result shown.

_ROUTES = {
    "hyperlink_display": ("xl_format", "manage_hyperlink",
                          {"action": "add", "target": "https://example.com", "display": "link"}),
    "hyperlink_no_display": ("xl_format", "manage_hyperlink",
                             {"action": "add", "target": "https://example.com"}),
    "import_neutralized": ("xl_edit", "import_data",
                           {"source": "=1+1", "fmt": "csv", "header": False}),
}


@pytest.mark.parametrize("cell", ["A1", "A2"])
@pytest.mark.parametrize("route", sorted(_ROUTES))
def test_fa2_every_writer_refuses_part_of_an_array(route, cell, tmp_path):
    p = _array_book(tmp_path / "arr.xlsx")
    before = _sha(p)
    tool, op, args = _ROUTES[route]
    payload, err, _ = fh.start().call(tool, {"op": op, "args": {"path": p, "location": cell,
                                                                **args}})
    assert err, f"{route} wrote {cell} inside array A1:A2: {payload}"
    assert _sha(p) == before


# sandbox.device_name reads the last component with
# os.path.basename, so a device named with a trailing colon ("CON:") or in the
# device namespace ("\.\CON", "\.\CONIN$", "\?\CONIN$") is not recognised:
# the op blocks on a console read until the call deadline and parks a thread.

_B = chr(92)
_DEVICES = ["CON:", "CONIN$:", _B * 2 + "." + _B + "CON", _B * 2 + "." + _B + "CONIN$",
            _B * 2 + "?" + _B + "CONIN$"]


@pytest.mark.skipif(sys.platform != "win32", reason="Windows device names")
@pytest.mark.timeout(180)
def test_fd2_every_device_spelling_refuses_instead_of_hanging():
    res = fh.stdio_session([("xl_read", {"op": "read_range",
                                         "args": {"path": d, "location": "A1"}}) for d in _DEVICES],
                           env={"KS4XL_CALL_TIMEOUT_S": "3"}, timeout=150)
    codes = {d: json.loads(r["text"])["error"]["code"] for d, r in zip(_DEVICES, res["results"])}
    assert all(c != "TIMEOUT" for c in codes.values()), codes
