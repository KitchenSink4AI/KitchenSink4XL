"""Every route that reads a workbook waits out another save's swap.

A save puts the new version in place with ReplaceFileW, which moves the old
file aside and then moves the new one in, so the workbook's path is missing
for a few milliseconds (MEASURED: 299 of 300 saves, up
to 9.3 ms). Every reading route, set_cell, the copy source, the snapshot and
the export answered "no such workbook" at once inside that gap (audit_styles
with a raw "[Errno 2]"). They now go through one helper
(core/swapwait.py require_workbook), which waits while a save has the old
file aside and answers at once when nothing does.

The swap is held open here for 0.4 s, far longer than a real one, so a route
that does not wait cannot pass by luck. No Excel.
"""
from __future__ import annotations

import os
import threading
import time
import uuid

import openpyxl
import pytest

from xlsx_mcp.core import swapwait
from xlsx_mcp.core.errors import WorkbookNotFound
from xlsx_mcp.ops import (backups, cells, dataio, format as format_ops, formulas,
                          inspectors, lifecycle, search, validation, view)


def _book(p):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "v0"
    wb.active["B1"] = 2
    wb.active["C1"] = "=B1*2"
    wb.save(p)


ROUTES = {
    "read_range": lambda p, t: cells.read_range(p, "A1:B1"),
    "get_cells": lambda p, t: cells.get_cells(p, ["A1"]),
    "query_range": lambda p, t: cells.query_range(p, "A1:B1", header=False),
    "set_cell": lambda p, t: cells.set_cell(p, "D1", 3),
    "get_workbook_metadata": lambda p, t: lifecycle.get_workbook_metadata(p),
    "get_grid_view": lambda p, t: view.get_grid_view(p),
    "find_cells": lambda p, t: search.find_cells(p, "v0"),
    "validate": lambda p, t: validation.validate(p),
    "audit_formulas": lambda p, t: formulas.audit_formulas(p),
    "audit_styles": lambda p, t: format_ops.audit_styles(p),
    "get_external_links": lambda p, t: inspectors.get_external_links(p),
    "inspect_vba": lambda p, t: inspectors.inspect_vba(p),
    "get_pivot": lambda p, t: inspectors.get_pivot(p),
    "get_connections": lambda p, t: inspectors.get_connections(p),
    "copy_workbook_source": lambda p, t: lifecycle.copy_workbook(p, str(t / "copy.xlsx")),
    "snapshot": lambda p, t: backups.create_snapshot(p, dest_dir=str(t / "snaps")),
    "export_file": lambda p, t: dataio.export_file(p, out_dir=str(t / "out")),
}


@pytest.mark.parametrize("legacy", [False, True], ids=["aside", "earlier_versions_aside"])
@pytest.mark.parametrize("route", sorted(ROUTES))
def test_a_call_inside_another_saves_swap_waits_it_out(tmp_path, route, legacy):
    folder = tmp_path / "work"
    folder.mkdir()
    (tmp_path / "snaps").mkdir()
    (tmp_path / "out").mkdir()
    p = folder / "Book.xlsx"
    _book(p)
    name = (f".ks4xl-write-{uuid.uuid4().hex}.replaced" if legacy
            else swapwait.aside_name("Book.xlsx", uuid.uuid4().hex[:12]))
    aside = folder / name
    os.replace(p, aside)  # ReplaceFileW's first move: the path is missing until the second

    def back():
        time.sleep(0.4)
        os.replace(aside, p)
    th = threading.Thread(target=back)
    th.start()
    try:
        ROUTES[route](str(p), tmp_path)
    finally:
        th.join()


def test_a_workbook_that_is_not_there_is_reported_at_once(tmp_path):
    t0 = time.monotonic()
    with pytest.raises(WorkbookNotFound):
        cells.read_range(str(tmp_path / "Nowhere.xlsx"), "A1")
    assert time.monotonic() - t0 < 0.2


def test_the_helper_only_reads(tmp_path):
    """The read-only tools reach it, so it may not change anything."""
    import inspect
    source = inspect.getsource(swapwait)
    for marker in ("os.remove", "os.replace", "os.rename", "os.unlink", "write_text",
                   "write_bytes", "mkdir", "safesave"):
        assert marker not in source, marker
