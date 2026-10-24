"""com_render_sheet judges the workbook before the session it runs in.

Rendering needs the Windows clipboard, and a locked or disconnected session
refuses the clipboard to every process. A workbook this server refuses on its
own (a part that declares a DTD, a decompression bomb, a part named twice) must
be refused as unsupported content in every session, not as a missing
clipboard in some sessions. The clipboard is switched off here, so the order
is pinned without needing a locked session. No Excel; every workbook is made
up and built by code.
"""
from __future__ import annotations

import zipfile

import openpyxl
import pytest

from xlsx_mcp.envelope import classify
from xlsx_mcp.ops import comtier


def _plain(tmp_path):
    src = tmp_path / "plain.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "made up"
    wb.save(src)
    return src


def _rewrite(src, dst, change):
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in zi.infolist():
            data = zi.read(i.filename)
            for name, d in change(i.filename, data):
                zo.writestr(name, d)


def _dtd(tmp_path):
    p = tmp_path / "dtd.xlsx"
    _rewrite(_plain(tmp_path), p, lambda n, d: [(n, b"<!DOCTYPE w>" + d if n == "xl/worksheets/sheet1.xml" else d)])
    return p


def _twice(tmp_path):
    p = tmp_path / "twice.xlsx"
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        _rewrite(_plain(tmp_path), p, lambda n, d: [(n, d)] + ([(n, d)] if n == "xl/worksheets/sheet1.xml" else []))
    return p


@pytest.mark.parametrize("make", [_dtd, _twice], ids=["dtd", "part_named_twice"])
def test_render_refuses_an_unsupported_package_even_without_a_clipboard(tmp_path, monkeypatch, make):
    p = make(tmp_path)
    monkeypatch.setattr(comtier, "_clipboard_available", lambda: False)
    with pytest.raises(Exception) as ei:
        comtier.com_render_sheet(str(p), str(tmp_path / "out.png"))
    assert classify(ei.value) == "UNSUPPORTED_CONTENT", str(ei.value)
    assert not (tmp_path / "out.png").exists()
