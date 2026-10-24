"""The no-ReplaceFile fallback keeps every named stream.

Where ReplaceFileW is not supported (some network and WebDAV drives), the
save renames the new version over the workbook and carries the workbook's
identity over by hand first. That carry took only the Zone.Identifier
stream; any other named stream on the workbook was lost on every writing
route, while ReplaceFileW keeps them all (MEASURED). The
carry now lists the streams (FindFirstStreamW) and copies each one. No Excel.
"""
from __future__ import annotations

import os
import sys

import openpyxl
import pytest

from xlsx_mcp.core import safesave
from xlsx_mcp.ops import backups, cells, lifecycle

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="NTFS streams")

STREAMS = {"Zone.Identifier": b"[ZoneTransfer]\r\nZoneId=3\r\n",
           "ks4verify": b"keep me",
           "big.bin": os.urandom(300_000)}


def _book(p, v="v0"):
    wb = openpyxl.Workbook()
    wb.active["A1"] = v
    wb.save(p)


def _dress(p):
    for name, data in STREAMS.items():
        with open(f"{p}:{name}", "wb") as fh:
            fh.write(data)


def _streams(p):
    out = {}
    for name in STREAMS:
        try:
            with open(f"{p}:{name}", "rb") as fh:
                out[name] = fh.read()
        except OSError:
            out[name] = None
    return out


@pytest.fixture
def no_replacefile(monkeypatch):
    real = safesave._replace_file()
    monkeypatch.setattr(safesave, "_REPLACE_FILE", (lambda *a: 0, real[1], lambda: 50, real[3]))


@pytest.mark.parametrize("route", ["set_cell", "restore", "copy_overwrite"])
def test_every_named_stream_survives_the_fallback(tmp_path, no_replacefile, route):
    p = tmp_path / "Book.xlsx"
    _book(p, "v1")
    if route == "restore":
        safesave.prepare_rotation(p).commit()  # prev is a link of v1 ...
        _book(tmp_path / "v2.xlsx", "v2")
        os.replace(tmp_path / "v2.xlsx", p)  # ... so v2 arrives as a new file
    _dress(p)
    if route == "set_cell":
        cells.set_cell(str(p), "A1", "edited")
    elif route == "restore":
        backups.manage_backups("restore", path=str(p), source="prev")
    else:
        src = tmp_path / "src.xlsx"
        _book(src, "copied")
        lifecycle.copy_workbook(str(src), str(p), overwrite=True)
    assert _streams(p) == STREAMS


def test_the_stream_list_names_each_stream(tmp_path):
    p = tmp_path / "Book.xlsx"
    _book(p)
    _dress(p)
    assert sorted(safesave._named_streams(str(p))) == sorted(STREAMS)
