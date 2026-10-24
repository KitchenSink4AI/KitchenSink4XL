"""com_save_with_password refuses a legacy .xls
before Excel starts, as every other writing route does.

The shared mutation path refused BIFF up front, but com_save_with_password
has its own open and had no such check; the password-to-modify check reads
OOXML packages only, so a .xls with a password to modify reached a
read-write open and Excel showed its Password dialog. Every route that saves
a workbook back to its own path now takes the same refusal. None of these
tests starts Excel: run_com is a tripwire.
"""
from __future__ import annotations

import re
import struct
from pathlib import Path

import pytest

from xlsx_mcp.core.errors import UnsupportedStructure

SRC = Path(__file__).resolve().parents[2] / "src" / "xlsx_mcp"


def _cfb_bytes(streams: list[str]) -> bytes:
    """A minimal OLE compound file: header plus one directory sector."""
    header = bytearray(512)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<H", header, 26, 3)
    struct.pack_into("<H", header, 28, 0xFFFE)
    struct.pack_into("<H", header, 30, 9)
    struct.pack_into("<H", header, 32, 6)
    struct.pack_into("<i", header, 44, 1)
    struct.pack_into("<i", header, 48, 0)
    for off in range(76, 512, 4):
        struct.pack_into("<i", header, off, -1)

    def entry(name: str, etype: int) -> bytes:
        e = bytearray(128)
        raw = name.encode("utf-16-le")
        e[0:len(raw)] = raw
        struct.pack_into("<H", e, 64, len(raw) + 2)
        e[66] = etype
        e[67] = 1
        for off in (68, 72, 76):
            struct.pack_into("<i", e, off, -1)
        return bytes(e)

    sector = bytearray(512)
    for i, e in enumerate([entry("Root Entry", 5)] + [entry(s, 2) for s in streams]):
        sector[i * 128:(i + 1) * 128] = e
    return bytes(header) + bytes(sector)


@pytest.fixture
def no_excel(monkeypatch):
    from xlsx_mcp.com import session as _session

    def tripwire(*a, **k):
        raise AssertionError("Excel was started for a legacy .xls on a writing route")
    monkeypatch.setattr(_session, "run_com", tripwire)


@pytest.mark.parametrize("kw", [{"password": "x"},
                                {"password": "", "current_password": "old"},
                                {"password": "new", "current_password": "old"}],
                         ids=["add", "remove", "rekey"])
def test_save_with_password_refuses_a_legacy_xls_before_excel(tmp_path, no_excel, kw):
    from xlsx_mcp.ops import comtier
    p = tmp_path / "old.xls"
    p.write_bytes(_cfb_bytes(["Workbook"]))
    before = p.read_bytes()
    with pytest.raises(UnsupportedStructure) as ei:
        comtier.com_save_with_password(str(p), **kw)
    msg = str(ei.value)
    assert "97-2003" in msg and "com_convert_format" in msg
    assert p.read_bytes() == before
    assert not (tmp_path / ".ks4xl-backups").exists()


def test_the_mutation_path_and_the_password_path_share_one_refusal():
    s = (SRC / "ops" / "comtier.py").read_text(encoding="utf-8")
    for fn in ("def _run_mutation", "def com_save_with_password"):
        start = s.index(fn)
        nxt = s.find("\ndef ", start + 1)
        assert "_refuse_legacy_biff(" in s[start:nxt], fn
    # and no second copy of the inline check survives anywhere in the module
    assert len(re.findall(r"== _hazard\.OLE_KIND_BIFF", s)) == 1
