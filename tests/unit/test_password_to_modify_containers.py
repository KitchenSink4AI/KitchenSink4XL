"""The password-to-modify check reads every form
Excel opens, not only a UTF-8 workbook.xml.

The earlier check regex-scanned the workbook part's raw bytes for an ASCII
<fileSharing ...>. Two ordinary inputs walked past it and Excel, opened for
editing, put its Password dialog on the desktop:

  * .xlsb: the workbook part is the binary xl/workbook.bin, where the
    setting is a BrtFileSharingIso (0x2A4) or BrtFileSharing (0x224)
    record, not text.
  * workbook.xml stored as UTF-16 (Excel opens it), where every
    character of "<fileSharing" is followed by a zero byte.

None of these tests starts Excel: run_com is a tripwire.
"""
from __future__ import annotations

import zipfile
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.core import hazard
from xlsx_mcp.core.errors import WorkbookProtected


# ------------------------------------------------------------ xlsb builder

def _rec(rt: int, body: bytes = b"") -> bytes:
    """One BIFF12 record: 1-2 byte type, 1-4 byte size (7 bits a byte)."""
    head = bytearray()
    if rt >= 0x80:
        head += bytes([(rt & 0x7F) | 0x80, rt >> 7])
    else:
        head.append(rt)
    n = len(body)
    while True:
        b = n & 0x7F
        n >>= 7
        head.append(b | (0x80 if n else 0))
        if not n:
            break
    return bytes(head) + body


def _wide(s: str) -> bytes:
    return len(s).to_bytes(4, "little") + s.encode("utf-16-le")


BEGIN_BOOK = _rec(0x83)
END_BOOK = _rec(0x84)
FILE_VERSION = _rec(0x80, b"\x00" * 16 + _wide("xl"))
SHEETS = _rec(0x8F) + _rec(0x9C, b"\x00" * 8 + _wide("rId1") + _wide("Sheet1")) + _rec(0x90)
#: Excel's own form for a password set today: the ISO hash record, then the
#: legacy record with no legacy verifier.
ISO = _rec(0x2A4, (100000).to_bytes(4, "little") + b"\x00\x00" + _wide("Owner")
           + (64).to_bytes(4, "little") + b"\x11" * 64
           + (16).to_bytes(4, "little") + b"\x22" * 16 + _wide("SHA-512"))
LEGACY_NONE = _rec(0x224, b"\x00\x00" + b"\x00\x00" + _wide("Owner"))
LEGACY_PW = _rec(0x224, b"\x00\x00" + b"\x3d\xcc" + _wide("Owner"))
RO_ONLY = _rec(0x224, b"\x01\x00" + b"\x00\x00" + _wide("Owner"))

_RELS = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
         b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
         b'relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>')
_CT = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
       b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
       b'<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>'
       b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
       b'</Types>')


def _xlsb(tmp_path, *records: bytes, name="book.xlsb", raw: bytes | None = None) -> str:
    out = str(tmp_path / name)
    body = raw if raw is not None else BEGIN_BOOK + FILE_VERSION + b"".join(records) + SHEETS + END_BOOK
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CT)
        z.writestr("_rels/.rels", _RELS)
        z.writestr("xl/workbook.bin", body)
    return out


# --------------------------------------------------------- xml re-encoder

MODERN = ('<fileSharing userName="Owner" algorithmName="SHA-512" hashValue="aGFzaA==" '
          'saltValue="c2FsdA==" spinCount="100000"/>')


def _xlsx(tmp_path, encode, sharing: str | None = MODERN, name="enc.xlsx") -> str:
    base = str(tmp_path / ("base_" + name))
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(base)
    out = str(tmp_path / name)
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in zin.infolist():
            b = zin.read(i.filename)
            if i.filename == "xl/workbook.xml":
                t = b.decode("utf-8")
                if sharing:
                    k = t.index("<workbookPr")
                    t = t[:k] + sharing + t[k:]
                b = encode(t)
            zout.writestr(i, b)
    return out


def _utf16le_bom(t):
    return b"\xff\xfe" + t.replace('encoding="UTF-8"', 'encoding="UTF-16"').encode("utf-16-le")


def _utf16be_bom(t):
    return b"\xfe\xff" + t.replace('encoding="UTF-8"', 'encoding="UTF-16"').encode("utf-16-be")


def _utf16le_nobom(t):
    return t.replace('encoding="UTF-8"', 'encoding="UTF-16"').encode("utf-16-le")


def _utf8_bom(t):
    return b"\xef\xbb\xbf" + t.encode("utf-8")


def _latin1_declared(t):
    return t.replace('encoding="UTF-8"', 'encoding="ISO-8859-1"').replace(
        'userName="Owner"', 'userName="Owéner"').encode("latin-1")


ENCODERS = {"utf16le_bom": _utf16le_bom, "utf16be_bom": _utf16be_bom,
            "utf16le_nobom": _utf16le_nobom, "utf8_bom": _utf8_bom,
            "latin1_declared": _latin1_declared}


@pytest.fixture
def no_excel(monkeypatch):
    from xlsx_mcp.com import session as _session

    def tripwire(*a, **k):
        raise AssertionError("Excel was started for a workbook with a password to modify")
    monkeypatch.setattr(_session, "run_com", tripwire)


# ---------------------------------------------------------------- xlsb

def test_xlsb_iso_hash_is_a_password_to_modify(tmp_path):
    got = hazard.write_reservation(_xlsb(tmp_path, ISO, LEGACY_NONE))
    assert got and got["password_to_modify"] is True
    assert got["reserved_by"] == "Owner"


def test_xlsb_legacy_verifier_is_a_password_to_modify(tmp_path):
    got = hazard.write_reservation(_xlsb(tmp_path, LEGACY_PW))
    assert got and got["password_to_modify"] is True


def test_xlsb_read_only_recommended_alone_and_plain_are_not_reserved(tmp_path):
    assert hazard.write_reservation(_xlsb(tmp_path, RO_ONLY, name="ro.xlsb")) is None
    assert hazard.write_reservation(_xlsb(tmp_path, name="plain.xlsb")) is None


def test_xlsb_part_that_cannot_be_walked_is_not_called_safe(tmp_path):
    """A truncated record: the check cannot rule a password out, so the
    writing routes must not hand the file to Excel for editing."""
    broken = BEGIN_BOOK + FILE_VERSION + b"\xa4\x05\x7f" + b"\x00" * 10
    got = hazard.write_reservation(_xlsb(tmp_path, raw=broken, name="broken.xlsb"))
    assert got is not None and got["password_to_modify"] is None
    with pytest.raises(WorkbookProtected) as ei:
        hazard.refuse_if_write_reserved(_xlsb(tmp_path, raw=broken, name="b2.xlsb"))
    assert "cannot read" in str(ei.value) and "rule out" in str(ei.value)


# ------------------------------------------------------- xml encodings

@pytest.mark.parametrize("enc", sorted(ENCODERS))
def test_password_to_modify_is_read_in_every_encoding(tmp_path, enc):
    got = hazard.write_reservation(_xlsx(tmp_path, ENCODERS[enc], name=f"{enc}.xlsx"))
    assert got and got["password_to_modify"] is True, enc


@pytest.mark.parametrize("enc", sorted(ENCODERS))
def test_no_false_alarm_in_any_encoding(tmp_path, enc):
    p = _xlsx(tmp_path, ENCODERS[enc], sharing='<fileSharing readOnlyRecommended="1"/>',
              name=f"ro_{enc}.xlsx")
    assert hazard.write_reservation(p) is None, enc
    p = _xlsx(tmp_path, ENCODERS[enc], sharing=None, name=f"plain_{enc}.xlsx")
    assert hazard.write_reservation(p) is None, enc


# ------------------------------------------------- every writing route, both

ROUTES = {
    "com_autofit": lambda c, p: c.com_autofit(p),
    "recalculate": lambda c, p: c.recalculate(p, engine="com"),
    "com_goal_seek": lambda c, p: c.com_goal_seek(p, "A2", 5, "A1"),
    "com_manage_pivot": lambda c, p: c.com_manage_pivot(p, "refresh", name="P"),
    "com_set_sparkline": lambda c, p: c.com_set_sparkline(p, "clear", location="B1"),
    "com_save_with_password": lambda c, p: c.com_save_with_password(p, "pw"),
}


@pytest.mark.parametrize("kind", ["xlsb", "utf16"])
@pytest.mark.parametrize("route", sorted(ROUTES))
def test_every_writing_route_refuses_before_excel(tmp_path, no_excel, route, kind):
    from xlsx_mcp.ops import comtier
    p = _xlsb(tmp_path, ISO, LEGACY_NONE) if kind == "xlsb" else _xlsx(tmp_path, _utf16le_bom)
    before = Path(p).read_bytes()
    with pytest.raises(WorkbookProtected) as ei:
        ROUTES[route](comtier, p)
    assert "password to modify" in str(ei.value)
    assert "Excel was not started" in str(ei.value)
    assert Path(p).read_bytes() == before
