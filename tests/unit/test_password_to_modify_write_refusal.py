"""A workbook with a "password to modify" never reaches
a read-write Excel open.

Such a file is an ordinary zip (workbook.xml carries <fileSharing
reservationPassword=...> or the modern hashValue form), so the encryption
sniff passes it. Opened read-write, Excel shows a Password dialog that
DisplayAlerts does not suppress; in the server's hidden Excel it appeared on
the user's desktop with Excel's main window. Every COM route that writes now
refuses it before Excel starts, and the shared open refuses a read-write open
of it whatever route reached it. None of these tests starts Excel: run_com
and Workbooks.Open are tripwires.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.core import hazard
from xlsx_mcp.core.errors import WorkbookProtected

SRC = Path(__file__).resolve().parents[2] / "src" / "xlsx_mcp"

LEGACY = b'<fileSharing userName="Owner" reservationPassword="CC3D"/>'
MODERN = (b'<fileSharing readOnlyRecommended="1" userName="Owner" '
          b'algorithmName="SHA-512" hashValue="aGFzaA==" saltValue="c2FsdA==" '
          b'spinCount="100000"/>')
RO_ONLY = b'<fileSharing readOnlyRecommended="1"/>'


def _book(tmp_path, sharing: bytes | None, name="reserved.xlsx"):
    base = str(tmp_path / ("base_" + name))
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.active["A2"] = 2
    wb.save(base)
    out = str(tmp_path / name)
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in zin.infolist():
            b = zin.read(i.filename)
            if i.filename == "xl/workbook.xml" and sharing:
                k = b.index(b"<workbookPr")
                b = b[:k] + sharing + b[k:]
            zout.writestr(i, b)
    return out


@pytest.fixture
def no_excel(monkeypatch):
    from xlsx_mcp.com import session as _session

    def tripwire(*a, **k):
        raise AssertionError("Excel was started for a workbook with a password to modify")
    monkeypatch.setattr(_session, "run_com", tripwire)
    return tripwire


# ------------------------------------------------------------- detection

def test_detects_legacy_and_modern_password_to_modify(tmp_path):
    assert hazard.write_reservation(_book(tmp_path, LEGACY, "a.xlsx"))["password_to_modify"]
    got = hazard.write_reservation(_book(tmp_path, MODERN, "b.xlsx"))
    assert got["password_to_modify"] and got["read_only_recommended"]
    assert got["reserved_by"] == "Owner"


def test_read_only_recommended_alone_and_plain_files_are_not_reserved(tmp_path):
    assert hazard.write_reservation(_book(tmp_path, RO_ONLY, "c.xlsx")) is None
    assert hazard.write_reservation(_book(tmp_path, None, "d.xlsx")) is None


def test_not_a_zip_is_not_reserved(tmp_path):
    p = tmp_path / "x.xlsx"
    p.write_bytes(b"not a zip")
    assert hazard.write_reservation(str(p)) is None


# ----------------------------------------------------- every writing route

ROUTES = {
    "com_autofit": lambda c, p: c.com_autofit(p),
    "recalculate": lambda c, p: c.recalculate(p, engine="com"),
    "com_goal_seek": lambda c, p: c.com_goal_seek(p, "A2", 5, "A1"),
    "com_manage_pivot": lambda c, p: c.com_manage_pivot(p, "refresh", name="P"),
    "com_set_sparkline": lambda c, p: c.com_set_sparkline(p, "clear", location="B1"),
    "com_save_with_password": lambda c, p: c.com_save_with_password(p, "pw"),
}


@pytest.mark.parametrize("sharing", [LEGACY, MODERN], ids=["legacy", "modern"])
@pytest.mark.parametrize("route", sorted(ROUTES))
def test_every_writing_com_route_refuses_before_excel(tmp_path, no_excel, route, sharing):
    from xlsx_mcp.ops import comtier
    p = _book(tmp_path, sharing)
    before = Path(p).read_bytes()
    with pytest.raises(WorkbookProtected) as ei:
        ROUTES[route](comtier, p)
    msg = str(ei.value)
    assert "password to modify" in msg
    assert "General Options" in msg and "Excel was not started" in msg
    assert Path(p).read_bytes() == before
    assert not (Path(p).parent / ".ks4xl-backups").exists() or not any(
        (Path(p).parent / ".ks4xl-backups").rglob("*.xlsx"))


def test_the_refusal_reaches_the_envelope_as_a_typed_code(tmp_path):
    from xlsx_mcp import envelope
    exc = hazard.write_reserved_refusal(_book(tmp_path, LEGACY))
    assert envelope.classify(exc) == "UNSUPPORTED_CONTENT"
    r = envelope.refusal(exc)
    assert "password to modify" in str(r)


# --------------------------------------------------- the shared open itself

class _FakeWorkbooks:
    def __init__(self):
        self.calls = []

    def Open(self, *a):
        self.calls.append(a)
        # The shared open checks that Excel opened THE file
        import types
        return types.SimpleNamespace(FullName=a[0], ReadOnly=bool(a[2]))


class _FakeApp:
    def __init__(self):
        self.Workbooks = _FakeWorkbooks()


def test_shared_open_refuses_a_read_write_open(tmp_path):
    from xlsx_mcp.com import session as _session
    app = _FakeApp()
    with pytest.raises(WorkbookProtected):
        _session.open_workbook(app, _book(tmp_path, LEGACY))
    assert app.Workbooks.calls == []


def test_shared_open_allows_a_read_only_open(tmp_path):
    from xlsx_mcp.com import session as _session
    app = _FakeApp()
    _session.open_workbook(app, _book(tmp_path, LEGACY), read_only=True)
    assert len(app.Workbooks.calls) == 1
    assert app.Workbooks.calls[0][2] is True  # ReadOnly


def test_shared_open_passes_a_plain_file(tmp_path):
    from xlsx_mcp.com import session as _session
    app = _FakeApp()
    _session.open_workbook(app, _book(tmp_path, RO_ONLY))
    assert len(app.Workbooks.calls) == 1


# ------------------------------------------------------------ source guard

def test_workbooks_open_is_called_only_by_the_shared_open():
    """Any new route that opens a workbook must go through
    session.open_workbook, which carries the F1 refusal."""
    offenders = []
    for f in SRC.rglob("*.py"):
        text = f.read_text(encoding="utf-8")
        for m in re.finditer(r"Workbooks\.Open\(", text):
            line = text.count("\n", 0, m.start()) + 1
            if not (f.name == "session.py" and f.parent.name == "com"):
                offenders.append(f"{f.relative_to(SRC)}:{line}")
    assert offenders == [], offenders
    s = (SRC / "com" / "session.py").read_text(encoding="utf-8")
    assert s.count("Workbooks.Open(") == 1
    body = s[s.index("def open_workbook"):s.index("def _assert_hygiene")]
    assert "refuse_if_write_reserved" in body


def test_every_read_write_open_in_comtier_has_the_early_refusal():
    s = (SRC / "ops" / "comtier.py").read_text(encoding="utf-8")
    for fn in ("def _run_mutation", "def com_save_with_password", "def recalculate"):
        start = s.index(fn)
        nxt = s.find("\ndef ", start + 1)
        assert "refuse_if_write_reserved" in s[start:nxt], fn
