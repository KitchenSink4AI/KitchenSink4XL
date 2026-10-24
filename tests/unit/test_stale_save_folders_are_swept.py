"""A private save folder left in the temp area is swept.

Every COM save has Excel write its copy into a private folder in the
system's temp area. A server killed in the middle of a save left
that folder there for good. The next COM save now removes such a folder,
but only when all of these hold:

- its name has this server's prefix (ks4xl-save-);
- it is a real folder (not a link or a junction) holding only files a save
  makes there: Excel's copy (book.<ext>), the scratch file Excel writes
  while it saves (eight hex digits), Excel's owner file;
- the server that made it is gone (its process ID is in the name), or the
  folder is more than an hour old (no save takes that long;
  an older folder without a process ID is judged by age alone).
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.com import session as com_session
from xlsx_mcp.ops import comtier

IS_WIN = sys.platform == "win32"
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
HOURS_AGO_2 = time.time() - 2 * 3600


def _dead_pid() -> int:
    """The process ID of a process that has ended."""
    p = subprocess.Popen([sys.executable, "-c", "pass"], stdin=subprocess.DEVNULL,
                         creationflags=NO_WINDOW)
    p.wait(30)
    return p.pid


def _plant(temp: Path, name: str, files=("book.xlsx",), age=None) -> Path:
    d = temp / name
    d.mkdir()
    for f in files:
        (d / f).write_bytes(b"PK\x03\x04 left by a killed save")
    if age is not None:
        for f in files:
            os.utime(d / f, (age, age))
        os.utime(d, (age, age))
    return d


class _Worker:
    pid = 0
    app = object()


class _Manager:
    def acquire(self):
        return _Worker()

    def release_to_pool(self, w):
        pass


class _FakeWB:
    def __init__(self, path):
        self.path = path
        self.FullName = os.path.abspath(path)
        self.ReadOnly = False
        self.Password = ""

    def SaveCopyAs(self, copy):
        wb = openpyxl.load_workbook(self.path)
        wb.active["A1"] = "edited by Excel"
        wb.save(copy)

    def Close(self, save):
        pass


@pytest.fixture
def temp(tmp_path, monkeypatch):
    t = tmp_path / "temp"
    t.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(t))
    return t


@pytest.fixture
def book(tmp_path, temp, monkeypatch):
    folder = tmp_path / "work"
    folder.mkdir()
    p = str(folder / "book.xlsx")
    wb = openpyxl.Workbook()
    wb.active["A1"] = "before"
    wb.save(p)
    monkeypatch.setattr(com_session, "run_com",
                        lambda label, job, timeout=None: job(_Manager()))
    monkeypatch.setattr(com_session, "_assert_hygiene", lambda app: ())
    monkeypatch.setattr(com_session, "_restore_hygiene", lambda app, *a: None)
    monkeypatch.setattr(com_session, "com_retry", lambda fn, label=None, **k: fn())
    monkeypatch.setattr(com_session, "open_workbook",
                        lambda app, p, password=None, **k: _FakeWB(p))
    return p


def _com_save(book):
    res = comtier._run_mutation("autofit(book.xlsx)", book, lambda app, wb: {"a": 1})
    assert res["ok"]


def test_the_next_com_save_sweeps_what_killed_saves_left(book, temp):
    dead = _dead_pid()
    gone = [
        _plant(temp, f"ks4xl-save-{dead}-a1b2c3d4"),
        _plant(temp, f"ks4xl-save-{dead}-scratch1", files=("book.xlsm", "4F2A9C10")),
        _plant(temp, f"ks4xl-save-{dead}-emptyabc", files=()),
        _plant(temp, f"ks4xl-save-{os.getpid()}-oldone12", age=HOURS_AGO_2),
        _plant(temp, "ks4xl-save-q1w2e3r4", age=HOURS_AGO_2),  # old-style name, an hour old
    ]
    _com_save(book)
    left = sorted(os.listdir(temp))
    for d in gone:
        assert d.name not in left, left
    assert left == [], left


def test_what_is_not_a_save_leftover_is_never_touched(book, temp, tmp_path):
    dead = _dead_pid()
    kept = [
        _plant(temp, f"ks4xl-save-{os.getpid()}-fresh123"),  # this server, a fresh one
        _plant(temp, "ks4xl-save-k3j4h5g6"),  # old-style name, fresh: the owner is unknown
        _plant(temp, f"ks4xl-save-{dead}-foreign1", files=("book.xlsx", "notes.txt"),
               age=HOURS_AGO_2),
        _plant(temp, "ks4xl-open-check-zz9", files=("plain.xlsx",), age=HOURS_AGO_2),
        _plant(temp, "someone-elses-folder", age=HOURS_AGO_2),
    ]
    sub = _plant(temp, f"ks4xl-save-{dead}-subdir12", age=HOURS_AGO_2)
    (sub / "inner").mkdir()
    (temp / f"ks4xl-save-{dead}-afile123").write_bytes(b"a file, not a folder")
    before = sorted(os.listdir(temp))
    contents = {d.name: sorted(os.listdir(d)) for d in kept + [sub]}
    _com_save(book)
    after = sorted(os.listdir(temp))
    assert after == before, (before, after)
    assert {d.name: sorted(os.listdir(d)) for d in kept + [sub]} == contents
    assert (sub / "inner").is_dir() and (sub / "book.xlsx").exists()


@pytest.mark.skipif(not IS_WIN, reason="junctions are Windows")
def test_a_junction_with_the_prefix_is_never_followed(book, temp, tmp_path):
    dead = _dead_pid()
    target = tmp_path / "someone-elses-work"
    target.mkdir()
    (target / "book.xlsx").write_bytes(b"their workbook")
    link = temp / f"ks4xl-save-{dead}-junction"
    subprocess.run(["cmd", "/c", "mklink", "/J", str(link), str(target)], check=True,
                   capture_output=True, stdin=subprocess.DEVNULL, creationflags=NO_WINDOW)
    try:
        _com_save(book)
        assert (target / "book.xlsx").read_bytes() == b"their workbook"
    finally:
        if link.exists():
            os.rmdir(link)


def test_a_save_makes_no_folder_in_the_temp_area(book, temp, monkeypatch):
    """Excel's copy is written beside the workbook, so a save makes no
    folder in the temp area at all; the sweep remains for the folders
    earlier versions' saves left there (ks4xl-save-<pid>-<random>, the
    owner judged from the name)."""
    made = []
    real = tempfile.mkdtemp

    def spy(*a, **k):
        d = real(*a, **k)
        made.append(os.path.basename(d))
        return d

    monkeypatch.setattr(tempfile, "mkdtemp", spy)
    _com_save(book)
    assert not [m for m in made if m.startswith("ks4xl-save-")], made
