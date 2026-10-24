"""Saves in a folder where the user may write files but not change their
permissions.

The folder grants the user Modify and OWNER RIGHTS (S-1-3-4) Modify, with
the folder's own inheritance removed: a common file-share hardening.
ReplaceFileW with flags 0 must write the replaced file's permissions onto
the new file, OWNER RIGHTS takes that right away from the file's owner, and
the call fails with ERROR_ACCESS_DENIED (5). Every save, restore and
overwrite was refused there with "another program had the file open", which
was not true; os.replace and ReplaceFileW with REPLACEFILE_IGNORE_ACL_ERRORS
(0x4) both succeed. The swap is now made once more with 0x4, and a workbook
whose own permissions could not come across is named in a warning.
"""
from __future__ import annotations

import shutil
import subprocess
import sys

import openpyxl
import pytest

from xlsx_mcp.core import access, safesave
from xlsx_mcp.core.errors import WorkbookLocked
from xlsx_mcp.ops import backups, cells, lifecycle

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Windows ACLs")
NOWIN = 0x08000000


def _sid() -> str:
    out = subprocess.run(["whoami", "/user", "/fo", "csv", "/nh"], capture_output=True,
                         text=True, creationflags=NOWIN, timeout=30).stdout.strip()
    return out.split(",")[-1].strip('"')


def _icacls(*args: str) -> None:
    r = subprocess.run(["icacls", *args], capture_output=True, text=True,
                       creationflags=NOWIN, timeout=60)
    assert r.returncode == 0, r.stdout + r.stderr


def _elevated() -> bool:
    import ctypes
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:  # noqa: BLE001
        return False


def _own_entries(p) -> list:
    """The file's permission entries that are not inherited from its folder."""
    r = subprocess.run(["icacls", str(p)], capture_output=True, text=True,
                       creationflags=NOWIN, timeout=60)
    lines = [ln.strip() for ln in r.stdout.splitlines() if ln.strip() and "Successfully" not in ln]
    lines[0] = lines[0].replace(str(p), "").strip()
    return [ln for ln in lines if "(I)" not in ln]


def _lock_down(folder) -> None:
    _icacls(str(folder), "/inheritance:r", "/grant:r", f"*{_sid()}:(OI)(CI)M",
            "*S-1-3-4:(OI)(CI)M")


@pytest.fixture
def owner_rights_folder(tmp_path):
    folder = tmp_path / "share"
    folder.mkdir()
    _lock_down(folder)
    yield folder
    shutil.rmtree(folder, ignore_errors=True)  # Modify includes delete


def _book(p, v):
    wb = openpyxl.Workbook()
    wb.active["A1"] = v
    wb.save(p)


def _a1(p):
    wb = openpyxl.load_workbook(p, read_only=True)
    try:
        return wb.active["A1"].value
    finally:
        wb.close()


def test_set_cell_saves_in_an_owner_rights_folder(owner_rights_folder):
    p = str(owner_rights_folder / "Book.xlsx")
    _book(p, "before")
    r = cells.set_cell(p, "A1", "after")
    assert r["ok"] is True
    assert _a1(p) == "after"
    assert not any("permissions" in w for w in r.get("warnings", [])), (
        "a workbook with only the folder's permissions lost nothing; no warning is due")


def test_copy_workbook_overwrites_in_an_owner_rights_folder(owner_rights_folder, tmp_path):
    src = str(tmp_path / "src.xlsx")
    _book(src, "copied")
    dst = str(owner_rights_folder / "Dest.xlsx")
    _book(dst, "old")
    r = lifecycle.copy_workbook(src, dst, overwrite=True)
    assert r["saved"] is True
    assert _a1(dst) == "copied"


def test_restore_works_in_an_owner_rights_folder(owner_rights_folder):
    p = str(owner_rights_folder / "Book.xlsx")
    _book(p, "v1")
    cells.set_cell(p, "A1", "v2")
    r = backups.manage_backups("restore", path=p, source="prev")
    assert r["restored"]
    assert _a1(p) == "v1"


@pytest.mark.skipif(_elevated(), reason="an elevated account holds the right to change "
                    "permissions through the Administrators entry its temp folder hands down, "
                    "so this lockdown does not meet it as written")
def test_a_workbook_whose_own_permissions_could_not_come_across_is_named(tmp_path):
    folder = tmp_path / "share"
    folder.mkdir()
    p = folder / "Book.xlsx"
    _book(p, "before")
    # The workbook's own explicit entry, set while the user still may.
    _icacls(str(p), "/grant", f"*{_sid()}:(R)")
    _lock_down(folder)
    try:
        r = cells.set_cell(str(p), "A1", "after")
        assert _a1(p) == "after"
        assert any("Book.xlsx was saved, but its own permissions could not be kept"
                   in w for w in r["warnings"]), r["warnings"]
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def test_access_denied_is_made_once_more_ignoring_permission_errors(tmp_path, monkeypatch):
    p = tmp_path / "Book.xlsx"
    _book(p, "before")
    if _own_entries(p):
        pytest.skip("this account's new files carry explicit permissions of their own "
                    "(an elevated account), so the workbook is named in a warning")
    new = tmp_path / "new.xlsx"
    _book(new, "after")
    fn, move, last_error, text = safesave._replace_file()
    calls = []

    def replace(dst, src, bak, flags, a, b):
        calls.append(flags)
        if flags == 0:
            return 0  # denied, as in an OWNER RIGHTS folder
        return fn(dst, src, bak, flags, a, b)

    monkeypatch.setattr(safesave, "_REPLACE_FILE",
                        (replace, move, lambda: 5 if calls[-1] == 0 else last_error(), text))
    assert safesave.promote(str(new), str(p)) == []
    assert calls == [0, 0x4]
    assert _a1(p) == "after"


def test_a_denied_replace_is_never_called_another_programs_hold(tmp_path, monkeypatch):
    p = tmp_path / "Book.xlsx"
    _book(p, "before")
    fn, move, last_error, text = safesave._replace_file()
    monkeypatch.setattr(safesave, "_REPLACE_FILE",
                        (lambda *a: 0, move, lambda: 5, text))
    with pytest.raises(WorkbookLocked) as info:
        cells.set_cell(str(p), "A1", "after")
    msg = str(info.value)
    assert "another program had the file open" not in msg
    assert "denied replacing it (error 5)" in msg
    assert _a1(p) == "before"
    assert access.replace_refusal_text(str(p), 32).endswith(
        "another program had the file open at the moment of the save and has let it go "
        "since. Retry.")
