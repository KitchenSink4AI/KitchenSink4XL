"""core/access.py: why a workbook cannot be written, named by its real cause.

A refusal that says "open in Excel" sends the user to close a program they
never opened when the real cause is the file's read-only attribute, its
permissions, or some other program holding it. The COM routes used to say
"open in Excel or locked by another process" for all of them; the file tier
already named the read-only attribute and now names the other two as well.

Two probes:

- the PLAIN probe opens the file read-write while sharing writing, which is
  what the file tier needs (it writes a temp and replaces the file);
- the DENY-WRITERS probe opens it the way Excel's own read-write open does,
  read and write with other writers refused. A program that holds the file
  open for writing and shares writing (some editors and sync tools, with no
  ~$ owner file) passes the plain probe; Excel then opens the workbook
  read-only and its save puts a "Save As" dialog on the screen.
  The COM routes use this probe before Excel starts.

Both probes also ask for DELETE access. Saving replaces the
file: Excel writes a new copy and swaps it in, and the file tier does the
same with its own temp. Neither can happen while another program holds the
file without sharing delete, which is the C runtime's default (fopen, and so
Python's open(path, "rb"), many viewers and scripts). Such a reader passed
both probes because it does not write; Excel then edited the workbook, could
not replace the original and put a visible "Save As" dialog on the screen,
and the file tier failed at the replace with raw system text. The DELETE
request makes the probe fail with a sharing violation for exactly that
holder, while a reader that shares delete (where the replace succeeds) and
no holder at all still pass (measured).
"""

from __future__ import annotations

import os
import stat
import sys
from pathlib import Path

READ_ONLY_ATTRIBUTE = "read_only_attribute"
IN_USE = "in_use"
ACCESS_DENIED = "access_denied"

_ERROR_ACCESS_DENIED = 5


def owner_file(path: str | os.PathLike) -> Path:
    """Excel's owner file (~$name) beside a workbook it has open."""
    p = Path(path)
    return p.parent / ("~$" + p.name)


def _cause_of(exc: OSError) -> str:
    winerror = getattr(exc, "winerror", None)
    if winerror is not None:
        return ACCESS_DENIED if winerror == _ERROR_ACCESS_DENIED else IN_USE
    return ACCESS_DENIED if isinstance(exc, PermissionError) else IN_USE


_KERNEL32 = None


def long_path(path: str) -> str:
    """`path` absolute, in the long-path form when it is over 240
    characters: \\\\?\\C:\\... for a drive path, \\\\?\\UNC\\server\\share\\...
    for a network path (a network path given the drive form named no file at
    all, so a 252-character UNC path was refused as "in use")."""
    target = os.path.abspath(path)
    if len(target) > 240 and not target.startswith("\\\\?\\"):
        if target.startswith("\\\\"):
            target = "\\\\?\\UNC\\" + target[2:]
        else:
            target = "\\\\?\\" + target
    return target


def _open_for_write(path: str, *, share_write: bool) -> int:
    """0 when `path` opens for read, write and delete; otherwise the Windows
    error code (32 sharing violation, 5 denied). share_write False refuses
    other writers, as Excel's read-write open does; True shares writing, as
    Python's own open(..., "r+b") does. DELETE is asked for in both, because
    every save replaces the file: a holder that does not share
    delete fails the probe. Python's open cannot be used to tell the causes
    apart: the C runtime maps both errors to the same EACCES."""
    import ctypes
    from ctypes import wintypes

    global _KERNEL32
    if _KERNEL32 is None:
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD,
                                    wintypes.DWORD, ctypes.c_void_p,
                                    wintypes.DWORD, wintypes.DWORD,
                                    wintypes.HANDLE]
        k32.CreateFileW.restype = wintypes.HANDLE
        k32.CloseHandle.argtypes = [wintypes.HANDLE]
        k32.CloseHandle.restype = wintypes.BOOL
        _KERNEL32 = k32
    target = long_path(path)
    generic_read, generic_write, delete = 0x80000000, 0x40000000, 0x00010000
    file_share_read, file_share_write, file_share_delete = 0x1, 0x2, 0x4
    open_existing, attribute_normal = 3, 0x80
    share = (file_share_read | file_share_delete
             | (file_share_write if share_write else 0))
    h = _KERNEL32.CreateFileW(target, generic_read | generic_write | delete,
                              share, None, open_existing, attribute_normal,
                              None)
    if h is None or h == ctypes.c_void_p(-1).value:
        return ctypes.get_last_error() or _ERROR_ACCESS_DENIED
    _KERNEL32.CloseHandle(h)
    return 0


def write_block(path: str | os.PathLike, *, deny_writers: bool = False
                ) -> str | None:
    """None when `path` can be opened for writing; otherwise the cause:
    READ_ONLY_ATTRIBUTE, IN_USE or ACCESS_DENIED. A missing or unreadable
    file returns None (the caller's own open reports it).

    The attribute is checked first because it is a property of the file
    rather than a race: st_mode's owner-write bit is exactly the read-only
    attribute on Windows and the ordinary permission bit elsewhere."""
    try:
        mode = os.stat(path).st_mode
    except OSError:
        return None
    if not mode & stat.S_IWRITE:
        return READ_ONLY_ATTRIBUTE
    if sys.platform == "win32":
        err = _open_for_write(os.fspath(path), share_write=not deny_writers)
        if err == 0:
            return None
        return ACCESS_DENIED if err == _ERROR_ACCESS_DENIED else IN_USE
    try:
        with open(path, "r+b"):
            pass
    except OSError as exc:
        return _cause_of(exc)
    return None


_TAIL = ("Every edit needs the file closed, xl_app's too; reads, exports and "
         "copies use the copy last saved to disk")


def in_use_text(path: str | os.PathLike) -> str:
    """Another program holds the file. "Open in Excel" only when Excel's
    owner file says so."""
    name = Path(path).name
    if owner_file(path).exists():
        return (f"{name} is open in Excel (or another process holds it); "
                "close it there, or wait for the other process to let it go, "
                "then retry. " + _TAIL)
    return (f"{name} is in use by another program, which holds it open and "
            "does not let others write to it; close it there, or wait for "
            "that program to let it go, then retry. " + _TAIL)


def access_denied_text(path: str | os.PathLike) -> str:
    name = Path(path).name
    return (f"{name}: this account is not allowed to change the file (its "
            "permissions, or its folder's, do not allow writing). No program "
            "is holding it. Ask whoever manages the file for write access, "
            "or save a copy elsewhere with copy_workbook")


def refusal_text(path: str | os.PathLike, cause: str) -> str:
    """The refusal sentence for a write_block cause."""
    if cause == READ_ONLY_ATTRIBUTE:
        from .package import read_only_refusal
        return read_only_refusal(os.fspath(path))
    if cause == ACCESS_DENIED:
        return access_denied_text(path)
    return in_use_text(path)


def read_denied(path: str | os.PathLike) -> bool:
    """True when the file's permissions refuse this account reading it (as
    opposed to another program holding it, or a lock on part of it). Only
    Windows can tell: an open for reading that shares everything fails with
    5 (denied) for permissions and 32 or 33 for a holder. Elsewhere the
    permission bits answer (os.access)."""
    if sys.platform != "win32":
        return not os.access(path, os.R_OK)
    import ctypes
    from ctypes import wintypes
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD,
                                wintypes.DWORD, ctypes.c_void_p,
                                wintypes.DWORD, wintypes.DWORD,
                                wintypes.HANDLE]
    k32.CreateFileW.restype = wintypes.HANDLE
    h = k32.CreateFileW(long_path(os.fspath(path)), 0x80000000, 0x7, None,
                        3, 0x80, None)
    if h is None or h == ctypes.c_void_p(-1).value:
        return ctypes.get_last_error() == _ERROR_ACCESS_DENIED
    k32.CloseHandle(h)
    return False


def folder_refuses_text(path: str | os.PathLike) -> str:
    """The folder does not let this account create the new version beside
    the workbook (it used to read "it may be open in Excel" with the raw
    system error)."""
    p = Path(path)
    return (f"{p.name} was not changed: this account cannot create files in "
            f"its folder ({p.parent}), and every edit first writes the new "
            "version there and then swaps it in. Ask whoever manages the "
            "folder for permission to create files in it, or save a copy "
            "elsewhere with copy_workbook")


def _folder_takes_new_files(folder: str) -> bool:
    """One create attempt. Not tempfile.mkstemp: on Windows it takes the
    permission error of a folder that refuses new files for a name clash
    and retries without end (measured: the call hung)."""
    import uuid
    probe = os.path.join(folder, f".ks4xl-write-{uuid.uuid4().hex}.probe")
    try:
        fd = os.open(probe, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except OSError:
        return False
    os.close(fd)
    try:
        os.remove(probe)
    except OSError:
        pass
    return True


def write_refusal_text(path: str | os.PathLike) -> str:
    """Writing the new version beside the workbook failed with an access
    error: name the cause (the folder, the file's own state, or a program
    that held it for a moment), never "may be open in Excel" with raw
    system text."""
    p = Path(path)
    if not _folder_takes_new_files(str(p.parent)):
        return folder_refuses_text(p)
    return replace_refusal_text(p)


def replace_refusal_text(path: str | os.PathLike,
                         winerror: int | None = None) -> str:
    """Swapping the new version in failed with an access error: the file's
    own state or a program holding it, by the probes; when they pass again,
    a program held it for a moment.

    Error 5 (access denied) is not a holder: the save already made the swap
    once more ignoring permission errors, so what is left is the folder
    refusing the replace, or a program that held the file for a moment. The answer says both and never claims a program had it open."""
    cause = write_block(path)
    if cause is not None:
        return refusal_text(path, cause)
    name = Path(path).name
    if winerror == 5:
        return (f"{name} was not changed: Windows denied replacing it (error "
                "5). This folder's permissions may not allow replacing files, "
                "or a program held the file for a moment. Retry; if it happens "
                "again, check that this account may change files in this "
                "folder, or save a copy elsewhere with copy_workbook.")
    return (f"{name} was not changed: another program had the file open at "
            "the moment of the save and has let it go since. Retry.")


def read_only_open_text(path: str | os.PathLike) -> str:
    """Excel was asked to open the workbook for editing and could open it
    only read-only: another program got it first."""
    name = Path(path).name
    return (f"{name} could be opened only read-only, because another program "
            "holds it open for writing, so Excel could not save it. Nothing "
            "was changed. Close it in the other program, or wait for that "
            "program to let it go, then retry. " + _TAIL)


__all__ = ["READ_ONLY_ATTRIBUTE", "IN_USE", "ACCESS_DENIED", "owner_file",
           "write_block", "in_use_text", "access_denied_text", "refusal_text",
           "read_only_open_text", "long_path", "read_denied",
           "folder_refuses_text", "write_refusal_text", "replace_refusal_text"]
