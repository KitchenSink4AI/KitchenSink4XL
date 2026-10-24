"""com/session.py: the serialized, timeout-bounded COM execution layer.

Live-COM field testing proved that unserialized concurrent
COM access produces silent data corruption, interleaved writes, modal-dialog
deadlocks, and unbounded hangs. KS4XL's COM tier therefore ships with, from
the first commit:

  1. PROCESS-WIDE SERIALIZATION: every COM entry point routes through ONE
     dedicated worker thread (ComExecutor). One thread means one COM apartment
     and structurally serial operations; there is no lock to forget.
  2. ALERT SUPPRESSION: DisplayAlerts=False and EnableEvents=False are set at
     worker spawn AND re-asserted around every operation (Excel can flip
     DisplayAlerts back after some calls); the prior value is restored after.
     Workers are private DispatchEx instances, so no user setting is touched.
  3. BOUNDED TIMEOUTS: every submitted operation gets a deadline (default 60s,
     KS4XL_COM_TIMEOUT tunable, per-call override). On timeout the worker
     Excel is reclaimed BY OWNED PID (which unblocks the stuck COM call), the
     executor re-arms with a fresh worker, and the caller gets a clean
     structured APP_BLOCKED refusal, never a hang. THE DEADLINE IS THE WHOLE
     WAIT: a caller asking for N seconds gets its refusal within N plus the
     taskkill, a fraction of a second. The eight-second restart-transient
     reap that used to run before the refusal was raised now runs on a
     background thread that shutdown joins.
  4. RETRY WITH BACKOFF: save/export-class calls retry on the documented
     busy/rejected HRESULTs (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER,
     VBA_E_IGNORE) with exponential backoff before giving up as APP_BUSY.
  5. HONEST STATUS: status() reports busy/contention/current-op/last-op truth,
     not a hardcoded "ready".

POLICY (the [EXCEL CLOSED] vs live-coexistence tension, resolved for v1):
the COM tier uses PRIVATE POOLED WORKER INSTANCES ONLY. It never calls
GetActiveObject onto a user's Excel session and never routes an edit into a
visible instance. When the target file is held open by any other process
(the user's Excel included), com_ operations refuse with WORKBOOK_LOCKED
semantics instead of touching the open copy. A stale ~$ lockfile left by a
crash degrades to a warning after a real write-probe, never a permanent
spurious refusal.

PID accounting: the ExcelInstanceManager journals every spawned PID; a
startup sweep reclaims stale journal-owned PIDs from crashed prior sessions;
shutdown quits the pool. Foreign EXCEL.EXE processes are never touched.
"""

from __future__ import annotations

import atexit
import contextvars
import os
import queue
import re
import threading
import time
from dataclasses import dataclass, field
from typing import Any, Callable

from ..core.errors import (
    ExcelBlocked,
    ExcelBusy,
    ExcelNotRunning,
    ValidationFailed,
    WorkbookLocked,
    XlMcpError,
)
from .instances import ExcelInstanceManager

DEFAULT_TIMEOUT = float(os.environ.get("KS4XL_COM_TIMEOUT", "60"))
#: Journal records older than this from a PRIOR session are treated as crash
#: leftovers and reclaimed at startup (by owned PID only).
STALE_JOURNAL_SECONDS = 900.0

#: Busy/rejected HRESULTs worth a backoff retry (signed 32-bit forms).
_RETRYABLE_HRESULTS = {
    -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED
    -2147417846,  # 0x8001010A RPC_E_SERVERCALL_RETRYLATER
    -2146777998,  # 0x800AC472 VBA_E_IGNORE (Excel busy / cell edit mode)
}


# --------------------------------------------------------------- availability

_AVAILABLE: tuple[bool, str] | None = None


def com_available(*, refresh: bool = False) -> tuple[bool, str]:
    """Cheap cached probe: (available, reason). Never spawns Excel."""
    global _AVAILABLE
    if _AVAILABLE is not None and not refresh:
        return _AVAILABLE
    import sys

    # sys.platform, not platform.system(): on Windows the latter reaches
    # platform._wmi_query, which can hang with no bound.
    if sys.platform != "win32":
        _AVAILABLE = (False, "COM requires Windows")
        return _AVAILABLE
    try:
        import win32com.client  # noqa: F401
    except Exception:
        _AVAILABLE = (False, "pywin32 is not installed")
        return _AVAILABLE
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                            r"Excel.Application\CLSID"):
            pass
    except OSError:
        _AVAILABLE = (False, "Excel is not installed (no COM registration)")
        return _AVAILABLE
    _AVAILABLE = (True, "Windows + pywin32 + Excel COM registration present")
    return _AVAILABLE


def require_com() -> None:
    ok, reason = com_available()
    if not ok:
        raise ExcelNotRunning(
            f"this tool drives the Excel application and cannot run here: "
            f"{reason}. The file-based tools keep working without Excel.")


# ------------------------------------------------------------ target guard


def guard_target_closed(path: str) -> list[str]:
    """Refuse when the target file cannot be opened for editing the way
    Excel opens it; degrade a STALE ~$ lockfile to a warning after a real
    write-probe. Returns advisory warnings. POLICY: com_ operations only ever
    open files in private pooled workers, never a user's open copy.

    The probe refuses other writers, as Excel's own read-write open does.
    The old probe shared writing, so a program holding the file
    open for writing and sharing it (some editors and sync tools, with no ~$
    owner file) passed; Excel then opened the workbook read-only and its save
    put a visible "Save As" dialog on the screen. The refusal names the real
    cause: the read-only attribute, the file's permissions, or another
    program ("open in Excel" only when Excel's owner file is there)."""
    from pathlib import Path

    from ..core import access as _access

    warnings: list[str] = []
    p = Path(path)
    lock_present = _access.owner_file(p).exists()
    if p.exists():
        cause = _access.write_block(p, deny_writers=True)
        if cause is not None:
            raise WorkbookLocked(_access.refusal_text(p, cause))
    if lock_present:
        warnings.append(
            f"a stale owner lockfile (~${p.name}) is present but the file is "
            "writable; a prior Excel session likely crashed. Proceeding; you "
            "can delete the ~$ file safely.")
    return warnings


# ----------------------------------------------------------------- plumbing


def com_retry(fn: Callable[[], Any], *, attempts: int = 4,
              base_delay: float = 0.5, label: str = "COM call") -> Any:
    """Run fn, retrying on the busy/rejected HRESULTs with exponential
    backoff. Raises ExcelBusy after the last attempt."""
    last: Exception | None = None
    for attempt in range(attempts):
        try:
            return fn()
        except Exception as exc:  # noqa: BLE001
            hr = getattr(exc, "hresult", None)
            args = getattr(exc, "args", ())
            code = hr if hr is not None else (args[0] if args else None)
            if code in _RETRYABLE_HRESULTS and attempt < attempts - 1:
                last = exc
                time.sleep(base_delay * (2 ** attempt))
                continue
            raise
    raise ExcelBusy(
        f"{label}: Excel stayed busy across {attempts} attempts "
        f"({type(last).__name__ if last else 'unknown'})")


#: The HRESULTs Excel actually hands back, in plain English. Without this a
#: refusal read "com_error (HRESULT -2147352567)" -- a pywin32 repr inside an
#: otherwise clean envelope, which tells a caller nothing about what to do.
_HRESULT_TEXT = {
    -2147352567: ("Excel rejected the call and supplied no message of its "
                  "own. That is what Excel returns when the arguments are "
                  "not valid for the current state of the workbook (a range "
                  "that does not fit the operation, a cell that holds no "
                  "formula, a value it cannot use)"),
    -2147418111: ("Excel was busy and rejected the call; it is running a "
                  "command or showing a dialog"),
    -2147417846: "Excel asked the caller to retry later; it stayed busy",
    -2146777998: ("Excel is in cell-edit mode or otherwise ignoring "
                  "automation; press Escape in Excel and retry"),
    -2147221164: "the Excel COM class is not registered on this machine",
    -2147023174: ("the Excel process stopped answering and the call could "
                  "not be delivered (RPC server unavailable): the worker "
                  "Excel is gone"),
    -2147023170: ("the Excel process died in the middle of this call (the "
                  "remote procedure call failed), so the operation did not "
                  "finish. This is what a crashed or force-closed Excel "
                  "returns"),
    -2147023169: ("the Excel process is no longer there to answer the call "
                  "(the remote procedure endpoint does not exist)"),
    -2147417848: ("the Excel object this call was made on has disconnected "
                  "from its client; the workbook or the process closed "
                  "underneath the operation"),
    -2147221019: ("the Excel object is no longer connected to a running "
                  "process; Excel closed underneath the operation"),
}

#: The HRESULTs that mean EXCEL IS GONE, as opposed to Excel refusing a call.
#: A taskkilled or crashed Excel returns 0x800706BE mid-body. These used to
#: land on BAD_PARAMS: the caller was told its arguments were wrong when in
#: fact the process had died, so an unattended caller would "fix" a
#: correct call instead of retrying. They now raise ExcelDisconnected,
#: which the envelope already maps to CONFLICT.
_PROCESS_GONE_HRESULTS = frozenset({
    -2147023174,  # 0x800706BA RPC_S_SERVER_UNAVAILABLE
    -2147023170,  # 0x800706BE RPC_S_CALL_FAILED
    -2147023169,  # 0x800706BF RPC_S_CALL_FAILED_DNE
    -2147417848,  # 0x80010108 RPC_E_DISCONNECTED
    -2147221019,  # 0x800401FD CO_E_OBJNOTCONNECTED
})


def _hresult_of(exc: Exception) -> int | None:
    """The HRESULT a pywin32 com_error carries, on either of the two shapes
    it uses (the .hresult attribute, or args[0] on the tuple form)."""
    hr = getattr(exc, "hresult", None)
    if isinstance(hr, int):
        return hr
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0]
    return None


def excel_process_died(exc: Exception) -> bool:
    """True when this COM failure means the Excel PROCESS went away mid-call
    (killed, crashed, or closed underneath us), rather than Excel refusing
    what it was asked to do.

    pywin32 reports a member lookup on an object whose Excel is gone as an
    AttributeError ("Open.SaveCopyAs") raised while handling the COM error;
    that error's HRESULT decides (it used to be answered BAD_PARAMS, which
    tells an assistant its arguments were wrong)."""
    if _hresult_of(exc) in _PROCESS_GONE_HRESULTS:
        return True
    if isinstance(exc, AttributeError):
        inner = exc.__context__ or exc.__cause__
        return inner is not None and _hresult_of(inner) in _PROCESS_GONE_HRESULTS
    return False


def excel_gone(pid: int, wait: float = 0.0) -> bool:
    """True when the Excel process ``pid`` has ended, waiting up to ``wait``
    seconds for one that is going. A COM call into an Excel that died can
    fail as a plain member lookup (pywin32 raises AttributeError("Open.Close")
    with no COM error behind it, MEASURED), so the process itself is the
    answer, not the exception."""
    from .instances import pid_alive
    deadline = time.monotonic() + max(0.0, wait)
    while True:
        if not pid_alive(pid):
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.1)


def excel_error_message(exc: Exception) -> str:
    """A human-facing message for a pywin32 com_error: Excel's own text when
    present, else a plain-English rendering of the HRESULT. Never a raw
    pywin32 repr."""
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
        return str(excepinfo[2]).strip()
    hr = _hresult_of(exc)
    if hr is not None:
        known = _HRESULT_TEXT.get(hr)
        if known:
            return known
        return (f"Excel rejected the call and supplied no message "
                f"(error code 0x{hr & 0xFFFFFFFF:08X})")
    return f"{type(exc).__name__}: {exc}"


#: OLE compound-file (CFB) signature: an ENCRYPTED OOXML workbook is a CFB
#: container, not a zip. A password-less open of one cannot be made to fail
#: fast (Excel prompts modally whatever arguments are supplied, and the call
#: hangs until the worker is killed), so encryption is detected by signature
#: BEFORE Excel is asked to open anything without a password. A SUPPLIED
#: password is a different matter: since the switch to positional delivery a
#: wrong one raises Excel's own wrong-password error in about a tenth of a
#: second (measured).
_CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def is_encrypted_package(path: str) -> bool:
    """True when the file is an OLE/CFB container that is (or could be)
    Excel's real encryption. A legacy BIFF .xls is ALSO a CFB container, and
    the bare 8-byte magic check used to call every genuine 97-2003 workbook
    "password-protected", refusing the one COM route (com_convert_format)
    that can actually help its owner. The stream-name
    sniff (core.hazard.ole_container_kind) tells the two apart; an OLE
    container the sniff cannot place stays treated as possibly-encrypted,
    because a password-less open of a real encrypted file hangs Excel on a
    modal prompt and that risk wins."""
    from ..core import hazard as _hazard
    kind = _hazard.ole_container_kind(path)
    return kind in (_hazard.OLE_KIND_ENCRYPTED, _hazard.OLE_KIND_UNKNOWN)


#: Passed as the Open password when the caller supplied none. Unencrypted
#: files ignore it; encrypted files never reach Open without a password
#: (the CFB pre-check refuses first), so this is defense in depth only.
_NO_PASSWORD_SENTINEL = "ks4xl.no-password.sentinel.7f3a9c"

#: Passed as WriteResPassword on every read-write open. A
#: workbook with a password to modify then fails the Open at once with
#: Excel's own wrong-password error instead of showing its Password dialog,
#: which DisplayAlerts does not suppress. MEASURED (Excel 16, one
#: open per process on a private desktop): plain .xlsx and .xlsb files and
#: an encrypted file with no password to modify open normally with it; a
#: password to modify on an .xlsx, .xlsb, .xls or encrypted .xlsx fails in
#: 0.08 to 0.19 s with no dialog. The file check in core.hazard still refuses
#: every form it can read before Excel starts; this catches what it cannot
#: see, chiefly a password to modify inside Excel's encryption.
#:
#: AT MOST 15 CHARACTERS, Excel's limit for a password to modify. MEASURED:
#: a 15-character stand-in behaves as above; a 16-character or longer one
#: fails EVERY open outright ("Open method of Workbooks class failed"),
#: plain files included. That is what an earlier note recorded with a
#: 33-character sentinel, and why no stand-in was used until now.
_NO_WRITE_PASSWORD_SENTINEL = "ks4xl-nw-5e1b"

#: Whether this Excel opens a workbook with no password to modify when
#: handed the stand-in. None until measured. It is measured only when a
#: read-write open with the stand-in fails on a file that opens read-only,
#: because that failure means either "this file has a password to modify"
#: or "this Excel refuses the stand-in on every file", and only a plain
#: workbook of our own tells the two apart. The Excel it was measured on
#: accepts it; the measurement is the safety net if another build does not.
_WRITE_STAND_IN_OK: bool | None = None


def _open(app, abs_path: str, *, read_only: bool, password: str,
          write_password: str | None):
    """The one Workbooks.Open call. POSITIONAL arguments (see
    open_workbook).

    A read-write open also passes Editable=True, the tenth argument.
    Without it Excel opens a template (.xltx, .xltm) as a NEW,
    unsaved workbook based on it ("Budget1"), the route edited that copy,
    and wb.Save wrote it as Budget1.xlsx into Excel's current folder (the
    user's Documents), silently replacing any file of that name, while the
    template stayed untouched and the call answered ok. MEASURED (Excel 16,
    private desktop): with Editable=True a template opens as itself (its
    own path, template format) and its save rewrites it in place; .xlsx,
    .xlsm, .xlsb, read-only-recommended and encrypted files open exactly as
    before; a password to modify still fails at once with the stand-in.
    Origin and Delimiter are passed as None: pythoncom.Missing there cuts
    the argument list short and Editable never reaches Excel (measured: the
    template opened as a copy again)."""
    args: tuple = (
        abs_path,        # FileName
        0,               # UpdateLinks: never
        read_only,       # ReadOnly
        None,            # Format
        password,        # Password (positional delivery is required)
        write_password,  # WriteResPassword: the stand-in, or None
        True,            # IgnoreReadOnlyRecommended
    )
    if not read_only:
        args += (
            None,        # Origin (text files only)
            None,        # Delimiter (text files only)
            True,        # Editable: a template opens as itself
        )
    wb = com_retry(lambda: app.Workbooks.Open(*args),
                   label=f"open {os.path.basename(abs_path)}")
    _note_opened(app, wb, abs_path)
    return wb


def _note_opened(app, wb, *also: str) -> None:
    """Record a workbook this server opened in its Excel, so a workbook
    there that was NOT recorded is known to be the user's (instances.
    users_claim): the pool then hands the instance over instead of closing
    their work."""
    from .instances import note_opened
    try:
        note_opened(app, str(getattr(wb, "FullName", "") or ""))
    except Exception:  # noqa: BLE001
        pass
    for name in also:
        note_opened(app, name)


def _close_quietly(wb) -> None:
    try:
        wb.Close(False)
    except Exception:  # noqa: BLE001
        pass


def _stand_in_ok(app) -> bool | None:
    """Measure, once per process, whether this Excel opens a workbook that
    has no password to modify when handed the stand-in. A plain workbook is
    made in a private temp folder by the same hidden Excel, then opened with
    the stand-in. None when the measurement itself failed (nothing is kept,
    so a later open measures again)."""
    global _WRITE_STAND_IN_OK
    if _WRITE_STAND_IN_OK is not None:
        return _WRITE_STAND_IN_OK
    import shutil
    import tempfile

    folder = tempfile.mkdtemp(prefix="ks4xl-open-check-")
    probe = os.path.join(folder, "plain.xlsx")
    wb = None
    try:
        try:
            wb = app.Workbooks.Add()
            _note_opened(app, wb, probe)
            com_retry(lambda: wb.SaveAs(probe, 51), label="save a check workbook")
            _close_quietly(wb)
            wb = None
        except Exception:  # noqa: BLE001
            return None
        try:
            wb = _open(app, probe, read_only=False,
                       password=_NO_PASSWORD_SENTINEL,
                       write_password=_NO_WRITE_PASSWORD_SENTINEL)
            _WRITE_STAND_IN_OK = True
        except Exception as exc:  # noqa: BLE001
            if excel_process_died(exc):
                raise
            _WRITE_STAND_IN_OK = False
        return _WRITE_STAND_IN_OK
    finally:
        if wb is not None:
            _close_quietly(wb)
        shutil.rmtree(folder, ignore_errors=True)


def _diagnose_refused_write_open(app, abs_path: str, pw: str, exc: Exception):
    """A read-write open handed the stand-in failed. Name the cause, or
    return the workbook opened the old way when that is safe.

    1. A read-only open never asks about a password to modify. If it fails
       too, the file itself did not open (a wrong current password, a
       damaged file): that error is the answer, raised as it came.
    2. It opens read-only, so the read-write failure was the stand-in: the
       file has a password to modify, unless this Excel refuses the
       stand-in on every file. A plain workbook of our own tells which.
    3. This Excel refuses the stand-in everywhere: an OOXML package was
       already cleared by the file check, so it opens the old way; anything
       the file check cannot read inside (Excel's encryption, a legacy
       container) cannot be shown to have no password to modify, so it
       is refused rather than risk the dialog."""
    from ..core import hazard as _hazard

    probe = _open(app, abs_path, read_only=True, password=pw,
                  write_password=None)
    _close_quietly(probe)
    ok = _stand_in_ok(app)
    if ok:
        raise _hazard.write_reserved_refusal(abs_path, excel_started=True) from exc
    if ok is False and _hazard.ole_container_kind(abs_path) is None:
        return _open(app, abs_path, read_only=False, password=pw,
                     write_password=None)
    raise _hazard.write_reserved_refusal(abs_path, unreadable=True,
                                         excel_started=True) from exc


def _refuse_read_only_open(wb, abs_path: str) -> None:
    """Excel was asked for a read-write open and gave a read-only one:
    another program got the file first. Saving it would put
    Excel's "Save As" dialog on the screen, so it is closed unsaved."""
    try:
        read_only = bool(getattr(wb, "ReadOnly", False))
    except Exception:  # noqa: BLE001
        read_only = True
    if read_only:
        from ..core import access as _access
        _close_quietly(wb)
        raise WorkbookLocked(_access.read_only_open_text(abs_path))


def is_target(wb, path: str) -> bool:
    """True when the workbook Excel has open IS the file at `path`, so a
    wb.Save writes that file and nothing else.

    The names are compared normalized (separators, case); when they differ
    in spelling (an 8.3 name, a SUBST or mapped drive, a junction), the
    resolved final paths decide. File identity is not used: a hard link is
    the same file under another name, and Excel's save replaces the name it
    knows, so the other name would keep the old content. A workbook in a
    folder that Office syncs (OneDrive, SharePoint) is named by its web
    address; it is the file when it has a saved location and the same name.
    A new workbook Excel made from a template has neither a location nor
    the file's name, which is the case this exists for."""
    try:
        full = str(getattr(wb, "FullName", "") or "")
    except Exception:  # noqa: BLE001
        return False
    if not full:
        return False
    target = os.path.abspath(path)
    if os.path.normcase(os.path.normpath(full)) == \
            os.path.normcase(os.path.normpath(target)):
        return True
    if full.lower().startswith(("http://", "https://")):
        try:
            name = str(getattr(wb, "Name", "") or "")
            where = str(getattr(wb, "Path", "") or "")
        except Exception:  # noqa: BLE001
            return False
        return bool(where) and \
            name.casefold() == os.path.basename(target).casefold()
    if not os.path.isabs(full):
        return False  # a workbook that has never been saved ("Budget1")
    try:
        return os.path.normcase(os.path.realpath(full)) == \
            os.path.normcase(os.path.realpath(target))
    except (OSError, ValueError):
        return False


def _other_workbook_refusal(wb, path: str):
    from ..core.errors import UnsupportedStructure
    try:
        shown = str(getattr(wb, "FullName", "") or getattr(wb, "Name", "") or "")
    except Exception:  # noqa: BLE001
        shown = ""
    name = os.path.basename(path)
    return UnsupportedStructure(
        f"Excel opened {name} as a different workbook"
        + (f" ({shown})" if shown else "")
        + ", not as the file itself, so an edit would have been saved "
        "somewhere else. Nothing was changed and nothing was saved. Excel "
        "does this with a template it will not open for editing: open the "
        "file in Excel, save it as an Excel workbook (.xlsx), and run the "
        "operation on that file.")


def _refuse_other_workbook(wb, abs_path: str) -> None:
    """The read-write open gave a workbook that is not the file: it is
    closed unsaved and refused before any edit is made."""
    if not is_target(wb, abs_path):
        exc = _other_workbook_refusal(wb, abs_path)
        _close_quietly(wb)
        raise exc


#: Excel's calculation modes (XlCalculation) and the Find arguments the
#: network check uses.
_XL_CALC_MANUAL = -4135
_XL_FORMULAS = -4123
_XL_PART = 2
#: How many of Excel's Find hits per function the look confirms one by one.
_FIND_LIMIT = 500


def _reach_in_excel(app, abs_path: str, pw: str) -> dict | None:
    """What in a workbook this server cannot read (an .xlsb, a legacy .xls,
    an encrypted workbook) can reach the network, found by Excel with
    calculation turned off.

    Excel takes its calculation mode from the first workbook it opens, and
    a workbook that asks for it is recalculated as it opens, before any
    call can stop it. MEASURED (Excel 16, private desktop, a listener on
    127.0.0.1): with a blank workbook open first and calculation set to
    manual, opening a workbook that fetches at open (WEBSERVICE, with
    fullCalcOnLoad) makes no request, and none while it is read, exported
    or closed; switching back to automatic makes the request at once. So
    the blank workbook goes first, calculation goes to manual, the file is
    opened read-only, read, and closed without calculating, and the blank
    one closes too. Nothing of either is saved. The caller then opens the
    file as usual only when nothing was found."""
    from ..core import netreach as _netreach
    anchor = None
    wb = None
    prior = None
    try:
        anchor = app.Workbooks.Add()
        _note_opened(app, anchor)
        try:
            prior = app.Calculation
        except Exception:  # noqa: BLE001
            prior = None
        app.Calculation = _XL_CALC_MANUAL
        wb = _open(app, abs_path, read_only=True, password=pw, write_password=None)
        return _reach_of_open_workbook(wb, _netreach)
    finally:
        if wb is not None:
            _close_quietly(wb)
        if anchor is not None:
            _close_quietly(anchor)
        try:
            if prior is not None and int(app.Workbooks.Count) > 0:
                app.Calculation = prior
        except Exception:  # noqa: BLE001
            pass


#: Conditional-format rule types (XlFormatConditionType) by what formula they
#: carry: the cell-value rule always has Formula1 and has Formula2 only for
#: between / not between; the expression rule has Formula1; the rule types
#: below hold no formula property that must be readable (colour scales, data
#: bars, icon sets, top-N, unique values, text, blanks, time period, average,
#: errors). Their Formula1 is still read when it is there, but a failed read
#: of it is not an unread rule.
_FC_CELL_VALUE, _FC_EXPRESSION = 1, 2
_FC_BETWEEN, _FC_NOT_BETWEEN = 1, 2
_FC_NO_REQUIRED_FORMULA = frozenset({3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 16, 17})


def _rule_formula_reads(fc) -> list[tuple[str, bool]]:
    """(property, required) pairs to read on one conditional-format rule. A
    rule whose type cannot be read, or is not a type this knows, is read in
    full and every read is required: a rule that cannot be classified cannot
    be treated as safe."""
    try:
        kind = int(fc.Type)
    except Exception:  # noqa: BLE001
        return [("Formula1", True), ("Formula2", True)]
    if kind in _FC_NO_REQUIRED_FORMULA:
        return [("Formula1", False), ("Formula2", False)]
    if kind == _FC_EXPRESSION:
        return [("Formula1", True), ("Formula2", False)]
    if kind == _FC_CELL_VALUE:
        try:
            two = int(fc.Operator) in (_FC_BETWEEN, _FC_NOT_BETWEEN)
        except Exception:  # noqa: BLE001
            two = True
        return [("Formula1", True), ("Formula2", two)]
    return [("Formula1", True), ("Formula2", True)]


def _reach_of_open_workbook(wb, netreach) -> dict | None:
    """network_reach's answer, read from a workbook Excel has open: cell
    formulas (Excel's own Find over formulas, each hit confirmed by its
    formula text), defined names, conditional-format rules and data
    connections. A sheet that cannot be searched counts as one that might
    reach the network, and so does any required read that fails (a defined
    name's target, the connection count or a name, the formula of a rule that
    has one): it is listed under "unread", never taken as "nothing found"."""
    formulas: list[dict] = []
    unread: list[str] = []
    for ws in wb.Worksheets:
        name = str(ws.Name)
        try:
            used = ws.UsedRange
            for fn in netreach.NETWORK_FUNCTIONS:
                first = used.Find(What=fn + "(", LookIn=_XL_FORMULAS, LookAt=_XL_PART,
                                  MatchCase=False)
                if first is None:
                    continue
                start = str(first.Address)
                cell, n = first, 0
                while True:
                    if n >= _FIND_LIMIT:
                        # more cells hold the name than this look confirms one by one:
                        # it cannot rule a call out, so it counts as one
                        formulas.append({"where": f"sheet {name} (more than {_FIND_LIMIT} cells hold "
                                         f"the text {fn}(, more than Excel's look confirms)",
                                         "functions": [fn]})
                        break
                    try:
                        text = str(cell.Formula2)
                    except Exception:  # noqa: BLE001
                        text = str(cell.Formula)
                    if fn in netreach.calls_in(text):
                        formulas.append({"where": f"{name}!{str(cell.Address).replace('$', '')}",
                                         "functions": [fn]})
                        break
                    cell = used.FindNext(cell)
                    n += 1
                    if cell is None or str(cell.Address) == start:
                        break
            rules: list[str] = []
            fcs = ws.Cells.FormatConditions
            for i in range(1, int(fcs.Count) + 1):
                fc = fcs.Item(i)
                for attr, required in _rule_formula_reads(fc):
                    try:
                        text = str(getattr(fc, attr) or "")
                    except Exception:  # noqa: BLE001
                        if required:
                            unread.append(f"{attr} of conditional format rule {i} on {name}")
                        continue
                    for f in netreach.calls_in(text):
                        if f not in rules:
                            rules.append(f)
            if rules:
                formulas.append({"where": f"a conditional format rule on {name}",
                                 "functions": rules})
        except Exception:  # noqa: BLE001
            formulas.append({"where": f"sheet {name} (Excel could not search it)",
                             "functions": ["unknown"]})
    try:
        names = list(wb.Names)
    except Exception:  # noqa: BLE001
        names = []
        unread.append("the defined names (Excel could not list them)")
    for n in names:
        try:
            label = str(n.Name)
        except Exception:  # noqa: BLE001
            label = "(unnamed)"
        try:
            funcs = netreach.calls_in(str(n.RefersTo or ""))
        except Exception:  # noqa: BLE001
            unread.append(f"the defined name {label} (Excel could not read what it refers to)")
            continue
        if funcs:
            formulas.append({"where": f"the defined name {label}", "functions": funcs})
    connections: list[str] = []
    try:
        count = int(wb.Connections.Count)
    except Exception:  # noqa: BLE001
        count = 0
        unread.append("the data connections (Excel could not count them)")
    for i in range(1, count + 1):
        try:
            connections.append(str(wb.Connections.Item(i).Name) or "(unnamed)")
        except Exception:  # noqa: BLE001
            # the connection exists; only its name is missing, and it is still one
            connections.append(f"(connection {i}, name unreadable)")
    if not formulas and not connections and not unread:
        return None
    out: dict = {"formulas": formulas, "connections": connections}
    if unread:
        out["unread"] = unread
    return out


#: What the network check found at this thread's shared opens. The COM worker
#: runs one job at a time on one thread, and run_com reads and clears this
#: around each job, so every op that opens a workbook through open_workbook
#: reports an allowed network the same way (_with_network_note).
_NETWORK = threading.local()

#: The sentence every COM op's success result carries, with network_allowed,
#: when allow_network let Excel open a workbook that can reach the network.
NETWORK_NOTE = ("Excel was allowed to contact the internet for this workbook "
                "because allow_network was set.")


def network_gate(app, abs_path: str, pw: str, allow_network: bool,
                 op: str | None) -> bool:
    """THE network check every COM open passes (open_workbook). Excel
    recalculates a workbook as it opens when the workbook asks for it (an
    edit by this server or by openpyxl asks; so does a legacy file), and
    some functions and data connections fetch from the network as Excel
    calculates. Unless allow_network, a workbook that can reach the network
    is refused here, before Excel opens it: read from its package or its
    text (core.netreach.file_reach), or, for a file this server cannot read,
    by Excel itself with calculation turned off (_reach_in_excel).

    With allow_network the same check runs to find out whether the call
    needed the permission: the answer is True when the workbook can reach
    the network (or the check could not rule it out), and the op's result
    then says so (_with_network_note). A check that fails with
    allow_network set never stops the open."""
    from ..core import netreach as _netreach
    if allow_network:
        try:
            how, reach = _netreach.file_reach(abs_path)
            if how == _netreach.READ_IN_EXCEL:
                reach = _reach_in_excel(app, abs_path, pw)
        except Exception:  # noqa: BLE001
            return False
        return bool(reach)
    how, reach = _netreach.file_reach(abs_path)
    looked = False
    if how == _netreach.READ_IN_EXCEL:
        reach = _reach_in_excel(app, abs_path, pw)
        looked = True
    if reach:
        raise _netreach.network_refusal(abs_path, reach, op or "open it in Excel",
                                        started=True, looked=looked)
    return False


def open_workbook(app, path: str, *, read_only: bool = False,
                  password: str | None = None, allow_network: bool = False,
                  network_op: str | None = None):
    """Workbooks.Open with the standing hygiene: no link updates, password
    passed POSITIONALLY, retry on busy. A repair-demand or refusal surfaces
    as Excel's own message.

    POSITIONAL ARGS ARE LOAD-BEARING (probe-proven on this machine, pywin32
    312 dynamic dispatch): Open with Password as a NAMED argument silently
    fails to deliver the password, the modal prompt appears on the hidden
    instance, and the call hangs until the timeout kills the worker. The
    same open with positional arguments delivers the password: a correct
    one opens in milliseconds and a wrong one raises Excel's own
    wrong-password error immediately. Signature:
    Open(FileName, UpdateLinks, ReadOnly, Format, Password,
    WriteResPassword, IgnoreReadOnlyRecommended). The sentinel makes an
    absent password fail fast instead of prompting; unencrypted files
    ignore it.

    A read-write open never leaves Excel waiting on a dialog:

    - a workbook with a password to modify is refused HERE, before Excel
      opens it, whenever the file check can read the setting, whatever
      route reached it. The mutating routes
      refuse it earlier still, before Excel starts;
    - the open hands Excel a stand-in password to modify, so a password the
      file check cannot see (inside Excel's encryption) fails at once
      instead of asking (_diagnose_refused_write_open names the cause);
    - a workbook Excel could open only read-only (another program holds it)
      is closed unsaved and refused;
    - the workbook Excel hands back must BE the file: a
      template is opened for editing (Editable=True), and anything Excel
      still opens as a new workbook is closed unsaved and refused, so no
      edit is ever made to a copy that a save would write elsewhere.

    A read-only open does not ask about a password to modify and is
    unchanged."""
    pw = password if password not in (None, "") else _NO_PASSWORD_SENTINEL
    abs_path = os.path.abspath(path)
    from ..core import hazard as _hazard
    # An XLM 4.0 macro sheet can hold an auto-run macro, and the worker's
    # AutomationSecurity=ForceDisable does not disable XLM macros, so no COM
    # open (read-only or read-write) opens such a file; it is refused here,
    # at the one choke point every COM open passes, before Excel starts.
    _hazard.refuse_xlm_macro_sheet(abs_path)
    # No COM open lets Excel reach the network unasked (network_gate).
    if network_gate(app, abs_path, pw, allow_network, network_op):
        _NETWORK.allowed = True
    if read_only:
        return _open(app, abs_path, read_only=True, password=pw,
                     write_password=None)
    _hazard.refuse_if_write_reserved(abs_path)
    stand_in = _WRITE_STAND_IN_OK is not False
    if not stand_in and _hazard.ole_container_kind(abs_path) is not None:
        # This Excel refuses the stand-in, and the file check cannot read
        # inside an OLE container (Excel's encryption, a legacy file), so a
        # password to modify cannot be excluded: refuse, never risk asking.
        raise _hazard.write_reserved_refusal(abs_path, unreadable=True,
                                             excel_started=True)
    try:
        wb = _open(app, abs_path, read_only=False, password=pw,
                   write_password=_NO_WRITE_PASSWORD_SENTINEL if stand_in
                   else None)
    except XlMcpError:
        raise
    except Exception as exc:  # noqa: BLE001
        if not stand_in or excel_process_died(exc):
            raise
        wb = _diagnose_refused_write_open(app, abs_path, pw, exc)
    _refuse_read_only_open(wb, abs_path)
    _refuse_other_workbook(wb, abs_path)
    return wb


#: Earlier versions had Excel write the edited copy into a private folder
#: of its own in the system's temp area (ks4xl-save-<pid>-<random>). Excel
#: writes every relative hyperlink relative to the folder it saves INTO, so
#: once that copy was put back beside the user's files each such link named
#: a path that does not exist. The copy is now written beside the workbook
#: (copy_path_for); these names remain so that the folders earlier
#: versions left in the temp area are still swept.
SAVE_FOLDER_PREFIX = "ks4xl-save-"

#: A save folder this old is a leftover whoever made it: no save keeps its
#: folder anywhere near this long.
STALE_SAVE_FOLDER_SECONDS = 60 * 60

#: ks4xl-save-<pid>-<random>, or the older ks4xl-save-<random>, where
#: <random> is tempfile.mkdtemp's eight letters, digits and underscores.
_SAVE_FOLDER_NAME = re.compile(r"ks4xl-save-(?:(\d{1,10})-)?[a-z0-9_]{8}")

#: What a save leaves in its folder: Excel's copy (book.<ext>), its owner
#: file (~$book.<ext>), and the scratch file Excel writes while it saves
#: (eight hex digits, sometimes .tmp).
_SAVE_FOLDER_FILE = re.compile(
    r"(?:~\$)?book\.[A-Za-z0-9]{1,8}|[0-9A-Fa-f]{8}(?:\.tmp)?")


def _save_leftover(entry: os.DirEntry, now: float) -> list[str] | None:
    """The files to remove when `entry` is a save folder a save left behind,
    else None (sweep_stale_save_folders says when)."""
    m = _SAVE_FOLDER_NAME.fullmatch(entry.name)
    if m is None:
        return None
    try:
        if not entry.is_dir(follow_symlinks=False) or entry.is_junction():
            return None
        files, newest = [], entry.stat(follow_symlinks=False).st_mtime
        with os.scandir(entry.path) as inner:
            for f in inner:
                if (not _SAVE_FOLDER_FILE.fullmatch(f.name)
                        or not f.is_file(follow_symlinks=False)
                        or f.is_junction()):
                    return None
                newest = max(newest, f.stat(follow_symlinks=False).st_mtime)
                files.append(f.path)
    except OSError:
        return None
    if now - newest > STALE_SAVE_FOLDER_SECONDS:
        return files
    pid = int(m.group(1)) if m.group(1) else None
    if pid is None or pid == os.getpid():
        return None
    from ..core import safesave as _safesave
    return None if _safesave._pid_alive(pid) else files


def sweep_stale_save_folders(root: str | None = None) -> list[str]:
    """Remove the private save folders that earlier saves left in the temp
    area: a server killed in the middle of a save left its folder, with
    Excel's copy, there for good. Run before each COM save.

    A folder goes only when ALL of these hold: its name has this server's
    prefix; it is a real folder (not a link or a junction) holding nothing
    but files a save makes there (Excel's copy, its owner file, the scratch
    file Excel writes while saving); and the server that made it is gone,
    or the folder is more than STALE_SAVE_FOLDER_SECONDS old. A folder of a
    live server (this one included) is judged by age alone, and so is a
    folder whose name carries no process ID. Files are removed one by
    one and then the empty folder, never a tree, so nothing outside the
    folder can be reached. Never raises; returns the folders removed."""
    import tempfile

    removed: list[str] = []
    try:
        now = time.time()
        with os.scandir(root or tempfile.gettempdir()) as it:
            entries = [e for e in it if e.name.startswith(SAVE_FOLDER_PREFIX)]
        for entry in entries:
            files = _save_leftover(entry, now)
            if files is None:
                continue
            try:
                for f in files:
                    os.remove(f)
                os.rmdir(entry.path)
                removed.append(entry.path)
            except OSError:
                continue  # another sweep got there first, or it is in use
    except Exception:  # noqa: BLE001 (a sweep never fails a save)
        pass
    return removed


#: Excel saves to a path of at most this many characters.
EXCEL_PATH_LIMIT = 218

#: The name Excel's copy gets beside the workbook: the save path's scratch
#: prefix (core.safesave.WRITE_TMP_PREFIX), six hex digits, and the
#: workbook's own extension. discard_saved removes nothing else.
_COPY_NAME = re.compile(r"\.ks4xl-write-[0-9a-f]{6}\.[A-Za-z0-9]{1,8}")


def copy_path_for(path: str) -> str:
    """Where Excel writes the edited copy of the workbook at `path`: beside
    it, in its own folder, under a short name the save path's stale-temp
    sweep knows.

    Excel writes every relative hyperlink and external link relative to the
    folder it saves into. MEASURED (Excel 16): a copy
    written into the workbook's own folder keeps every hyperlink and every
    external link exactly as Excel's own Save in place writes them; one in a
    subfolder shifts them a level; one in the temp area breaks them all.
    Excel killed in the middle of the copy leaves only that file (0 bytes),
    nothing else.

    Refused before anything changes (LIMIT_EXCEEDED) when the copy's full
    path would be longer than the 218 characters Excel saves to: Excel's
    copy could not be written there, and writing it anywhere else would
    break the workbook's links."""
    import uuid

    from ..core import safesave as _safesave
    target = os.path.abspath(path)
    folder = os.path.dirname(target) or "."
    ext = os.path.splitext(target)[1]
    copy = ""
    for _ in range(32):
        copy = os.path.join(
            folder, f"{_safesave.WRITE_TMP_PREFIX}{uuid.uuid4().hex[:6]}{ext}")
        if not os.path.lexists(copy):
            break
    if len(copy) > EXCEL_PATH_LIMIT:
        exc = XlMcpError(
            f"{os.path.basename(target)} was not changed. Excel saves a workbook "
            "by writing a copy beside it first, and that copy's full path "
            f"would be {len(copy)} characters, over the {EXCEL_PATH_LIMIT} "
            "Excel can save to. Move the workbook to a folder with a shorter "
            "path and retry; the xl_edit tools, which do not use Excel, have "
            "no such limit.")
        exc.code = "LIMIT_EXCEEDED"
        raise exc
    return copy


def save_workbook(wb, *, label: str, path: str, copy: str | None = None) -> str:
    """Write the edited workbook to a NEW file and return that file's path,
    for every route that saves a workbook back to its own path. Excel never
    replaces the workbook's own file. The caller closes the
    workbook in Excel and then puts the copy in place with put_in_place,
    or drops it with discard_saved.

    The copy goes beside the workbook (copy_path_for, or `copy` when the
    caller chose it already), so Excel writes the workbook's relative links
    exactly as they are.

    Excel's own save replaced the file itself: it writes the new version
    beside it and swaps it in. When another program held the file without
    sharing delete (a script or viewer reading it with the C runtime's
    default sharing), the swap failed and Excel put a visible "Save As"
    dialog on the screen, which DisplayAlerts does not suppress, until the
    call timed out; the edited copy stayed beside the user's file under a
    name like D421B010. The check before Excel starts (guard_target_closed)
    catches a holder that is already there, not one that opens the file
    while Excel works. MEASURED: such a reader arriving after that
    check brought the dialog back.

    Workbook.SaveCopyAs writes the workbook as it is in Excel, in its own
    format, to another file, and leaves the open workbook and its file
    alone. MEASURED (Excel 16, private desktop): .xlsx, .xlsm, .xlsb,
    .xltx and .xltm copies keep their format and hold the edit; a password
    set on the workbook encrypts the copy, which then opens with the new
    password and not the old one, and an empty one decrypts it; the file
    stays byte for byte as it was; nothing new appears beside it but Excel's
    owner file while the workbook is open; the copy takes 0.01 to 0.14 s.

    - Never on a read-only workbook: Excel could open it only
      read-only because another program holds it, and the copy could not
      be put in place either.
    - Never unless the workbook IS `path`: a new workbook
      Excel made from a template is not the file, and putting it in place
      of the template would replace the template with something else."""
    try:
        read_only = bool(getattr(wb, "ReadOnly", False))
    except Exception:  # noqa: BLE001
        read_only = True
    if read_only:
        from ..core import access as _access
        raise WorkbookLocked(_access.read_only_open_text(
            getattr(wb, "FullName", None) or "the workbook"))
    if not is_target(wb, path):
        raise _other_workbook_refusal(wb, path)
    sweep_stale_save_folders()  # what earlier saves left in the temp area
    from ..core.package import _sweep_stale_write_tmps
    _sweep_stale_write_tmps(os.path.dirname(os.path.abspath(path)))
    if copy is None:
        copy = copy_path_for(path)
    try:
        try:
            com_retry(lambda: wb.SaveCopyAs(copy), label=label)
        except Exception:
            from ..core import access as _access
            folder = os.path.dirname(os.path.abspath(path)) or "."
            if not _access._folder_takes_new_files(folder):
                raise WorkbookLocked(_access.folder_refuses_text(path)) from None
            raise
        if not os.path.isfile(copy) or os.path.getsize(copy) == 0:
            raise XlMcpError(
                f"{label}: Excel reported the save done but wrote nothing. "
                f"{os.path.basename(path)} was not changed.")
    except BaseException:
        discard_saved(copy)
        raise
    return copy


def discard_saved(copy: str | None) -> None:
    """Remove a copy that save_workbook wrote beside the workbook. Only that
    one file, and only under the copy's own name pattern; never a folder.
    An Excel that was just ended may hold it for a moment, so the removal
    is tried a few times; what still cannot go is swept by a later save
    once it is an hour old."""
    if not copy or not _COPY_NAME.fullmatch(os.path.basename(copy)):
        return
    from ..core import access as _access
    target = _access.long_path(os.path.abspath(copy))
    for pause in (0.0, 0.05, 0.1, 0.2, 0.4, 0.8):
        if pause:
            time.sleep(pause)
        try:
            os.remove(target)
            return
        except FileNotFoundError:
            return
        except OSError:
            continue


def check_saved_copy(path: str, *, encrypted: bool = False,
                     integrity_baseline: set[str] | None = None
                     ) -> Callable[[str], None]:
    """The checks Excel's copy must pass BEFORE it replaces the workbook at
    `path`, for put_in_place's `check`.

    They used to run on the workbook after the copy had replaced it, with
    the backup put back on a failure; with backup=false there was nothing
    to put back and a bad save stayed in place. On the copy, a bad save
    never replaces the workbook, backup or not:

    - an encrypted save: the copy must be Excel's encryption, an OLE
      compound file and not a plain package (a plain one means the
      password did not take);
    - anything else: the structural verify (core.verify.structural_check),
      and, given the broken links the workbook already had, the integrity
      check, which fails only on links the save broke."""
    from ..core import verify as _verify

    name = os.path.basename(path)

    def check(copy: str) -> None:
        if encrypted:
            if not _is_encrypted_container(copy):
                raise ValidationFailed(
                    f"Excel's copy is not encrypted, so the password did not "
                    f"take; {name} was not changed.")
            return
        ok, reasons = _verify.structural_check(copy)
        if not ok:
            raise ValidationFailed(
                f"Excel's copy failed the structural verify, so {name} was "
                "not changed: " + "; ".join(reasons))
        if integrity_baseline is not None:
            ok, broken = _verify.integrity_check(copy, integrity_baseline)
            if not ok:
                raise ValidationFailed(
                    "Excel's copy has broken internal links Excel would "
                    f"repair or refuse, so {name} was not changed: "
                    + "; ".join(broken[:8]))

    return check


def _is_encrypted_container(path: str) -> bool:
    """Excel's encryption writes an OLE compound file holding the encrypted
    package; a zip is a plain package."""
    import zipfile
    try:
        with open(path, "rb") as fh:
            head = fh.read(len(_CFB_MAGIC))
    except OSError:
        return False
    return head == _CFB_MAGIC and not zipfile.is_zipfile(path)


def file_sha256(path: str) -> str | None:
    """SHA-256 of a file's bytes, or None when it cannot be read."""
    import hashlib

    from ..core import access as _access
    h = hashlib.sha256()
    try:
        with open(_access.long_path(os.path.abspath(path)), "rb") as fh:
            for block in iter(lambda: fh.read(1 << 20), b""):
                h.update(block)
    except OSError:
        return None
    return h.hexdigest()


def put_in_place(copy: str, path: str, *,
                 stamp: tuple[int, int] | None = None,
                 check: Callable[[str], None] | None = None,
                 warnings: list[str] | None = None,
                 ticket: Any = None) -> str | None:
    """Put the copy save_workbook wrote in place of the workbook at `path`,
    after Excel has closed the workbook. The caller holds the
    workbook's write lock (ops/comtier.py) in the thread that calls
    this, never the COM worker's.

    Excel wrote the copy beside the workbook (copy_path_for), under
    a .ks4xl-write- name the save path sweeps if a process dies here; it
    replaces the workbook through core.safesave.promote, the file tier's
    own route: ReplaceFileW keeps the file's permissions, Mark of the Web,
    attributes and creation time, a passing reader is retried, and the
    original stays exactly as it was if the replace cannot be done. A
    program that still holds the workbook gets a plain refusal naming it,
    never a dialog, and nothing of this call is left beside the file: the
    copy is removed whenever it was not put in place.

    stamp is the workbook's core.safesave.file_stamp from when the
    save began. Just before the replace the file must still be that file:
    a program that does not use the write lock may have written to it while
    Excel worked, and its change wins. The copy is then dropped and the
    call refused in the file tier's own words (CONFLICT), the workbook
    untouched.

    check (check_saved_copy) runs first, on the copy where Excel wrote
    it: a copy that fails never replaces the workbook, and the workbook and
    its folder are as they were. Returns the SHA-256 of the bytes that went
    in place, so the caller can confirm the file is the checked copy.
    warnings receives what the promote reports, such as a workbook
    whose own permissions a folder would not let it keep. ticket is
    the backup rotation staged for this save; the promote lands it after
    the swap, and it is dropped on any failure here."""
    import shutil
    import uuid

    from ..core import access as _access
    from ..core import deadline as _deadline
    from ..core import safesave as _safesave

    target = os.path.abspath(path)
    folder = os.path.dirname(target) or "."
    if os.path.normcase(os.path.dirname(os.path.abspath(copy))) == os.path.normcase(folder):
        beside = copy  # Excel wrote it here
    else:
        beside = os.path.join(
            folder,
            f"{_safesave.WRITE_TMP_PREFIX}{uuid.uuid4().hex}{os.path.splitext(target)[1]}")
    try:
        if check is not None:
            check(copy)
        digest = file_sha256(copy)
        _deadline.commit_or_raise("the save")
        if beside != copy:
            try:
                shutil.copyfile(_access.long_path(copy), _access.long_path(beside))
            except PermissionError:
                raise WorkbookLocked(_access.write_refusal_text(target)) from None
        if stamp is not None and _safesave.file_stamp(target) != stamp:
            raise _safesave.changed_on_disk_error(target)
        try:
            promoted = _safesave.promote(beside, target, ticket=ticket)
        except PermissionError as exc:
            raise WorkbookLocked(_access.replace_refusal_text(
                target, getattr(exc, "winerror", None))) from None
        if warnings is not None:
            warnings.extend(promoted)
        return digest
    except OSError as exc:
        from ..core.package import _os_write_error
        raise _os_write_error(os.path.basename(target), exc) from exc
    finally:
        if beside != copy:
            try:
                os.remove(_access.long_path(beside))
            except OSError:
                pass  # promoted, so it is the workbook now; or never written
        discard_saved(copy)  # gone already when it was promoted


def _assert_hygiene(app) -> tuple[Any, Any]:
    """Re-assert DisplayAlerts=False / EnableEvents=False before an op;
    return the prior values so the wrapper can restore them after."""
    prior_alerts = None
    prior_events = None
    try:
        prior_alerts = app.DisplayAlerts
        app.DisplayAlerts = False
    except Exception:
        pass
    try:
        prior_events = app.EnableEvents
        app.EnableEvents = False
    except Exception:
        pass
    return prior_alerts, prior_events


def _restore_hygiene(app, prior_alerts, prior_events) -> None:
    try:
        if prior_alerts is not None:
            app.DisplayAlerts = prior_alerts
    except Exception:
        pass
    try:
        if prior_events is not None:
            app.EnableEvents = prior_events
    except Exception:
        pass


# ---------------------------------------------------------------- executor


@dataclass
class _Job:
    label: str
    fn: Callable[[ExcelInstanceManager], Any]
    done: threading.Event = field(default_factory=threading.Event)
    result: Any = None
    error: BaseException | None = None
    generation: int = 0


class ComExecutor:
    """One dedicated COM worker thread owning the pooled Excel instance.

    Serialization is structural: the single thread runs one job at a time in
    one COM apartment. submit() enforces the deadline; a timed-out job
    poisons the current worker generation (owned PIDs are taskkilled, which
    unblocks the stuck COM call) and the next submit re-arms fresh."""

    def __init__(self, journal_path: str | os.PathLike | None = None):
        self._journal_path = journal_path
        self._queue: "queue.Queue[_Job]" = queue.Queue()
        self._thread: threading.Thread | None = None
        self._manager: ExcelInstanceManager | None = None
        self._generation = 0
        self._state = threading.Lock()
        self._swept_startup = False
        # honest status counters
        self._busy_with: str | None = None
        self._busy_since: float | None = None
        self._last_op: str | None = None
        self._last_op_seconds: float | None = None
        self._last_op_ok: bool | None = None
        self._ops_completed = 0
        self._timeouts = 0
        self._contention_waits = 0
        #: Background restart-transient reaps started by _poison. They are
        #: off the caller's deadline but NOT unaccounted for: shutdown joins
        #: them and status() reports whether one is running.
        self._reap_threads: list[threading.Thread] = []

    # --- worker lifecycle -------------------------------------------------

    def _ensure_worker(self) -> None:
        with self._state:
            if self._thread is not None and self._thread.is_alive():
                return
            self._generation += 1
            gen = self._generation
            self._manager = ExcelInstanceManager(self._journal_path)
            if not self._swept_startup:
                self._swept_startup = True
                try:
                    self._sweep_stale_journal(self._manager)
                except Exception:
                    pass
            t = threading.Thread(
                target=self._worker_loop, args=(gen, self._manager),
                name=f"ks4xl-com-worker-{gen}", daemon=True)
            self._thread = t
            t.start()

    @staticmethod
    def _sweep_stale_journal(manager: ExcelInstanceManager) -> dict:
        """JOURNAL REPLAY at startup: reclaim owned Excel PIDs that no live
        server is responsible for. BY OWNED PID ONLY, AND ONLY WHEN THE PID
        STILL NAMES THE PROCESS WE JOURNALED.

        Why the identity check: this sweep used to ask ``pid_alive(pid)``, which only answers "is
        SOME EXCEL.EXE holding that PID right now". After a crash leaves a
        record behind and Windows recycles the PID onto the user's Excel,
        that question says yes and the next server start force-killed their
        unsaved work. ``reclaim_journaled_pid`` compares the recorded
        creation time first and drops a recycled record without killing
        anything.

        Two populations:

          - ABANDONED: the record names an owning SERVER process that is gone
            (crash, or a clean exit whose Quit never landed). Nobody is
            coming back for these, so they are reclaimed at once. Waiting out
            the 15-minute age window instead is how an invisible EXCEL.EXE
            survived a clean shutdown, outlived the grace period, and could
            not be reaped by the next server.
          - AGED: older than STALE_JOURNAL_SECONDS with no owner recorded
            (a record written before owner tracking). Unchanged.

        A record whose owner is STILL ALIVE belongs to a concurrent session
        and is never touched."""
        from .instances import reclaim_journaled_pid

        data = manager.journal.records()
        abandoned = manager.journal.abandoned_pids()
        now = time.time()
        outcomes: dict[str, list[int]] = {}
        for pid_s, rec in list(data.items()):
            try:
                pid = int(pid_s)
            except ValueError:
                continue
            age = now - float(rec.get("spawned_at", now))
            owner = rec.get("owner_pid") if isinstance(rec, dict) else None
            has_owner = isinstance(owner, int) and not isinstance(owner, bool)
            # The age clause is for records with no owner recorded only: a
            # record whose owner is alive belongs to a concurrent session,
            # however long that session has kept its Excel.
            if pid not in abandoned and (has_owner or age < STALE_JOURNAL_SECONDS):
                continue
            outcome = reclaim_journaled_pid(manager.journal, pid, rec)
            outcomes.setdefault(outcome, []).append(pid)
        return outcomes

    def _worker_loop(self, gen: int, manager: ExcelInstanceManager) -> None:
        # Every job here runs for a caller thread that may hold the
        # workbook's write lock (ops/comtier.py); a hold taken on this thread
        # would wait for that caller, so it fails at once instead.
        from ..core import safesave as _safesave
        _safesave.forbid_write_lock_on_this_thread(
            "the COM worker thread runs jobs for callers that hold the "
            "workbook's write lock; take it in the calling thread")
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            while True:
                try:
                    job = self._queue.get(timeout=1.0)
                except queue.Empty:
                    with self._state:
                        if gen != self._generation:
                            return  # superseded; exit quietly
                    continue
                with self._state:
                    if gen != self._generation:
                        # poisoned while idle; refuse the job back to a live
                        # generation by re-queueing it for the new worker.
                        self._queue.put(job)
                        return
                    self._busy_with = job.label
                    self._busy_since = time.monotonic()
                t0 = time.monotonic()
                ok = False
                try:
                    job.result = job.fn(manager)
                    ok = True
                except BaseException as exc:  # noqa: BLE001
                    job.error = exc
                finally:
                    # The job is over, however it ended: its Excel no
                    # longer speaks for the next job's failures.
                    from .instances import forget_current
                    forget_current()
                dt = time.monotonic() - t0
                with self._state:
                    self._busy_with = None
                    self._busy_since = None
                    self._last_op = job.label
                    self._last_op_seconds = round(dt, 3)
                    self._last_op_ok = ok
                    self._ops_completed += 1
                    stale = gen != self._generation
                if stale:
                    # a stale (timed-out) job's caller already got its
                    # refusal; setting done now would be a lie. The result
                    # is dropped and this dead-generation worker exits.
                    return
                job.done.set()
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- the public entry point ------------------------------------------

    def submit(self, label: str,
               fn: Callable[[ExcelInstanceManager], Any],
               timeout: float | None = None) -> Any:
        """Run fn(manager) on the COM worker thread, serialized with every
        other COM operation, bounded by timeout seconds."""
        require_com()
        deadline = DEFAULT_TIMEOUT if timeout is None else float(timeout)
        if deadline <= 0:
            deadline = DEFAULT_TIMEOUT
        self._ensure_worker()
        with self._state:
            if self._busy_with is not None or not self._queue.empty():
                self._contention_waits += 1
        # The job runs under a copy of the SUBMITTING
        # call's context, so the call's deadline (core/deadline.py) is active
        # on the COM worker thread and Excel's own saves commit against it:
        # once the call has answered TIMEOUT, Excel never saves. Before, the
        # long-lived worker thread saw no deadline at all.
        ctx = contextvars.copy_context()
        job = _Job(label=label, fn=lambda manager: ctx.run(fn, manager),
                   generation=self._generation)
        self._queue.put(job)
        if not job.done.wait(deadline):
            self._poison(job)
            raise ExcelBlocked(
                f"{label} exceeded the {deadline:.0f}s timeout; the worker "
                "Excel instance was reclaimed (by owned PID) and the COM "
                "layer re-armed. The file was not promoted by this call; "
                "retry, raise timeout_seconds, or split the operation.")
        if job.error is not None:
            raise job.error
        return job.result

    def _poison(self, job: _Job) -> None:
        """Timeout remediation: retire the worker generation and reclaim its
        owned Excel PIDs (taskkill unblocks a stuck synchronous COM call).

        OBSERVED (COM-tier gate): killing an Excel that is showing a modal
        dialog can trigger a transient Office restart-recovery EXCEL.EXE a
        moment later. When NO foreign Excel existed at kill time, any PID
        appearing right after the kill is that artifact of our own kill; it
        is journaled and reaped. With any foreign Excel present the reap is
        skipped entirely (never risk a user process).

        THE REAP RUNS IN THE BACKGROUND. It watches for six seconds and settles for two,
        and it fires exactly when no foreign Excel exists, which is the
        unattended case: every time. Running it inline before submit() raised
        put those 8.9 seconds on the caller's clock, so a tool asked for a
        3-second deadline returned at 11.9 and the documented ceiling of a
        default 60-second call was really 69. The kill that unblocks the
        stuck COM call still happens inline (it is fast and it is what the
        refusal promises); only the watch-and-settle moves off the deadline.
        shutdown() joins the reap, so the zero-orphan guarantee is unchanged.
        """
        with self._state:
            self._timeouts += 1
            self._generation += 1
            manager = self._manager
            self._thread = None
            self._manager = None
        if manager is None:
            return
        try:
            from .instances import list_excel_pids

            # Another live server's Excel is foreign here too, so
            # the restart-transient reap never runs while one exists.
            foreign_before = list_excel_pids() - manager.reclaimable_pids()
            manager.force_reclaim()
        except Exception:
            return
        if not foreign_before:
            self._start_background_reap(manager)

    def _start_background_reap(self, manager: ExcelInstanceManager) -> None:
        def _run() -> None:
            try:
                self._reap_restart_transients(manager)
            except Exception:  # noqa: BLE001
                pass

        t = threading.Thread(target=_run, name="ks4xl-com-reap", daemon=True)
        with self._state:
            self._reap_threads = [x for x in self._reap_threads if x.is_alive()]
            self._reap_threads.append(t)
        t.start()

    def join_reaps(self, timeout: float = 15.0) -> bool:
        """Block until every background restart-transient reap has finished.
        Called by shutdown() (and therefore by atexit) so moving the reap off
        the caller's deadline cannot leave an orphan behind. Returns True when
        all of them finished inside the timeout."""
        deadline = time.monotonic() + float(timeout)
        with self._state:
            threads = list(self._reap_threads)
        for t in threads:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            t.join(remaining)
        done = all(not t.is_alive() for t in threads)
        if done:
            with self._state:
                self._reap_threads = [x for x in self._reap_threads
                                      if x.is_alive()]
        return done

    @staticmethod
    def _reap_restart_transients(manager: ExcelInstanceManager,
                                 watch_seconds: float = 6.0) -> None:
        from .instances import list_excel_pids, reclaim_journaled_pid

        deadline = time.monotonic() + watch_seconds
        seen: set[int] = set()
        while time.monotonic() < deadline:
            for pid in list_excel_pids():
                if pid not in seen:
                    seen.add(pid)
                    manager.journal.record(pid)
            time.sleep(0.5)
        # allow self-exit (the recovery instance usually quits on its own),
        # then reclaim whatever we journaled that still lives.
        time.sleep(2.0)
        for pid in sorted(seen):
            reclaim_journaled_pid(manager.journal, pid)

    # --- status / shutdown ------------------------------------------------

    def status(self) -> dict:
        ok, reason = com_available()
        with self._state:
            manager = self._manager
            busy = self._busy_with
            busy_for = (round(time.monotonic() - self._busy_since, 3)
                        if self._busy_since is not None else None)
            out = {
                "com_available": ok,
                "com_availability": reason,
                "worker_alive": bool(self._thread and self._thread.is_alive()),
                "busy": busy is not None,
                "current_op": busy,
                "current_op_seconds": busy_for,
                "queued_ops": self._queue.qsize(),
                "last_op": self._last_op,
                "last_op_seconds": self._last_op_seconds,
                "last_op_ok": self._last_op_ok,
                "ops_completed": self._ops_completed,
                "timeouts": self._timeouts,
                "contention_waits": self._contention_waits,
                "default_timeout_seconds": DEFAULT_TIMEOUT,
                # A timeout's restart-transient reap runs off the caller's
                # deadline; say so rather than let it look like nothing is
                # happening.
                "reaping_in_background": any(
                    t.is_alive() for t in self._reap_threads),
            }
        pooled = None
        owned: list[int] = []
        journaled: list[int] = []
        abandoned: list[int] = []
        try:
            # The JOURNAL is on disk and is the record of what this server has
            # spawned; reading it through the live manager only made
            # journaled_pids honest AFTER a worker existed, so a fresh server
            # reported "journaled_pids: []" next to a stranded EXCEL.EXE it
            # had every record of.
            jm = manager if manager is not None else \
                ExcelInstanceManager(self._journal_path)
            journaled = sorted(jm.journal.owned_pids())
            abandoned = sorted(jm.journal.abandoned_pids())
            if manager is not None:
                pooled = manager._pool.pid if manager._pool else None
                owned = sorted({w.pid for w in manager._workers})
        except Exception:
            pass
        out["pooled_instance_pid"] = pooled
        out["tracked_worker_pids"] = owned
        if manager is not None and getattr(manager, "handed_over", None):
            # Excels this server left running because a workbook was opened
            # in them from outside (they are the user's now).
            out["left_to_user"] = list(manager.handed_over)
        out["journaled_pids"] = journaled
        out["abandoned_journal_pids"] = abandoned
        if abandoned:
            out["note"] = (
                "journaled Excel PIDs whose owning server process is gone; "
                "the next COM call's startup sweep reclaims them, or call "
                "com_status after any com_ tool to trigger it")
        return out

    def shutdown(self, *, reclaim: bool = True,
                 graceful_timeout: float = 8.0) -> None:
        """Teardown that actually ends the worker process.

        TWO STEPS, and the second one is why this exists. The old shutdown
        called ``manager.quit(w)`` straight from the calling thread, which at
        interpreter exit is the MAIN thread -- a different COM apartment from
        the single worker thread that created the proxy. The cross-apartment
        call raised, the ``except Exception: pass`` swallowed it, Quit never
        reached Excel, and the invisible worker outlived the server
        indefinitely: one was seen alive through a clean shutdown and past
        the 210s grace, with a fresh server unable to reap it.

        So: ask the WORKER THREAD to Quit (in its own apartment, bounded), and
        then reclaim by owned PID, which needs no apartment at all and is the
        only step that can be trusted at interpreter exit. Foreign EXCEL.EXE
        processes are never touched, and neither is the pooled Excel of
        another kitchensink4xl server: the journal is one file per user,
        and reclaim takes only this server's records and those whose server
        is gone (reclaimable_pids).

        A background restart-transient reap (see _poison) is joined FIRST, so
        moving that work off the caller's timeout deadline cannot turn into an
        orphan at exit."""
        self.join_reaps()
        with self._state:
            manager = self._manager
            thread = self._thread
            worker_live = bool(thread and thread.is_alive())
        if manager is None:
            self._retire_generation()
            return
        if worker_live:
            done = threading.Event()

            def _quit_in_apartment(mgr: ExcelInstanceManager) -> None:
                for w in list(mgr._workers):
                    try:
                        mgr.quit(w)
                    except Exception:
                        pass
                done.set()

            job = _Job(label="shutdown", fn=_quit_in_apartment,
                       generation=self._generation)
            self._queue.put(job)
            job.done.wait(graceful_timeout)
        self._retire_generation()
        if not reclaim:
            return
        # Excel defers process exit by minutes after Quit (exp 7). Waiting
        # that out at shutdown is not an option, and leaving the process to
        # "probably exit" is exactly what stranded one per session, so the
        # owned PIDs are reclaimed now. force_reclaim forgets a PID only once
        # it is confirmed gone, so a kill that fails stays journaled for the
        # next startup sweep.
        try:
            manager.force_reclaim()
        except Exception:
            pass

    def _retire_generation(self) -> None:
        with self._state:
            self._generation += 1
            self._thread = None
            self._manager = None


_EXECUTOR: ComExecutor | None = None
_EXECUTOR_LOCK = threading.Lock()


def owned_excel_pids() -> list[int]:
    """The Excel PIDs this server's pool started and still tracks, for the
    exit log line. Reads only; never starts the executor or Excel."""
    ex = _EXECUTOR
    manager = getattr(ex, "_manager", None) if ex is not None else None
    if manager is None:
        return []
    try:
        return sorted(manager.owned_pids())
    except Exception:  # noqa: BLE001 (a log line never fails the exit)
        return []


def get_executor() -> ComExecutor:
    global _EXECUTOR
    with _EXECUTOR_LOCK:
        if _EXECUTOR is None:
            _EXECUTOR = ComExecutor()
            atexit.register(_EXECUTOR.shutdown)
        return _EXECUTOR


def run_com(label: str, fn: Callable[[ExcelInstanceManager], Any],
            timeout: float | None = None) -> Any:
    """Module-level convenience over the process-wide executor. Every route
    that starts the private hidden Excel passes here, so this is where
    KS4XL_FORBID=xl_app closes them all."""
    from .. import runtime as _runtime

    why = _runtime.excel_forbidden()
    if why:
        from ..core.errors import SurfaceRefusal
        from ..surface import copy as _copy

        raise SurfaceRefusal("NOT_AVAILABLE", f"{label.split('(')[0]} needs Excel, and {why}",
                             details={"decided_by": _copy.SETTINGS_OWNER})
    return get_executor().submit(label, _with_network_note(fn), timeout=timeout)


def _with_network_note(fn: Callable[[ExcelInstanceManager], Any]):
    """fn, run so that its result says when allow_network let Excel open a
    workbook that can reach the network: network_allowed: true and
    network_note. The note is made where the open is (open_workbook), so
    an op that opens a workbook through it cannot leave it out."""
    def run(manager):
        _NETWORK.allowed = False
        try:
            out = fn(manager)
            allowed = bool(getattr(_NETWORK, "allowed", False))
        finally:
            _NETWORK.allowed = False
        if allowed and isinstance(out, dict):
            out = dict(out)
            out["network_allowed"] = True
            out["network_note"] = NETWORK_NOTE
        return out
    return run


# ------------------------------------------------------------- opens-clean


def opens_clean(path: str, timeout: float | None = None,
                allow_network: bool = False) -> dict:
    """The authoritative corruption smoke test: open the file in a pooled
    hidden worker and report whether Excel accepts it without a repair
    prompt. Under DisplayAlerts=False a repair-demand THROWS (the bare-FILTER
    spike case), so a clean Open IS the verdict. Read-only; never saves."""
    require_com()
    abs_path = os.path.abspath(path)
    from ..core import swapwait as _swapwait
    _swapwait.require_workbook(abs_path, f"no such workbook: {abs_path}")
    from ..core import hazard as _hazard
    if _hazard.has_xlm_macro_sheet(abs_path):
        # An XLM macro sheet is not opened by the COM route (ForceDisable
        # cannot disable XLM), so the honest answer is that Excel was not
        # asked, not that it refused.
        return {
            "opens_clean": None,
            "xlm_macro_sheet": True,
            "note": "the COM route does not open a workbook with an Excel "
                    "4.0 (XLM) macro sheet, so whether Excel would accept it "
                    "was not checked; the file-based tools read it without "
                    "running anything",
        }
    if is_encrypted_package(abs_path):
        # A password-less open of an encrypted file HANGS on a modal prompt
        # (probe-proven; no supplied argument suppresses it), so the honest
        # answer comes from the signature, not from Excel.
        return {
            "opens_clean": False,
            "encrypted": True,
            "excel_says": "the file is password-protected (CFB-encrypted "
                          "package); Excel would prompt for a password",
            "note": "supply the password (com_validate_opens_clean "
                    "password=...) to verify an encrypted file",
        }

    def body(manager: ExcelInstanceManager) -> dict:
        w = manager.acquire()
        wb = None
        prior = _assert_hygiene(w.app)
        try:
            try:
                wb = open_workbook(w.app, abs_path, read_only=True,
                                   allow_network=allow_network,
                                   network_op="validate that it opens clean")
            except XlMcpError as exc:
                if "network_reach" in (getattr(exc, "detail", None) or {}):
                    raise  # the network check's refusal, not Excel's verdict
                return {
                    "opens_clean": False,
                    "instance_pid": w.pid,
                    "excel_says": excel_error_message(exc),
                    "note": ("Excel refused the file or demanded a repair; "
                             "under suppressed alerts a repair prompt "
                             "surfaces as this refusal"),
                }
            except Exception as exc:  # noqa: BLE001
                return {
                    "opens_clean": False,
                    "instance_pid": w.pid,
                    "excel_says": excel_error_message(exc),
                    "note": ("Excel refused the file or demanded a repair; "
                             "under suppressed alerts a repair prompt "
                             "surfaces as this refusal"),
                }
            sheets = int(wb.Worksheets.Count)
            name = str(wb.Name)
            return {"opens_clean": True, "instance_pid": w.pid,
                    "workbook": name, "worksheets": sheets}
        finally:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            _restore_hygiene(w.app, *prior)
            manager.release_to_pool(w)

    return run_com(f"opens_clean({os.path.basename(abs_path)})", body,
                   timeout=timeout)


__all__ = [
    "ComExecutor", "get_executor", "run_com", "com_available", "require_com",
    "guard_target_closed", "com_retry", "open_workbook", "opens_clean",
    "excel_error_message", "excel_process_died", "DEFAULT_TIMEOUT",
    "STALE_JOURNAL_SECONDS", "sweep_stale_save_folders",
]
