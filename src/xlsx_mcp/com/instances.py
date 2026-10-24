"""com/instances.py: the multi-instance COM manager.

Grounded in measurements on a real Excel install.
Excel is a TRUE multi-instance COM server under DispatchEx, so the danger is
not singleton contention (that was PowerPoint) but PID-precise accounting:
Quit can defer process exit by minutes, a crashed client leaves a real zombie
that Quit cannot reap, and GetActiveObject returns the OLDEST registrant which
may be a worker, not the user.

This manager therefore:
  - spawns isolated workers with DispatchEx and identifies each new PID by
    diffing the EXCEL.EXE process set across the spawn (robust for invisible,
    workbook-less instances whose Hwnd is 0),
  - JOURNALS every spawned PID to disk at spawn time, TOGETHER WITH THAT
    PROCESS'S CREATION TIME, so a crash still leaves an owned-PID record for
    the next startup sweep and the sweep can prove the PID still names the
    same process,
  - EXCLUDES self-spawned PIDs from GetActiveObject when locating the user's
    instance,
  - pools one worker for reuse (a pooled instance saves ~2s per op),
  - sweeps zombies in two phases: a minutes-scale grace window, then taskkill
    BY OWNED PID ONLY, and only after the PID's creation time still matches
    what was journaled. It never touches a foreign EXCEL.EXE, including one
    that Windows has handed our old PID to.

IMPORTANT: pywin32 is imported lazily inside the methods that need it, so the
module imports cleanly on any platform and in headless CI. Nothing here spawns
Excel at import time.
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path

EXCEL_IMAGE = "EXCEL.EXE"
# Minutes-scale grace: Quit was measured deferring exit by 1-3 min.
DEFAULT_GRACE_SECONDS = 210.0

#: Warnings only; silent unless the host configures logging. A COM server on
#: stdio must never print to stdout, so nothing here does.
_log = logging.getLogger(__name__)


# ------------------------------------------------------------- PID plumbing

def list_excel_pids() -> set[int]:
    """Every live EXCEL.EXE PID on the machine. psutil when present, else
    `tasklist` (a substitute with no loss of fidelity)."""
    try:
        import psutil  # type: ignore

        out: set[int] = set()
        for p in psutil.process_iter(["name", "pid"]):
            nm = (p.info.get("name") or "")
            if nm.upper() == EXCEL_IMAGE:
                out.add(int(p.info["pid"]))
        return out
    except Exception:
        return _tasklist_pids()


#: Every helper process gets its own stdin, stdout and stderr, never the
#: server's: the MCP pipe is stdin, and a child that inherits a full stderr
#: pipe can hang at startup on Windows.
#: capture_output gives it its own stdout and stderr; no window either.
_CHILD = {"stdin": subprocess.DEVNULL,
          "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}


def _tasklist_pids() -> set[int]:
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {EXCEL_IMAGE}", "/FO", "CSV",
             "/NH"],
            capture_output=True, text=True, timeout=30, **_CHILD,
        )
    except Exception:
        return set()
    pids: set[int] = set()
    for line in proc.stdout.splitlines():
        parts = [c.strip('"') for c in line.split('","')]
        if len(parts) >= 2 and parts[0].strip('"').upper() == EXCEL_IMAGE:
            try:
                pids.add(int(parts[1]))
            except ValueError:
                pass
    return pids


def pid_alive(pid: int) -> bool:
    """True when this PID is a LIVE EXCEL.EXE. Deliberately image-scoped: a
    recycled PID belonging to some other program must never look like one of
    our workers."""
    return pid in list_excel_pids()


def _process_alive(pid: int) -> bool:
    """True when ANY process holds this PID. Used only to ask whether the
    SERVER that journaled a worker is still running; never to decide whether
    to kill an Excel."""
    if pid <= 0:
        return False
    try:
        import psutil  # type: ignore

        return psutil.pid_exists(pid)
    except Exception:
        pass
    try:
        r = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV",
                            "/NH"], capture_output=True, text=True, timeout=30,
                           **_CHILD)
        return str(pid) in r.stdout
    except Exception:
        return True  # cannot tell: assume alive, never reap on a guess


def process_create_time(pid: int) -> float | None:
    """The creation time of the process holding this PID, or None when it
    cannot be read (no psutil, the process is gone, access denied).

    This is the IDENTITY of a PID. Windows recycles PIDs, so "PID 30092 is a
    live EXCEL.EXE" does not mean "PID 30092 is the EXCEL.EXE we spawned";
    creation time is what distinguishes them."""
    if pid <= 0:
        return None
    try:
        import psutil  # type: ignore

        return float(psutil.Process(int(pid)).create_time())
    except Exception:
        return None


def _own_start_time() -> float | None:
    return process_create_time(os.getpid())


#: psutil reports creation time as a POSIX float and the JSON round trip is
#: exact, so this tolerance only absorbs clock/platform jitter. It is far
#: below the interval in which a PID could plausibly be recycled.
CREATE_TIME_TOLERANCE = 0.05

#: How far a journal record's owner_started may be from the owning server's
#: creation time and still name that server. Real records carry
#: the exact time; this only absorbs a record written by hand or late. A
#: recycled PID belongs to a process created long after its old owner.
OWNER_TIME_TOLERANCE = 2.0


def journal_create_time(rec: object) -> float | None:
    """The Excel creation time stored in a journal record, or None for a
    LEGACY record written before identity tracking existed."""
    if not isinstance(rec, dict):
        return None
    ct = rec.get("create_time")
    if isinstance(ct, bool) or not isinstance(ct, (int, float)):
        return None
    return float(ct)


def identity_verdict(pid: int, rec: object) -> str:
    """Does this PID still name the Excel the journal recorded?

      ``"match"``     the creation times agree: this is our process.
      ``"mismatch"``  the creation times differ: Windows recycled the PID onto
                      a DIFFERENT Excel, which we have no claim to.
      ``"unknown"``   no recorded creation time (a legacy record) or none
                      readable now (no psutil, access denied). The caller
                      falls back to the pre-identity behaviour.
    """
    expected = journal_create_time(rec)
    if expected is None:
        return "unknown"
    actual = process_create_time(pid)
    if actual is None:
        return "unknown"
    return ("match" if abs(actual - expected) <= CREATE_TIME_TOLERANCE
            else "mismatch")


#: The class of Excel's main window.
EXCEL_MAIN_WINDOW_CLASS = "XLMAIN"


def excel_window_shown(pid: int) -> bool:
    """True when the Excel process ``pid`` shows its main window on any
    desktop of this window station: someone can see it and work in it.

    The pooled Excel is hidden for its whole life. Excel shows it when a
    workbook is opened in it from outside: a DDE open (a file association,
    another program) lands in an instance that answers DDE and makes it
    visible and user-controlled (MEASURED). Such an instance holds the
    user's work and is never closed, quit or killed. This reads window
    attributes only (EnumDesktopWindows, GetClassName, IsWindowVisible),
    which never waits on the process, so it is safe while Excel is stuck.
    False off Windows, and when the windows cannot be read."""
    if os.name != "nt" or pid <= 0:
        return False
    try:
        import ctypes
        from ctypes import wintypes as w

        u32 = ctypes.WinDLL("user32", use_last_error=True)
        enum_win = ctypes.WINFUNCTYPE(w.BOOL, w.HWND, w.LPARAM)
        enum_desk = ctypes.WINFUNCTYPE(w.BOOL, w.LPWSTR, w.LPARAM)
        u32.GetWindowThreadProcessId.argtypes = [w.HWND, ctypes.POINTER(w.DWORD)]
        u32.IsWindowVisible.argtypes = [w.HWND]
        u32.GetClassNameW.argtypes = [w.HWND, w.LPWSTR, ctypes.c_int]
        u32.GetProcessWindowStation.restype = w.HANDLE
        u32.EnumDesktopsW.argtypes = [w.HANDLE, enum_desk, w.LPARAM]
        u32.OpenDesktopW.restype = w.HANDLE
        u32.OpenDesktopW.argtypes = [w.LPCWSTR, w.DWORD, w.BOOL, w.DWORD]
        u32.EnumDesktopWindows.argtypes = [w.HANDLE, enum_win, w.LPARAM]
        u32.CloseDesktop.argtypes = [w.HANDLE]
        found: list[int] = []
        cls = ctypes.create_unicode_buffer(64)

        def on_window(hwnd, _lp):
            owner = w.DWORD()
            u32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
            if owner.value == pid and u32.IsWindowVisible(hwnd):
                u32.GetClassNameW(hwnd, cls, 64)
                if cls.value == EXCEL_MAIN_WINDOW_CLASS:
                    found.append(1)
                    return False
            return True

        names: list[str] = []

        def on_desktop(name, _lp):
            names.append(name)
            return True

        window_cb, desktop_cb = enum_win(on_window), enum_desk(on_desktop)
        u32.EnumDesktopsW(u32.GetProcessWindowStation(), desktop_cb, 0)
        for name in names:
            h = u32.OpenDesktopW(name, 0, False, 0x0001 | 0x0040)  # READOBJECTS | ENUMERATE
            if not h:
                continue
            try:
                u32.EnumDesktopWindows(h, window_cb, 0)
            finally:
                u32.CloseDesktop(h)
            if found:
                return True
        if not names:
            u32.EnumWindows.argtypes = [enum_win, w.LPARAM]
            u32.EnumWindows(window_cb, 0)
        return bool(found)
    except Exception:  # noqa: BLE001
        return False


def _return_to_user(pid: int, remote_prior: object = None,
                    timeout: float = 5.0) -> bool:
    """Best effort, bounded: reach an Excel that is being left to the user
    through one of its workbook windows, and put back what the pool turned
    off (alerts, so closing a changed workbook asks to save; events; screen
    updating; the DDE option), and keep it running when nothing else holds
    it. Used where the server has no COM handle of its own on that Excel (a
    COM timeout, another server's leftover). False when it was not reached
    in time; nothing is ever closed here."""
    import threading

    done: list[bool] = []

    def attach() -> None:
        try:
            import pythoncom
            import win32com.client
            import win32gui
            import win32process
        except Exception:  # noqa: BLE001
            return
        pythoncom.CoInitialize()
        try:
            tops: list[int] = []
            kids: list[int] = []

            def on_top(h, _):
                if win32process.GetWindowThreadProcessId(h)[1] == pid and \
                        win32gui.GetClassName(h) == EXCEL_MAIN_WINDOW_CLASS:
                    tops.append(h)
                return True

            def on_child(h, _):
                if win32gui.GetClassName(h) == "EXCEL7":
                    kids.append(h)
                return True

            win32gui.EnumWindows(on_top, None)
            for top in tops:
                win32gui.EnumChildWindows(top, on_child, None)
            for hwnd in kids:
                try:
                    _r, lres = win32gui.SendMessageTimeout(
                        hwnd, 0x003D, 0, 0xFFFFFFF0, 0x0002, 2000)  # WM_GETOBJECT, OBJID_NATIVEOM
                    if not lres:
                        continue
                    obj = pythoncom.ObjectFromLresult(lres, pythoncom.IID_IDispatch, 0)
                    app = win32com.client.Dispatch(obj).Application
                except Exception:  # noqa: BLE001
                    continue
                for attr, val in _user_settings(remote_prior):
                    try:
                        setattr(app, attr, val)
                    except Exception:  # noqa: BLE001
                        pass
                done.append(True)
                return
        except Exception:  # noqa: BLE001
            pass
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass

    t = threading.Thread(target=attach, name="ks4xl-return-excel", daemon=True)
    t.start()
    t.join(timeout)
    return bool(done)


def _user_settings(remote_prior: object) -> tuple:
    """What an Excel handed to the user gets back, in order: everything the
    pool turned off, then visible and user-controlled (so it stays running
    when the server lets go of it)."""
    return (("DisplayAlerts", True), ("EnableEvents", True),
            ("ScreenUpdating", True),
            ("IgnoreRemoteRequests", bool(remote_prior) if isinstance(remote_prior, bool) else False),
            ("Visible", True), ("UserControl", True))


def taskkill(pid: int) -> bool:
    """Force-kill one PID. The CALLER guarantees this PID is journal-owned;
    this function never enumerates or guesses at foreign processes."""
    try:
        r = subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           capture_output=True, text=True, timeout=30, **_CHILD)
        return r.returncode == 0
    except Exception:
        return False


# --------------------------------------------------------------- journal

#: How long a journal write waits for a reader that has the file open (another
#: server reading it at start, a status read) before it gives up.
JOURNAL_REPLACE_WAIT_S = 2.0

#: How long an update waits for another server's update of the journal.
JOURNAL_LOCK_WAIT_S = 5.0

#: Updates from the threads of this process, one at a time.
_JOURNAL_THREAD_LOCK = threading.Lock()


@contextmanager
def _journal_file_lock(path: Path):
    """One update of the journal at a time across every server of this user:
    an OS lock on a small file beside the journal, which the system drops if
    its holder dies. Yields whether the lock was had; after
    JOURNAL_LOCK_WAIT_S the update goes ahead without it rather than be lost."""
    lock_path = path.with_name(path.name + ".lock")
    try:
        fd = os.open(lock_path, os.O_RDWR | os.O_CREAT, 0o600)
    except OSError:
        yield False
        return
    locked = False
    try:
        deadline = time.monotonic() + JOURNAL_LOCK_WAIT_S
        while True:
            try:
                if os.name == "nt":
                    import msvcrt
                    os.lseek(fd, 0, 0)
                    msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
                locked = True
                break
            except OSError:
                if time.monotonic() >= deadline:
                    break
                time.sleep(0.005)
        yield locked
    finally:
        if locked:
            try:
                if os.name == "nt":
                    import msvcrt
                    os.lseek(fd, 0, 0)
                    msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(fd, fcntl.LOCK_UN)
            except OSError:
                pass
        os.close(fd)


class PidJournal:
    """A tiny disk journal of spawned PIDs. Survives a client crash so the
    next startup sweep can reclaim owned zombies. JSON: pid -> record.

    Every kitchensink4xl server of the user shares it. A write waits out a
    reader that has it open, each update (read, change, write) holds a lock
    all servers share, so no server loses another's record, and a journal
    that still cannot be written is logged: it is never the user's workbook
    being locked, and the Excels this server started are still reclaimed
    when it exits."""

    def __init__(self, path: str | os.PathLike):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def _load(self) -> dict[str, dict]:
        data = self._read()
        return {} if data is None else data

    def _read(self) -> dict[str, dict] | None:
        """The records; {} when there is no journal or it is not JSON; None
        when it could not be read for now (another program holds it)."""
        deadline = time.monotonic() + JOURNAL_REPLACE_WAIT_S
        while True:
            try:
                text = self.path.read_text("utf-8")
            except FileNotFoundError:
                return {}
            except OSError:
                if time.monotonic() >= deadline:
                    return None
                time.sleep(0.01)
                continue
            try:
                data = json.loads(text or "{}")
            except ValueError:
                return {}
            return data if isinstance(data, dict) else {}

    def _save(self, data: dict[str, dict]) -> bool:
        """Write the journal through a temporary file of this writer's own,
        waiting out a reader that has the journal open."""
        tmp = self.path.with_name(
            f"{self.path.name}.{os.getpid()}-{threading.get_ident()}.tmp")
        try:
            tmp.write_text(json.dumps(data, indent=2), "utf-8")
            deadline = time.monotonic() + JOURNAL_REPLACE_WAIT_S
            while True:
                try:
                    os.replace(tmp, self.path)
                    return True
                except PermissionError:
                    if time.monotonic() >= deadline:
                        raise
                    time.sleep(0.01)
        except OSError as exc:
            _log.warning("ks4xl: the Excel process journal %s could not be written (%s). "
                         "The Excels this server started are still ended when it exits, "
                         "but a crash would leave this change out.", self.path, exc)
            try:
                tmp.unlink()
            except OSError:
                pass
            return False

    def _change(self, change) -> None:
        """Read, change and write the journal under the shared lock. `change`
        edits the records in place and says whether it changed anything."""
        with _JOURNAL_THREAD_LOCK, _journal_file_lock(self.path):
            data = self._read()
            if data is None:
                _log.warning("ks4xl: the Excel process journal %s could not be read; "
                             "this change to it is left out.", self.path)
                return
            if change(data):
                self._save(data)

    def record(self, pid: int) -> None:
        rec = {"spawned_at": time.time(), "status": "owned",
               # The IDENTITY of the Excel process itself. Without
               # it a record was a bare PID, and a sweep that found
               # "some EXCEL.EXE holds this PID" force-killed it:
               # after a PID recycle that Excel is the USER'S, with
               # their unsaved work in it.
               # The owner side below already did exactly this;
               # the Excel side never got it.
               "create_time": process_create_time(pid),
               # The SERVER process that owns this Excel. A record
               # whose owner is gone is a crash/exit leftover and
               # the next startup can reclaim it AT ONCE, instead
               # of waiting out the 15-minute age window while an
               # invisible EXCEL.EXE sits on the machine. A record
               # whose owner is still alive belongs to a concurrent
               # session and is left alone.
               "owner_pid": os.getpid(),
               "owner_started": _own_start_time()}

        def put(data):
            data[str(pid)] = rec
            return True
        self._change(put)

    def forget(self, pid: int) -> None:
        def drop(data):
            return data.pop(str(pid), None) is not None
        self._change(drop)

    def annotate(self, pid: int, **fields) -> None:
        """Add fields to an existing record (nothing when there is none)."""
        def note(data):
            if not isinstance(data.get(str(pid)), dict):
                return False
            data[str(pid)].update(fields)
            return True
        self._change(note)

    def owned_pids(self) -> set[int]:
        """Every Excel PID in the journal, this server's and every other
        kitchensink4xl server's of this user (the journal is one file per
        user). For telling the user's own Excel from ours; NEVER the list of
        what this server may kill (reclaimable_pids is)."""
        return {int(p) for p in self._load()}

    def records(self) -> dict[str, dict]:
        """The raw journal, for the startup sweep and for honest status."""
        return self._load()

    def mine_pids(self) -> set[int]:
        """Journaled Excel PIDs that THIS server process spawned:
        the record names this process as its owner, by PID and creation
        time."""
        out: set[int] = set()
        for pid_s, rec in self.records().items():
            if _record_is_mine(rec):
                try:
                    out.add(int(pid_s))
                except ValueError:
                    pass
        return out

    def abandoned_pids(self) -> set[int]:
        """Owned Excel PIDs whose OWNING SERVER is gone: crash or clean-exit
        leftovers that nobody is going to reap. Gone means its PID is free,
        or now names a process created at another time (Windows recycled the
        PID). Records with a live owner (a concurrent session)
        and records from before owner tracking are left to the age-based
        sweep."""
        out: set[int] = set()
        for pid_s, rec in self.records().items():
            if _owner_gone(rec):
                try:
                    out.add(int(pid_s))
                except ValueError:
                    pass
        return out


def _record_is_mine(rec: object) -> bool:
    """The record names this server process as the Excel's owner."""
    if not isinstance(rec, dict) or rec.get("owner_pid") != os.getpid():
        return False
    recorded = rec.get("owner_started")
    if isinstance(recorded, bool) or not isinstance(recorded, (int, float)):
        return True  # no time recorded: the PID is this process's own
    mine = _own_start_time()
    return mine is None or abs(mine - float(recorded)) <= OWNER_TIME_TOLERANCE


def _owner_gone(rec: object) -> bool:
    """The server named as the record's owner is no longer running: its
    PID is free, or it now names a process created at another
    time. False for this process's own records, for records from before
    owner tracking, and whenever liveness cannot be told."""
    if not isinstance(rec, dict):
        return False
    owner = rec.get("owner_pid")
    if not isinstance(owner, int) or isinstance(owner, bool) or owner == os.getpid():
        return False
    if not _process_alive(owner):
        return True
    recorded = rec.get("owner_started")
    if isinstance(recorded, bool) or not isinstance(recorded, (int, float)):
        return False
    actual = process_create_time(owner)
    return actual is not None and abs(actual - float(recorded)) > OWNER_TIME_TOLERANCE


# --------------------------------------------------- the one reclaim point

def reclaim_journaled_pid(journal: PidJournal, pid: int,
                          rec: object = None) -> str:
    """Reap ONE journal-owned Excel, and the ONLY sanctioned way to do it.
    Every kill path (sweep, force_reclaim, the startup journal replay, the
    restart-transient reap) goes through here so the identity check cannot be
    forgotten in one of them.

    Returns what happened:

      ``"gone"``         no live EXCEL.EXE holds the PID; the record is
                         released.
      ``"killed"``       identity verified (or a legacy record with no
                         recorded identity); taskkilled and released.
      ``"kill_failed"``  still alive after taskkill; the record is KEPT so
                         the next sweep retries (forgetting a live zombie
                         orphans it forever).
      ``"pid_reuse"``    a DIFFERENT Excel now holds this PID. NOTHING is
                         killed and the record is dropped, because the
                         process we journaled is already gone and the one
                         standing there is somebody else's, most likely the
                         user's with unsaved work in it.
      ``"handed_to_user"`` the Excel shows its window (excel_window_shown):
                         the user has it, with a workbook of theirs in it.
                         NOTHING is killed, the record is dropped so no later
                         sweep touches it, and its alerts are put back where
                         it can be reached.
    """
    if rec is None:
        rec = journal.records().get(str(pid))
    if not pid_alive(pid):
        journal.forget(pid)
        return "gone"
    if identity_verdict(pid, rec) == "mismatch":
        _log.warning(
            "ks4xl: journaled Excel PID %s is now a different process "
            "(recorded creation time %s, actual %s); Windows recycled the "
            "PID. Killing nothing and dropping the record.",
            pid, journal_create_time(rec), process_create_time(pid))
        journal.forget(pid)
        return "pid_reuse"
    if excel_window_shown(pid):
        journal.forget(pid)
        reached = _return_to_user(pid, (rec or {}).get("remote_prior")
                                  if isinstance(rec, dict) else None)
        _log.warning(
            "ks4xl: Excel PID %s shows its window, so a workbook was opened "
            "in it from outside this server; it is left running for the "
            "user and dropped from the journal (alerts put back: %s).",
            pid, reached)
        return "handed_to_user"
    if taskkill(pid):
        journal.forget(pid)
        return "killed"
    if not pid_alive(pid):
        journal.forget(pid)
        return "gone"
    return "kill_failed"


# ------------------------------------------- what this server opened in Excel

#: The workbooks this server opened in each Excel it spawned, keyed by the
#: id of the instance's COM object (spawn starts every new instance with an
#: empty set). A workbook open in the pooled Excel that is NOT here was opened
#: by someone else, and the instance is the user's from then on.
_OPENED: dict[int, set[str]] = {}


def workbook_key(full_name: str) -> str:
    """How a workbook open in Excel is recognised: its full path, normalized
    (case, separators), or, for one never saved or named by a web address,
    its name folded."""
    if not full_name:
        return ""
    if os.path.isabs(full_name):
        return os.path.normcase(os.path.normpath(full_name))
    return full_name.casefold()


def note_opened(app: object, full_name: str) -> None:
    """Record that this server opened ``full_name`` in ``app``."""
    key = workbook_key(full_name)
    if key and app is not None:
        _OPENED.setdefault(id(app), set()).add(key)


def opened_in(app: object) -> set[str]:
    return set(_OPENED.get(id(app), ()))


#: Every Excel this server process left running for the user.
HANDED_OVER: list[dict] = []

#: The Excel the current COM job on this thread uses (set by acquire), so a
#: failure can be told apart from Excel having ended under it. Forgotten
#: when the job lets that Excel go (release, hand-over, quit) and when the
#: job ends, so a later job's refusal is never judged by an earlier job's
#: Excel.
_THREAD = threading.local()


def current_pid() -> int | None:
    """The PID of the Excel the COM job in flight on this thread acquired."""
    return getattr(_THREAD, "pid", None)


def forget_current(pid: int | None = None) -> None:
    """This thread's job no longer uses that Excel (any, when pid is None)."""
    if pid is None or getattr(_THREAD, "pid", None) == pid:
        _THREAD.pid = None


# ------------------------------------------------------------ the manager

@dataclass
class Worker:
    app: object  # the win32com Excel.Application COM object
    pid: int
    busy: bool = False
    #: The instance's IgnoreRemoteRequests before the pool set it True (the
    #: user's own Excel option). Put back before any Quit, since a clean Quit
    #: writes Excel's options (MEASURED); a killed Excel writes nothing.
    remote_prior: object = None


@dataclass
class SweepResult:
    checked: list[int] = field(default_factory=list)
    killed: list[int] = field(default_factory=list)
    exited_on_own: list[int] = field(default_factory=list)
    still_waiting: list[int] = field(default_factory=list)
    skipped_foreign: list[int] = field(default_factory=list)
    #: Journaled PIDs that Windows recycled onto a different Excel. Nothing
    #: was killed for these and their records were dropped.
    pid_reuse_dropped: list[int] = field(default_factory=list)
    #: Excels left running for the user because they show their window: a
    #: workbook was opened in them from outside this server.
    handed_to_user: list[int] = field(default_factory=list)


#: msoAutomationSecurityForceDisable. Excel started for automation defaults
#: to msoAutomationSecurityLow (1), which enables every macro in a workbook
#: the server opens, with no prompt. Forcing it to ForceDisable (3) disables
#: the VBA project and every auto-run macro (Workbook_Open, Auto_Open) for
#: every open this instance makes. It is a property of our private, hidden,
#: single-session worker, so nothing the user sees or keeps is changed; a
#: fresh worker is spawned for every pool, so there is nothing to restore.
#: It does NOT cover Excel 4.0 (XLM) macro sheets -- Microsoft documents that
#: exclusion -- which com.session.open_workbook refuses before Excel opens
#: the file.
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def _force_disable_active_content(app) -> None:
    """Disable VBA and auto-run macros for every workbook this worker opens.
    Best-effort: a build that does not expose AutomationSecurity leaves the
    other suppressions (EnableEvents=False, and the XLM refusal at the open
    choke point) in force."""
    try:
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    except Exception:  # noqa: BLE001
        pass


class ExcelInstanceManager:
    """Owns worker EXCEL.EXE instances by PID. Never touches a PID it did not
    spawn. Safe to construct on any platform (no COM at construction)."""

    def __init__(self, journal_path: str | os.PathLike | None = None):
        if journal_path is None:
            journal_path = Path(os.environ.get("TEMP", ".")) / \
                "ks4xl_pid_journal.json"
        self.journal = PidJournal(journal_path)
        self._workers: list[Worker] = []
        self._pool: Worker | None = None
        #: Excels this server left running for the user (pid, why, when);
        #: one list for the process, so it outlives a re-armed manager.
        self.handed_over: list[dict] = HANDED_OVER

    # --- spawn / identify -------------------------------------------------

    def _new_app(self):
        import win32com.client as win32  # lazy
        app = win32.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        try:
            app.EnableEvents = False
        except Exception:
            pass
        _force_disable_active_content(app)
        return app

    def spawn(self) -> Worker:
        """DispatchEx a private worker, identify its PID by process-set diff,
        journal it, and track it. The diff is authoritative for invisible,
        workbook-less instances (Hwnd is 0 for them)."""
        before = list_excel_pids()
        app = self._new_app()
        prior = self._ignore_dde(app)
        pid = self._resolve_pid(app, before)
        self.journal.record(pid)
        if prior is not None:
            self.journal.annotate(pid, remote_prior=prior)
        _OPENED[id(app)] = set()
        w = Worker(app=app, pid=pid, remote_prior=prior)
        self._workers.append(w)
        return w

    @staticmethod
    def _ignore_dde(app) -> bool | None:
        """Make the instance ignore DDE, so a workbook another program or
        the shell opens through DDE never lands in this hidden Excel (where
        the user would work without knowing, and the server's exit, a COM
        timeout or a sweep would discard it). Returns the value it had, the
        user's own Excel option, to put back before any Quit; None when it
        could not be read or set."""
        try:
            prior = bool(app.IgnoreRemoteRequests)
            app.IgnoreRemoteRequests = True
            return prior
        except Exception:  # noqa: BLE001
            return None

    def _resolve_pid(self, app, before: set[int]) -> int:
        # Primary: the PID that appeared across the spawn.
        deadline = time.monotonic() + 10.0
        while time.monotonic() < deadline:
            new = list_excel_pids() - before
            if len(new) == 1:
                return next(iter(new))
            if len(new) > 1:
                # Concurrency: fall back to Hwnd mapping to disambiguate.
                hpid = self._pid_via_hwnd(app)
                if hpid in new:
                    return hpid
                # Cannot disambiguate safely; take the lowest new PID and
                # journal all new ones so none is orphaned.
                for p in new:
                    self.journal.record(p)
                return sorted(new)[0]
            time.sleep(0.05)
        # Last resort: Hwnd mapping (needs a window; may be 0).
        hpid = self._pid_via_hwnd(app)
        if hpid:
            return hpid
        raise RuntimeError("could not resolve spawned EXCEL.EXE PID")

    @staticmethod
    def _pid_via_hwnd(app) -> int:
        try:
            import win32process  # lazy
            hwnd = int(app.Hwnd)
            if not hwnd:
                return 0
            _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            return int(pid)
        except Exception:
            return 0

    # --- user-instance location (self-exclusion) --------------------------

    def owned_pids(self) -> set[int]:
        """Every Excel a kitchensink4xl server of this user spawned: this
        one's workers plus every journal record. Used to tell the user's own
        Excel from ours (find_user_instance); never to decide what to kill."""
        live = {w.pid for w in self._workers}
        return live | self.journal.owned_pids()

    def reclaimable_pids(self) -> set[int]:
        """The Excel PIDs THIS server may reclaim: its own
        workers, the journal records it owns, and records whose owning
        server is gone. The journal is one file per user, shared by every
        kitchensink4xl server; reclaiming every record in it made one
        server's exit or COM timeout kill another live server's pooled
        Excel in the middle of its call (MEASURED)."""
        live = {w.pid for w in self._workers}
        return live | self.journal.mine_pids() | self.journal.abandoned_pids()

    def find_user_instance(self):
        """GetActiveObject returns the OLDEST registrant, which may be one of
        our workers. Return (app, pid) for the user's instance, or None if the
        only attachable instance is self-spawned. NEVER assume GetActiveObject
        means the user.

        Process-set guard first: if EVERY live EXCEL.EXE is journal-owned there
        is no user instance, full stop. This is robust even when a worker's
        Hwnd is 0 (invisible, workbook-less workers have no window, so the
        Hwnd->PID self-exclusion cannot see them; the PID-set check can). Note
        a deferred-exit worker/user still counts as live until its process
        actually goes, so a lingering PID is honestly reported as present."""
        candidates = list_excel_pids() - self.owned_pids()
        if not candidates:
            return None
        try:
            import win32com.client as win32  # lazy
            app = win32.GetActiveObject("Excel.Application")
        except Exception:
            return None
        pid = self._pid_via_hwnd(app)
        if pid and pid in self.owned_pids():
            # the oldest registrant is a worker, but a non-owned instance
            # exists; we cannot attach to it specifically via GetActiveObject,
            # so report its PID without a worker handle (never hand back a
            # worker as if it were the user).
            return (None, sorted(candidates)[0])
        return (app, pid)

    # --- pooling ----------------------------------------------------------

    def acquire(self) -> Worker:
        """Reuse the pooled worker if idle, else spawn and pool one. A pooled
        worker that holds a workbook this server did not open, or that has
        been made visible, is the user's now: it is handed over (never
        reused, never closed) and a fresh worker is spawned."""
        w = self._pool
        if w is not None and pid_alive(w.pid):
            why = self.users_claim(w)
            if why is None:
                w.busy = True
                _THREAD.pid = w.pid
                return w
            if why:
                self.hand_over(w, why)
            else:
                # Excel did not answer: never reuse it, and never end it
                # here; the exit's reclaim decides, by its window.
                self._pool = None
        w = self.spawn()
        self._pool = w
        w.busy = True
        _THREAD.pid = w.pid
        return w

    def release_to_pool(self, w: Worker) -> None:
        w.busy = False
        forget_current(w.pid)

    # --- the user's claim -------------------------------------------------

    def users_claim(self, w: Worker) -> str | None:
        """Why this worker is the user's now, or None when it is still only
        this server's. "" when Excel could not be asked. COM: call it on the
        worker thread. The user's claim is a workbook open in it that this
        server did not open, or the instance shown (Visible or
        UserControl), which Excel does when it opens a workbook through DDE."""
        try:
            app = w.app
            mine = opened_in(app)
            books = app.Workbooks
            foreign = []
            for i in range(1, int(books.Count) + 1):
                try:
                    full = str(books.Item(i).FullName or "")
                except Exception:  # noqa: BLE001
                    full = ""
                if not full or workbook_key(full) not in mine:
                    foreign.append(os.path.basename(full) if full else "an unsaved workbook")
            if foreign:
                return ("it holds a workbook this server did not open: "
                        + ", ".join(foreign[:3])
                        + ("" if len(foreign) <= 3 else f" (+{len(foreign) - 3} more)"))
            if bool(app.Visible) or bool(app.UserControl):
                return "it was made visible"
            return None
        except Exception:  # noqa: BLE001
            return ""

    def hand_over(self, w: Worker, why: str) -> None:
        """Leave a worker to the user: put back everything the pool turned
        off (alerts, so closing a changed workbook asks to save; events;
        screen updating; the user's DDE option), show it and keep it running
        when the server lets go, then drop it from the pool, the workers and
        the journal, so no exit, timeout or sweep of any server touches it.
        COM: call it on the worker thread."""
        forget_current(w.pid)
        app = w.app
        for attr, val in _user_settings(w.remote_prior):
            try:
                setattr(app, attr, val)
            except Exception:  # noqa: BLE001
                pass
        _OPENED.pop(id(app), None)
        w.app = None
        if w in self._workers:
            self._workers.remove(w)
        if self._pool is w:
            self._pool = None
        self.journal.forget(w.pid)
        self.handed_over.append({"pid": w.pid, "why": why, "at": time.time()})
        _log.warning("ks4xl: Excel PID %s is left running for the user: %s. "
                     "It is no longer this server's.", w.pid, why)

    # --- teardown ---------------------------------------------------------

    def quit(self, w: Worker) -> None:
        """Ask a worker to Quit and drop our reference. Exit may DEFER by
        minutes (exp 7); do not verify here. The sweep reclaims stragglers.

        A worker the user has a claim on (users_claim) is handed over
        instead: Quit with alerts off would close their workbook unsaved
        without asking. Before a Quit the user's DDE option is put back,
        since a clean Quit writes Excel's options."""
        forget_current(w.pid)
        why = self.users_claim(w) if w.app is not None else None
        if why:
            self.hand_over(w, why)
            return
        if isinstance(w.remote_prior, bool):
            try:
                w.app.IgnoreRemoteRequests = w.remote_prior
            except Exception:  # noqa: BLE001
                pass
        try:
            w.app.Quit()
        except Exception:
            pass
        _OPENED.pop(id(w.app), None)
        try:
            # Releasing the reference lets Excel self-exit on a clean path.
            w.app = None
        except Exception:
            pass
        if w in self._workers:
            self._workers.remove(w)
        if self._pool is w:
            self._pool = None

    def sweep(self, grace_seconds: float = DEFAULT_GRACE_SECONDS,
              poll: float = 2.0) -> SweepResult:
        """Two-phase zombie sweep. Wait out the grace window; any JOURNAL-OWNED
        PID still alive past it gets taskkilled. Foreign PIDs are recorded and
        skipped, never touched."""
        res = SweepResult()
        owned = self.reclaimable_pids()  # never another server's
        res.checked = sorted(owned)
        deadline = time.monotonic() + grace_seconds
        pending = set(owned)
        while pending and time.monotonic() < deadline:
            for pid in list(pending):
                if not pid_alive(pid):
                    res.exited_on_own.append(pid)
                    self.journal.forget(pid)
                    pending.discard(pid)
            if pending:
                time.sleep(poll)
        # Grace elapsed: taskkill any owned PID still alive, identity first.
        for pid in sorted(pending):
            self._record_reclaim(res, pid, reclaim_journaled_pid(
                self.journal, pid))
        return res

    def _record_reclaim(self, res: SweepResult, pid: int, outcome: str) -> None:
        {"gone": res.exited_on_own, "killed": res.killed,
         "kill_failed": res.still_waiting,
         "pid_reuse": res.pid_reuse_dropped,
         "handed_to_user": res.handed_to_user}[outcome].append(pid)
        if outcome == "handed_to_user":
            for w in [w for w in self._workers if w.pid == pid]:
                _OPENED.pop(id(w.app), None)
                w.app = None
                self._workers.remove(w)
                if self._pool is w:
                    self._pool = None
            self.handed_over.append({"pid": pid, "why": "it shows its window",
                                     "at": time.time()})

    def force_reclaim(self) -> SweepResult:
        """Immediate reclaim for tests / shutdown: taskkill every PID this
        server may reclaim (reclaimable_pids: its own and abandoned ones)
        still alive with NO grace wait (used when we accept the deferred-exit
        cost is not worth waiting for). Foreign PIDs untouched, another live
        server's Excel untouched, and a PID Windows recycled onto
        somebody else's Excel is dropped, not killed."""
        res = SweepResult()
        owned = self.reclaimable_pids()
        res.checked = sorted(owned)
        records = self.journal.records()
        for pid in sorted(owned):
            self._record_reclaim(res, pid, reclaim_journaled_pid(
                self.journal, pid, records.get(str(pid))))
        return res


__all__ = [
    "ExcelInstanceManager", "Worker", "SweepResult", "PidJournal",
    "list_excel_pids", "pid_alive", "taskkill", "DEFAULT_GRACE_SECONDS",
    "process_create_time", "journal_create_time", "identity_verdict",
    "reclaim_journaled_pid", "CREATE_TIME_TOLERANCE",
]
