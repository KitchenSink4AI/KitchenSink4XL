"""A cross-process lockfile never outlives its holder (2.0.0rc5).

A server waiting for the write lock reads the lockfile on every poll. On
Windows an ordinary read opens the file without FILE_SHARE_DELETE, so a
holder whose release lands inside that read cannot delete it. The old
release swallowed the failed unlink; the file kept the holder's token; the
holder's next acquisition took it for a nested hold of its own
(owns_lockfile False) and never released it again. Every other server was
then refused for that workbook ("held for 12s", "held for 22s", ...) until
the holder process exited, however idle the holder was.

The fix (core/lockrelease.py):
- the release retries a failed unlink and reports whether the file went;
- a lockfile still ours after the retries goes to a sweeper thread;
- the outermost hold in the process adopts a lockfile with its own token.

Opening the lockfile for reading is exactly what a waiting server does, so
the tests below hold such a handle across a release. The two-server test
does it with real stdio servers.
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import textwrap
import threading
import time
from pathlib import Path

import pytest

import family_harness as fh
from xlsx_mcp import server as srv
from xlsx_mcp.core import lockrelease, safesave

# The src folder of the xlsx_mcp this test imported, so a child process runs
# the same code as the test.
SRC = Path(safesave.__file__).resolve().parents[2]
IS_WIN = sys.platform == "win32"
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

needs_windows = pytest.mark.skipif(
    not IS_WIN, reason="only Windows refuses to delete a file another process has open"
)


def _book(tmp_path: Path) -> Path:
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"not a workbook; the lock never reads it")
    return book


def _lock_of(book: Path) -> Path:
    return safesave.slot_dir(book, create=True) / safesave.LOCK_FILE_NAME


def _wait_until(predicate, seconds: float) -> bool:
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        if predicate():
            return True
        time.sleep(0.02)
    return predicate()


# ------------------------------------------------------------- file route


@needs_windows
def test_a_release_blocked_by_a_reader_is_retried_until_it_goes(tmp_path):
    """The usual case: a waiter's read is brief, so a retry succeeds and
    nothing is left over, without the sweeper."""
    book = _book(tmp_path)
    lock = _lock_of(book)
    with safesave.write_lock(book):
        reader = open(lock, "rb")
        threading.Timer(0.03, reader.close).start()
    assert not lock.exists()
    assert safesave._SWEEPER.pending() == []


@needs_windows
def test_the_next_hold_adopts_a_lockfile_the_release_could_not_delete(tmp_path):
    """The reader stays open through the whole release, so every retry
    fails. The holder's next write must not take the leftover for a nested
    hold; it adopts it and releases it."""
    book = _book(tmp_path)
    lock = _lock_of(book)
    with safesave.write_lock(book):
        reader = open(lock, "rb")
    try:
        assert lock.exists(), "the reader should have blocked every retry"
        assert json.loads(lock.read_text(encoding="utf-8"))["token"] == safesave._OWNER_TOKEN
    finally:
        reader.close()
    with safesave.write_lock(book):
        pass
    assert not lock.exists()
    assert safesave._SWEEPER.pending() == []


@needs_windows
def test_the_sweeper_removes_the_leftover_without_another_write(tmp_path):
    """A holder that never writes that workbook again must not keep everyone
    else out until it exits."""
    book = _book(tmp_path)
    lock = _lock_of(book)
    with safesave.write_lock(book):
        reader = open(lock, "rb")
    assert lock.exists()
    assert str(lock) in safesave._SWEEPER.pending()
    reader.close()
    assert _wait_until(lambda: not lock.exists(), 10.0), "the sweeper never removed it"
    assert _wait_until(lambda: safesave._SWEEPER.pending() == [], 5.0)


_WAITER = textwrap.dedent(
    """
    import sys, time
    sys.path.insert(0, sys.argv[2])
    from xlsx_mcp.core import safesave
    safesave.LOCK_WAIT_SECONDS = float(sys.argv[3])
    t0 = time.monotonic()
    try:
        with safesave.write_lock(sys.argv[1]):
            pass
        print("GOT", round(time.monotonic() - t0, 2), flush=True)
    except safesave.MutationLockTimeout as exc:
        print("REFUSED", exc, flush=True)
    """
)


@needs_windows
def test_another_process_gets_the_workbook_after_a_blocked_release(tmp_path):
    """After a release that could not delete the lockfile, a second process
    asks for the same workbook while the first one stays alive and idle. It
    used to be refused until the first process exited; now it gets in."""
    book = _book(tmp_path)
    lock = _lock_of(book)
    with safesave.write_lock(book):
        reader = open(lock, "rb")
    reader.close()
    proc = subprocess.run(
        [sys.executable, "-c", _WAITER, str(book), str(SRC), "6"],
        capture_output=True, text=True, timeout=60, stdin=subprocess.DEVNULL,
        creationflags=NO_WINDOW,
    )
    assert proc.stdout.startswith("GOT"), proc.stdout + proc.stderr
    assert not lock.exists()


def test_a_nested_hold_still_leaves_the_lockfile_to_the_outer_hold(tmp_path):
    """Adoption is for the outermost hold only; a nested hold in the same
    process must not take the file over and delete it under its parent."""
    book = _book(tmp_path)
    lock = _lock_of(book)
    with safesave.write_lock(book):
        with safesave.write_lock(book):
            assert lock.exists()
        assert lock.exists(), "the nested hold deleted the outer hold's lockfile"
    assert not lock.exists()


def test_the_release_never_deletes_another_holders_lockfile(tmp_path):
    lock = tmp_path / safesave.LOCK_FILE_NAME
    lock.write_text(json.dumps({
        "pid": os.getpid(), "token": "someone-else", "time": time.time(),
    }), encoding="utf-8")
    assert safesave._release_lockfile(lock) is True  # nothing of ours there
    assert lock.exists()


def test_a_release_of_a_lockfile_that_is_already_gone_is_done(tmp_path):
    assert safesave._release_lockfile(tmp_path / safesave.LOCK_FILE_NAME) is True


# ------------------------------------------------------- the sweeper itself


def test_the_sweeper_waits_while_a_hold_in_this_process_is_open(tmp_path):
    """It acts only with the lock's in-process mutex in hand, so it can
    never delete a lockfile that a hold in this process is using."""
    lock = tmp_path / "x.lock"
    lock.write_text("{}", encoding="utf-8")
    mutex = threading.RLock()
    sweeper = lockrelease.Sweeper("test-sweeper")
    mutex.acquire()
    try:
        sweeper.add("k", lock, mutex, lambda p: os.path.lexists(p))
        time.sleep(0.5)
        assert lock.exists(), "the sweeper deleted a lockfile while its mutex was held"
    finally:
        mutex.release()
    assert _wait_until(lambda: not lock.exists(), 10.0)
    assert _wait_until(lambda: sweeper.pending() == [], 5.0)


def test_the_sweeper_leaves_a_lockfile_that_is_no_longer_ours(tmp_path):
    lock = tmp_path / "x.lock"
    lock.write_text("{}", encoding="utf-8")
    sweeper = lockrelease.Sweeper("test-sweeper")
    sweeper.add("k", lock, threading.RLock(), lambda p: False)
    assert _wait_until(lambda: sweeper.pending() == [], 5.0)
    assert lock.exists()


class _HeldAtFirstTry:
    """A lock's mutex whose first try by the sweeper waits until the test
    lets it go on, so the test can change the queue in between."""

    def __init__(self) -> None:
        self.real = threading.RLock()
        self.trying = threading.Event()
        self.go_on = threading.Event()
        self._first = True

    def acquire(self, *args, **kwargs):
        if self._first and threading.current_thread().name == "test-sweeper":
            self._first = False
            self.trying.set()
            assert self.go_on.wait(10), "the test never let the sweeper go on"
        return self.real.acquire(*args, **kwargs)

    def release(self):
        self.real.release()


def test_a_sweep_retires_only_the_entry_it_processed(tmp_path):
    """The sweeper used to retire whatever entry had the same path. A
    newer entry for the same lockfile, added while the sweep was
    under way, was dropped with the old one, and the lockfile it named was
    never deleted."""
    lock = tmp_path / "x.lock"
    lock.write_text("{}", encoding="utf-8")
    mutex = _HeldAtFirstTry()
    sweeper = lockrelease.Sweeper("test-sweeper")
    # The old entry: by the time it is swept, nothing of its own is there.
    sweeper.add("k", lock, mutex, lambda p: False)
    assert mutex.trying.wait(10), "the sweeper never came round"
    # A new hold forgot it and could not delete its own lockfile.
    sweeper.forget("k")
    sweeper.add("k", lock, mutex, lambda p: os.path.lexists(p))
    mutex.go_on.set()
    assert _wait_until(lambda: not lock.exists(), 10.0), \
        "the newer entry was dropped with the old one, and its lockfile stayed"
    assert _wait_until(lambda: sweeper.pending() == [], 5.0)


class _PausedAfterRelease:
    """The sweeper is paused right after it lets the real mutex go, which
    is where the old code still had its bookkeeping to do."""

    def __init__(self) -> None:
        self.real = threading.RLock()
        self.released = threading.Event()
        self.go_on = threading.Event()

    def acquire(self, *args, **kwargs):
        return self.real.acquire(*args, **kwargs)

    def release(self):
        self.real.release()
        if threading.current_thread().name == "ks4xl-write-lock-sweeper" \
                and not self.released.is_set():
            self.released.set()
            assert self.go_on.wait(10), "the test never let the sweeper go on"


@needs_windows
def test_a_newer_leftover_is_still_swept_when_an_old_sweep_finishes(tmp_path, monkeypatch):
    """End to end. A reader blocks the release, so the
    sweeper gets the leftover. When the sweeper has deleted it, a new hold
    publishes the lockfile again and a second reader blocks that release
    too. The old sweep then finishes. The new leftover must still be swept
    once the reader lets go, so another process gets in without this
    process writing again."""
    book = _book(tmp_path)
    lock = _lock_of(book)
    key = safesave.canonical_key(book)
    mutex = _PausedAfterRelease()
    monkeypatch.setitem(safesave._MUTEXES, key, mutex)
    with safesave.write_lock(book):
        reader = open(lock, "rb")
    assert str(lock) in safesave._SWEEPER.pending()
    reader.close()
    assert mutex.released.wait(10), "the sweeper never deleted the first leftover"
    assert not lock.exists()
    with safesave.write_lock(book):
        reader = open(lock, "rb")
    try:
        assert str(lock) in safesave._SWEEPER.pending(), "the second release was not queued"
        mutex.go_on.set()
        time.sleep(0.3)
        assert str(lock) in safesave._SWEEPER.pending(), \
            "the old sweep dropped the newer leftover"
    finally:
        reader.close()
    assert _wait_until(lambda: not lock.exists(), 10.0), "the newer leftover was never swept"
    proc = subprocess.run(
        [sys.executable, "-c", _WAITER, str(book), str(SRC), "3"],
        capture_output=True, text=True, timeout=60, stdin=subprocess.DEVNULL,
        creationflags=NO_WINDOW,
    )
    assert proc.stdout.startswith("GOT"), proc.stdout + proc.stderr
    assert _wait_until(lambda: safesave._SWEEPER.pending() == [], 5.0)


# ------------------------------------------------------ two real servers


class _Server:
    """One real stdio server (the Community launcher), as a client sees it."""

    def __init__(self) -> None:
        self.proc = subprocess.Popen(
            fh.launch_cmd(), stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL, env=fh._env(None), creationflags=NO_WINDOW)
        self._lines: queue.Queue = queue.Queue()
        self._id = 0
        threading.Thread(target=self._pump, daemon=True).start()
        self._request("initialize", {
            "protocolVersion": "2025-06-18", "capabilities": {},
            "clientInfo": {"name": "lock-release", "version": "0"}})
        self._send({"jsonrpc": "2.0", "method": "notifications/initialized"})

    def _pump(self) -> None:
        for line in iter(self.proc.stdout.readline, b""):
            if line.strip():
                self._lines.put(line)

    def _send(self, msg: dict) -> None:
        self.proc.stdin.write((json.dumps(msg) + "\n").encode("utf-8"))
        self.proc.stdin.flush()

    def _request(self, method: str, params: dict, timeout: float = 120.0) -> dict:
        self._id += 1
        rid = self._id
        self._send({"jsonrpc": "2.0", "id": rid, "method": method, "params": params})
        end = time.monotonic() + timeout
        while True:
            msg = json.loads(self._lines.get(timeout=max(0.1, end - time.monotonic())))
            if msg.get("id") == rid:
                return msg

    def call(self, tool: str, op: str, args: dict) -> tuple[dict, bool]:
        res = self._request("tools/call", {"name": tool, "arguments": {"op": op, "args": args}})
        result = res.get("result") or {}
        content = result.get("content") or []
        text = content[0]["text"] if content else json.dumps(res.get("error"))
        try:
            payload = json.loads(text)
        except ValueError:
            payload = {"raw": text}
        return payload, bool(result.get("isError")) or "error" in res

    def close(self) -> None:
        try:
            self.proc.stdin.close()
        except OSError:
            pass
        try:
            self.proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            self.proc.kill()
            self.proc.wait(timeout=30)


def _hold_open_when_it_appears(lock: Path, seconds: float) -> tuple[threading.Event, threading.Thread]:
    """Play a waiting server's poll: open the lockfile the moment it exists
    and keep the handle (no FILE_SHARE_DELETE, like Path.read_text) open for
    `seconds`, which spans the holder's release."""
    caught = threading.Event()

    def run() -> None:
        end = time.monotonic() + 60
        while time.monotonic() < end:
            try:
                handle = open(lock, "rb")
            except OSError:
                time.sleep(0.0005)
                continue
            caught.set()
            try:
                time.sleep(seconds)
            finally:
                handle.close()
            return

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    return caught, thread


@needs_windows
@pytest.mark.timeout(240)
def test_two_servers_a_blocked_release_does_not_lock_the_other_server_out(tmp_path):
    """The defect as a user met it, with two real stdio servers on one
    workbook. Server A writes while a reader holds the lockfile open across
    A's release, then sits idle. Server B must still be able to write; on the
    old code it was refused ("held for 12s") until A exited, and A's own next
    write left the file behind again."""
    book = tmp_path / "shared.xlsx"
    srv.create_workbook(str(book))
    lock = safesave.slot_dir(book) / safesave.LOCK_FILE_NAME
    a = b = None
    try:
        a, b = _Server(), _Server()
        caught, thread = _hold_open_when_it_appears(lock, 2.0)
        payload, err = a.call("xl_edit", "set_cell", {"path": str(book), "location": "A1", "value": 1})
        assert not err, payload
        assert caught.wait(30), "the harness never saw the lockfile, so nothing was tested"
        thread.join(30)
        # A is idle and alive. B asks for the same workbook.
        t0 = time.monotonic()
        payload, err = b.call("xl_edit", "set_cell", {"path": str(book), "location": "A2", "value": 2})
        assert not err, f"B was refused after {time.monotonic() - t0:.1f}s: {payload}"
        # A writes again (it must not trip over a leftover of its own) and B again.
        payload, err = a.call("xl_edit", "set_cell", {"path": str(book), "location": "A3", "value": 3})
        assert not err, payload
        t0 = time.monotonic()
        payload, err = b.call("xl_edit", "set_cell", {"path": str(book), "location": "A4", "value": 4})
        assert not err, f"B was refused after {time.monotonic() - t0:.1f}s: {payload}"
        assert _wait_until(lambda: not lock.exists(), 10.0), "a lockfile was left behind"
    finally:
        for server in (a, b):
            if server is not None:
                server.close()
