"""Release an advisory lockfile so that it never outlives its holder.

The cross-process write lock is an advisory file: ``write.lock`` in a
workbook's slot folder (core/safesave.py). A server waiting for it reads the
file on every poll. On Windows an ordinary open does not share DELETE, so a
holder whose release lands inside a waiter's read cannot delete the file: the
unlink fails with a sharing violation (WinError 32). The release used to
swallow that error. The file stayed with the holder's token in it, the
holder's next acquisition took it for a nested hold of its own and never
released it again, and every other server was refused for that workbook until
the holder process exited.

Three rules close it, and the lock applies all three:

1. Retry. ``release`` retries a failed unlink with a short backoff. A
   waiter's read lasts microseconds, so a later attempt succeeds.
2. Verify. ``release`` answers whether anything of ours is still there,
   instead of assuming the unlink worked.
3. Never leave a holder. A lockfile still ours after the retries goes to a
   ``Sweeper`` thread that keeps trying while the process lives, and the
   outermost acquisition in the process adopts a lockfile carrying its own
   token (no hold in this process is open, so it can only be a leftover) and
   releases it when that hold ends.

Readers elsewhere (antivirus, backup and sync clients) can hold the file open
for longer than a waiter does, which is why the sweeper exists at all.

A sweep retires only the queue entry it worked on, and does it while it still
holds the lock's in-process mutex. Letting go of the mutex first and then
removing the entry by path would be a defect: a new
hold in this process could run in between, forget the old entry, fail its own
release and queue a new entry for the same lockfile, and the old sweep then
removed THAT entry because the path matched. The new lockfile would be left
with nobody to clean it up, and another server would be refused until this
process held that lock again. So each entry carries a generation number, and
a sweep retires an entry only if the generation is still the one it
processed.

Breaking a lockfile whose holder is gone has one winner (``break_stale``).
Two waiters that had both read a dead holder's lockfile used to both delete
"the lockfile", and the slower one deleted the LIVE lock the faster one had
just published, so both held the lock. A breaker now works under an OS lock
on a guard file beside the lockfile, reads the lockfile again, and deletes it
only while it is still the very file it judged.

The module is the same code in every KitchenSink4AI editor that keeps such a
lockfile; only this docstring names the XL lock.
"""

from __future__ import annotations

import errno
import os
import sys
import threading
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Callable

#: Pauses between unlink attempts in one release, about 0.8 s in all.
RETRY_DELAYS: tuple[float, ...] = (0.01, 0.02, 0.05, 0.1, 0.2, 0.4)

#: The sweeper's first pause and its longest one.
SWEEP_FIRST_S = 0.05
SWEEP_MAX_S = 2.0

#: ``deletable(lock_path)``: the lockfile is there and this process may
#: remove it (it carries our token, or it cannot be read at all).
Deletable = Callable[[Path], bool]


def present(lock_path: Path) -> bool:
    """Whether the lockfile's name is there. Never raises: a file that is
    being deleted can make stat fail, and pathlib's exists() re-raises that
    on Python 3.12."""
    return os.path.lexists(lock_path)


def unlink_once(lock_path: Path, deletable: Deletable) -> bool:
    """One attempt. True once nothing of ours is left at ``lock_path``."""
    if not deletable(lock_path):
        return True
    try:
        lock_path.unlink()
    except FileNotFoundError:
        return True
    except OSError:
        return not deletable(lock_path)
    return True


def release(lock_path: Path, deletable: Deletable) -> bool:
    """Delete ``lock_path`` while it is ours, retrying a failed unlink.

    Returns True once nothing of ours is left there (deleted, gone already,
    or now another holder's), and False when the file is still ours after
    every retry, so the caller can hand it to a ``Sweeper``.
    """
    for delay in (0.0, *RETRY_DELAYS):
        if delay:
            time.sleep(delay)
        if unlink_once(lock_path, deletable):
            return True
    return False


class Sweeper:
    """Keeps deleting lockfiles a release could not, while the process lives.

    Each entry carries the in-process mutex of its lock. The sweeper only
    acts while it holds that mutex, taken without waiting, so it never
    touches a lockfile that a hold in this process is using; that hold
    adopts the file and releases it itself. A finished entry is retired
    inside that same hold, and only by its generation, so a newer entry for
    the same lockfile is never taken for the one that was just swept.
    """

    def __init__(self, name: str) -> None:
        self._name = name
        self._guard = threading.Lock()
        #: key -> (generation, lock path, mutex, deletable)
        self._items: dict[str, tuple[int, Path, threading.RLock, Deletable]] = {}
        self._generation = 0
        self._thread: threading.Thread | None = None

    def add(
        self, key: str, lock_path: Path, mutex: threading.RLock, deletable: Deletable
    ) -> None:
        with self._guard:
            self._generation += 1
            self._items[key] = (self._generation, lock_path, mutex, deletable)
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(
                    target=self._run, name=self._name, daemon=True
                )
                self._thread.start()

    def forget(self, key: str) -> None:
        with self._guard:
            self._items.pop(key, None)

    def pending(self) -> list[str]:
        """The lockfiles still waiting to be deleted (status and tests)."""
        with self._guard:
            return [str(item[1]) for item in self._items.values()]

    def _run(self) -> None:
        pause = SWEEP_FIRST_S
        while True:
            time.sleep(pause)
            pause = min(pause * 2, SWEEP_MAX_S)
            with self._guard:
                items = list(self._items.items())
                if not items:
                    self._thread = None
                    return
            for key, (generation, lock_path, mutex, deletable) in items:
                if not mutex.acquire(blocking=False):
                    continue  # a hold in this process is open; it adopts the file
                try:
                    if unlink_once(lock_path, deletable):
                        # Still inside the mutex: no hold in this process can
                        # forget this entry and queue a newer one meanwhile.
                        self._retire(key, generation)
                finally:
                    mutex.release()

    def _retire(self, key: str, generation: int) -> None:
        """Drop the entry a sweep finished, and only that entry: one that has
        been replaced since (a newer generation under the same key and path)
        stays queued."""
        with self._guard:
            current = self._items.get(key)
            if current is not None and current[0] == generation:
                del self._items[key]


# ------------------------------------------------- breaking a stale lockfile

#: The guard beside a lockfile, ``<lockfile>.guard``, that a breaker holds.
GUARD_SUFFIX = ".guard"

#: How long one break waits for the guard; the acquire loop then looks again.
GUARD_WAIT_SECONDS = 1.0


def _os_lock(fd: int) -> bool | None:
    """Take the OS lock on the open guard file without waiting: True when
    taken, False when another breaker holds it, None when this file system
    offers no such lock."""
    try:
        if sys.platform == "win32":
            import msvcrt
            os.lseek(fd, 0, os.SEEK_SET)
            msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        return True
    except OSError as exc:
        if exc.errno in (errno.EACCES, errno.EAGAIN, errno.EWOULDBLOCK,
                         errno.EDEADLK):
            return False
        return None


def _os_unlock(fd: int) -> None:
    try:
        if sys.platform == "win32":
            import msvcrt
            os.lseek(fd, 0, os.SEEK_SET)
            msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            fcntl.flock(fd, fcntl.LOCK_UN)
    except OSError:
        pass  # closing the file drops the lock as well


def _still_the_file(fd: int, guard: Path) -> bool:
    """POSIX: the guard file this process locked is still the one at the
    path (a breaker deletes it while holding the lock, so a lock taken on
    the old file after that is no guard). Windows refuses to delete a file
    that anyone has open, so there it always is."""
    if sys.platform == "win32":
        return True
    try:
        held, named = os.fstat(fd), os.stat(guard)
    except OSError:
        return False
    return (held.st_ino, held.st_dev) == (named.st_ino, named.st_dev)


@contextmanager
def break_guard(lock_path: Path):
    """Hold the guard for breaking ``lock_path``. Yields True while it is
    held; False when another breaker kept it for the whole wait (the caller
    looks again on its next pass); None when the guard file cannot be made
    or this file system has no OS lock, where the caller breaks on the
    re-read alone. The system drops an OS lock when its holder dies, so the
    guard never goes stale, and the guard file is removed after use."""
    guard = lock_path.with_name(lock_path.name + GUARD_SUFFIX)
    end = time.monotonic() + GUARD_WAIT_SECONDS
    got: bool | None = False
    fd: int | None = None
    while True:
        try:
            fd = os.open(str(guard), os.O_RDWR | os.O_CREAT, 0o600)
        except OSError:
            fd, got = None, None
        else:
            got = _os_lock(fd)
            if got and not _still_the_file(fd, guard):
                _os_unlock(fd)
                got = False
            if got:
                break
            os.close(fd)
            fd = None
            if got is None:
                break
        if time.monotonic() >= end:
            break
        time.sleep(0.01)
    if fd is None:
        yield got
        return
    try:
        yield True
    finally:
        if sys.platform != "win32":
            # Gone while still locked: a breaker that opened the old file
            # sees, once it has the lock, that it is no longer the guard.
            try:
                os.unlink(guard)
            except OSError:
                pass
        _os_unlock(fd)
        os.close(fd)
        if sys.platform == "win32":
            # Windows deletes it only when nobody has it open, so a breaker
            # holding it now keeps it; otherwise it is gone, and the next
            # breaker makes a new one.
            try:
                os.unlink(guard)
            except OSError:
                pass


def break_stale(lock_path: Path, judged: dict,
                read_info: Callable[[Path], dict],
                unreadable_grace: float) -> None:
    """Delete a lockfile whose holder is gone, with one winner.

    ``judged`` is what the caller read and judged: the dead holder's record,
    or {} for a file that cannot be read at all, which goes only while it
    still cannot be read and is older than ``unreadable_grace`` seconds.
    Under ``break_guard`` the file is read again (``read_info``) and deleted
    only while it is still that file. Nothing else can change it meanwhile:
    a publisher cannot create over it, its holder is gone, and every other
    breaker waits for the guard. A lockfile published since the caller
    looked is never touched. Renaming the lockfile aside and checking it
    afterwards was not used: a breaker that lost the race would have moved
    a live lock away, and putting it back cannot be made atomic. A failure
    is left to the caller's acquire loop, which looks again."""
    with break_guard(lock_path) as held:
        if held is False:
            return  # another breaker has the guard
        now = read_info(lock_path)
        if judged:
            if now != judged:
                return  # broken already; whatever is there now is not it
        else:
            if now:
                return  # it can be read now: no longer abandoned
            try:
                born = lock_path.stat().st_mtime
            except OSError:
                return  # gone already
            if time.time() - born <= unreadable_grace:
                return
        try:
            lock_path.unlink()
        except OSError:
            pass


__all__ = [
    "GUARD_SUFFIX", "GUARD_WAIT_SECONDS", "RETRY_DELAYS", "SWEEP_FIRST_S",
    "SWEEP_MAX_S", "Sweeper", "break_guard", "break_stale", "present",
    "release", "unlink_once",
]
