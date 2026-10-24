"""One deadline per call (the same design as the PDF server's core/deadline.py).

The dispatcher gives every call ONE Deadline and runs the op in a worker
thread under it. The ops run in process, so a thread that overran its bound
cannot be killed. What CAN be guaranteed is the invariant

    a TIMEOUT answer means nothing was saved, and a call that saved always
    reports it.

The save is the point of no return. Immediately before its first change on
disk (the backup rotation, then the atomic replace), the safe-save routine
(WorkbookPackage.save, the lifecycle atomic write behind create_workbook
and copy_workbook, manage_backups restore, and the text/image export writers)
calls
`commit_or_raise()` on the active deadline. The dispatcher's wall-clock
backstop calls `cancel()` before it answers TIMEOUT. The two are mutually
exclusive under one lock:

* cancel first: the commit fails and the writer raises TIMEOUT without
  touching the file, so "nothing was saved" is true;
* commit first: cancel fails and the backstop waits for the real result
  (a rename and a return), so a call that saved never reports that it did
  not.

Outside a dispatched call (the in-process test convention, scripts) the
active deadline is unbounded and commit always succeeds.

EVERY FINAL WRITE COMMITS (the same design as the Word editor). Not only the
safe save: every site that changes the disk in a way the caller will see
calls `commit_or_raise()` immediately before its first change. That covers
the writers of NEW files (exports, the image
extract, manage_backups snapshot), manage_backups purge (a delete), and
every save made through Excel itself (the xl_app mutation save, the
encrypted save, SaveAs for conversion, ExportAsFixedFormat for PDF,
Chart.Export for a render). The COM worker thread runs each job under a
copy of the submitting call's context (com/session.py ComExecutor.submit),
so Excel's saves see the same deadline: once the call has answered
TIMEOUT, Excel never saves. A commit is idempotent: a call that committed
once is past its point of no return and later writes in the same call
proceed. A unit test enumerates the write sites
statically and fails on one that does not commit first.

Excel has no live route: the COM tier only ever edits a private hidden copy
it opened itself, so an edit that never reached its save is discarded with
that instance, not left in a document the user can see.

AFTER A COMMIT THE WAIT IS STILL BOUNDED. The backstop waits for a
committed call for up to `post_commit_wait()` more seconds; if even that
runs out (Excel hung inside its own Save), the answer is TIMEOUT with
saved: null and a message that says the save had started and the file must
be checked. It never claims "nothing was saved" for a call that committed.
"""

from __future__ import annotations

import contextvars
import math
import threading
import time
from contextlib import contextmanager
from typing import Iterator

NOTHING_SAVED = "nothing was saved"

#: Seconds the backstop keeps waiting for a call whose save committed,
#: beyond its bound and grace, at the least (the bound itself when larger).
MIN_POST_COMMIT_WAIT_S = 60.0


class Deadline:
    """A wall-clock budget with a cancel/commit handshake."""

    def __init__(self, seconds: float | None) -> None:
        self.bound = seconds
        self._expires = math.inf if seconds is None else time.monotonic() + float(seconds)
        self._lock = threading.Lock()
        self._cancelled = False
        self._committed = False

    @property
    def bounded(self) -> bool:
        return self.bound is not None

    @property
    def cancelled(self) -> bool:
        return self._cancelled

    @property
    def committed(self) -> bool:
        return self._committed

    def post_commit_wait(self) -> float:
        """How much longer the backstop waits for a committed call."""
        return max(MIN_POST_COMMIT_WAIT_S, self.bound or 0.0)

    def remaining(self) -> float:
        """Seconds left; 0 once expired or cancelled; inf when unbounded."""
        if self._cancelled:
            return 0.0
        return max(0.0, self._expires - time.monotonic())

    def expired(self) -> bool:
        return self._cancelled or self.remaining() <= 0.0

    def cancel(self) -> bool:
        """The backstop gave up. False when a save already committed."""
        with self._lock:
            if self._committed:
                return False
            self._cancelled = True
            return True

    def try_commit(self) -> bool:
        """The writer is about to change the disk. False when cancelled or out
        of time. Idempotent: once committed, the call is past its point of no
        return and every later write in it proceeds."""
        with self._lock:
            if self._committed:
                return True
            if self._cancelled or time.monotonic() >= self._expires:
                return False
            self._committed = True
            return True

    def timeout_error(self, what: str = "the call"):
        from .errors import SurfaceRefusal
        bound = f"{self.bound:g}s " if self.bounded else ""
        if self._committed:
            # Only reached when a committed save outran the post-commit wait.
            return SurfaceRefusal(
                "TIMEOUT", f"{what} ran past its {bound}limit while saving; the save had "
                "started and may have finished, so check the file before retrying",
                details={"bound_s": self.bound, "saved": None})
        return SurfaceRefusal(
            "TIMEOUT", f"{what} ran past its {bound}limit and was stopped; {NOTHING_SAVED}",
            details={"bound_s": self.bound, "saved": False})

    def commit_or_raise(self, what: str = "the call") -> None:
        if not self.try_commit():
            raise self.timeout_error(what)


_CURRENT: contextvars.ContextVar[Deadline | None] = contextvars.ContextVar(
    "ks4xl_deadline", default=None)


def current() -> Deadline:
    """The active call's deadline, or an unbounded one outside a call."""
    d = _CURRENT.get()
    return d if d is not None else Deadline(None)


@contextmanager
def active(d: Deadline) -> Iterator[Deadline]:
    token = _CURRENT.set(d)
    try:
        yield d
    finally:
        _CURRENT.reset(token)


def commit_or_raise(what: str = "the save") -> None:
    """The safe-save hook: commit the active call's deadline, or refuse."""
    d = _CURRENT.get()
    if d is not None:
        d.commit_or_raise(what)
