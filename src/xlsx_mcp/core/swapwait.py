"""Reading a workbook while a save swaps it in.

A save puts the new version of a workbook in place with ReplaceFileW (see
``promote`` in the save module). ReplaceFileW moves the workbook's old file
aside, under a name beside it, and then moves the new file in, so for a few
milliseconds the workbook's path is missing. A call that looks in that
instant waits here instead of answering that the workbook does not exist.
Office's own saves leave the same short gap.

A save that is stopped between the two moves (a crash, a client that ends
its server) leaves the workbook under the aside name, whose words say what
it holds. The answer for a missing workbook then says where it is. Reads
never move it back; the next change or the next start of the server does.

This module only looks: it lists a folder and asks whether a path exists.
Every route that reads a workbook reaches it, the read-only tools included,
so nothing here may change a file.
"""

from __future__ import annotations

import hashlib
import os
import re
import time

from .errors import WorkbookNotFound

#: Pauses while a save has the old workbook aside and the path is missing
#: for an instant (about 0.9 s at most).
_SWAP_POLL = (0.002, 0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.5)

#: The aside name earlier versions gave the old file:
#: ".ks4xl-write-<32 hex>.replaced". It does not say which workbook it was.
_LEGACY_ASIDE = re.compile(r"\.ks4xl-write-[0-9a-f]{32}\.replaced")


def aside_prefix(workbook_name: str) -> str:
    """The start of the name a save gives a workbook's old file while the
    new one is moved in, for this workbook name only."""
    tag = hashlib.sha1(
        os.path.normcase(workbook_name).encode("utf-8", "surrogatepass")
    ).hexdigest()[:8]
    return f".ks4xl-replaced-{tag}-"


def _utf16_head(text: str, units: int) -> str:
    """The longest start of ``text`` that fits in ``units`` UTF-16 code units
    (Windows counts a name's length in those)."""
    out, used = [], 0
    for ch in text:
        used += 2 if ord(ch) > 0xFFFF else 1
        if used > units:
            break
        out.append(ch)
    return "".join(out)


def aside_name(workbook_name: str, ident: str) -> str:
    """The name a save gives the workbook's old file while the new one is
    moved in. It starts with the prefix that marks it as this workbook's,
    and then says in words what it holds, because a save that is stopped at
    that instant leaves the workbook under this name:
    ".ks4xl-replaced-<tag>-<id> Budget.xlsx (as it was before a save).tmp".
    A long workbook name is cut short so that the whole name stays inside
    Windows' 255-unit limit. The .tmp ending keeps indexers, previewers and
    sync clients away from a file that exists for milliseconds."""
    return (f"{aside_prefix(workbook_name)}{ident} "
            f"{_utf16_head(workbook_name, 120)} (as it was before a save).tmp")


def _extended(path: str) -> str:
    """``path`` in the form Windows needs past 260 characters."""
    if os.name != "nt" or len(path) <= 240 or path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + path


def _listing(folder: str) -> list[str]:
    try:
        return os.listdir(_extended(folder or "."))
    except OSError:
        return []


def asides(path: str | os.PathLike) -> list[str]:
    """This workbook's old files that a save has aside, newest first.
    Normally none, or one for the few milliseconds of a swap. One that stays
    means a save did not finish: it holds the workbook as it was before
    that save."""
    folder, name = os.path.split(os.path.abspath(os.fspath(path)))
    prefix = aside_prefix(name)
    found = [os.path.join(folder, entry) for entry in _listing(folder)
             if entry.startswith(prefix)]

    def stamp(candidate: str) -> int:
        try:
            return os.stat(_extended(candidate)).st_mtime_ns
        except OSError:
            return 0

    return sorted(found, key=lambda c: (stamp(c), c), reverse=True)


def legacy_asides(path: str | os.PathLike) -> list[str]:
    """Old files that a save by an earlier version left aside in this
    workbook's folder. Their names do not say which workbook
    they were, so they are only named, never moved on a guess."""
    folder = os.path.dirname(os.path.abspath(os.fspath(path)))
    return sorted(os.path.join(folder, entry) for entry in _listing(folder)
                  if _LEGACY_ASIDE.fullmatch(entry))


def aside_in_place(path: str | os.PathLike) -> bool:
    """Whether a save has this workbook's old file aside right now (or an
    earlier version's save has an old file aside in its folder).

    ReplaceFileW moves the old file to a backup name and then moves the new
    one into its place, so the path itself is missing for an instant. The
    backup name is there for exactly that instant (and until it is
    removed), which is how a reader tells a workbook that is being swapped
    from one that is not there."""
    folder, name = os.path.split(os.path.abspath(os.fspath(path)))
    prefix = aside_prefix(name)
    return any(entry.startswith(prefix) or _LEGACY_ASIDE.fullmatch(entry)
               for entry in _listing(folder))


def exists_after_swap(path: str | os.PathLike) -> bool:
    """Whether the workbook is at its path, waiting out another save's swap.

    True at once for a workbook that is there. For a missing one, True once
    it is back while a save has its old file aside; False at once, with no
    wait, when no save has it aside, so a workbook that is not there is
    reported as such straight away."""
    target = _extended(os.path.abspath(os.fspath(path)))
    for pause in (0.0, *_SWAP_POLL):
        if pause:
            time.sleep(pause)
        if os.path.exists(target):
            return True
        if not aside_in_place(path):
            return os.path.exists(target)
    return os.path.exists(target)


def workbook_not_found(path: str | os.PathLike, first: str) -> WorkbookNotFound:
    """The refusal for a workbook that is not at its path. ``first`` is the
    route's own sentence ("no such workbook: ..."). When a save that did
    not finish left the workbook aside, the answer says where it is and how
    it comes back; a read never moves it (reads change nothing), a change
    or a restart does."""
    target = os.path.abspath(os.fspath(path))
    name = os.path.basename(target)
    found = asides(target)
    if found:
        exc = WorkbookNotFound(
            f"{first}. A save that did not finish left the workbook as it was "
            f"at {found[0]}; its content is safe there. The next change to the "
            "workbook through this server, or the next start of the server, "
            f"puts it back; you can also rename that file to {name}.")
        exc.hint = ("Nothing in the workbook was lost. Make the change you "
                    "wanted (it puts the workbook back first), restart the "
                    "server, or rename the file named in the message back to "
                    "the workbook's name.")
        return exc
    old = legacy_asides(target)
    if old:
        listed = ", ".join(old[:3]) + (f" and {len(old) - 3} more" if len(old) > 3 else "")
        exc = WorkbookNotFound(
            f"{first}. A save by an earlier version that did not finish left "
            f"a workbook aside in this folder: {listed}. If it is {name}, "
            f"rename it back to {name}; nothing removes it.")
        exc.hint = ("Open the file named in the message to check it is this "
                    "workbook, then rename it back to the workbook's name.")
        return exc
    return WorkbookNotFound(first)


def require_workbook(path: str | os.PathLike, first: str) -> None:
    """The one check every route makes before it reads a workbook: the
    workbook is at its path, or comes back from another save's swap within
    the wait. Otherwise WorkbookNotFound, with ``first`` as its first
    sentence and, when a save that did not finish left the workbook aside,
    where it is."""
    if exists_after_swap(path):
        return
    raise workbook_not_found(path, first)


def through_swap(path: str | os.PathLike, read_once):
    """``read_once()``, and once more if the workbook went missing under it
    because another save swapped it in at that very instant and it came
    back."""
    try:
        return read_once()
    except FileNotFoundError:
        if not exists_after_swap(path):
            raise
    return read_once()


__all__ = ["aside_prefix", "aside_name", "asides", "legacy_asides",
           "aside_in_place", "exists_after_swap", "workbook_not_found",
           "require_workbook", "through_swap"]
