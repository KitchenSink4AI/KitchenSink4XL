"""Slot-based backup rotation + per-file write serialization for the save path.

Ported from KitchenSink4Word core/safesave.py, the SAFETY CORE foundation
(backup before every change). The logic is format-agnostic;
only the backup-folder name (.ks4w-backups -> .ks4xl-backups), the slot file
extension (.docx -> .xlsx), the error base (WordMcpError -> XlMcpError), and
the process name in the lock-timeout message changed. Everything else is the
proven two-slot hardlink design.

Backup model:

- A hidden ``.ks4xl-backups/`` folder sits next to each mutated workbook
  (dot-prefix everywhere; FILE_ATTRIBUTE_HIDDEN is additionally set on
  Windows, non-fatal if that fails).
- Inside it, one subfolder per workbook (the workbook's own file name,
  hash-suffixed when the name is very long; Korean/unicode names work).
- Exactly two stable slots per workbook (SLOT_POLICY below):
    ``prev.xlsx``   - state before the most recent mutation, rotates every call
    ``anchor.xlsx`` - session-start state, rotates only when the workbook has
                      been idle ANCHOR_IDLE_SECONDS or more, measured from the
                      prev slot's mtime (no state database).
  (The slot extension is cosmetic: restore copies the slot bytes back onto
  the original path regardless of the source's real extension, so an .xlsm
  source is preserved intact in a prev.xlsx slot.)
- Rotation mechanism (crash-window-free): HARDLINK the current target file to
  a temp name inside the slot folder, then os.replace that link onto the slot.
  The workbook is NEVER absent from its own path at any instant. Hardlink
  failure (non-NTFS, cloud placeholders, cross-volume) falls back to
  shutil.copy2. Link/replace calls are retried with backoff to ride out
  transient AV sharing violations.

Write serialization (fixes the parallel read-modify-save race):

- ``write_lock(path)`` context manager, held across the FULL
  read-modify-validate-save cycle of a mutation.
- In-process: one threading.RLock per file, keyed on
  normcase(realpath(path)).
- Cross-process: an advisory lockfile inside the workbook's slot folder
  carrying PID + a per-process-instance TOKEN + a timestamp minted at
  publication. The file is written complete and linked into place, so it is
  never observable half-written. Stale locks (dead PID, our own recycled PID
  under a foreign token, under a local host identity) are broken;
  otherwise acquisition waits up to LOCK_WAIT_SECONDS and then refuses with
  MutationLockTimeout naming the holder. Release removes the file only while
  it still carries our token.

  The token and the atomic publish close a contention leak reproduced 2 of
  2 times at four or more concurrent writers: with a two-syscall create, a concurrent acquirer read the empty file, saw no
  PID, called it stale, deleted a LIVE lock and took its own; two holders
  then unlinked each other's files, the lock outlived every writer, and the
  refusal named a PID that had never written anything.

  A release that cannot delete the file never leaves it behind
  (core/lockrelease.py): a waiter polling the lockfile holds it open for a
  moment, and on Windows that makes the holder's unlink fail. The release
  retries, then hands a lockfile that is still ours to a sweeper thread, and
  the outermost hold in this process adopts a lockfile carrying our own
  token instead of mistaking it for a nested hold and never releasing it.

  Breaking a stale lockfile has one winner: it happens under
  an OS lock on a guard file beside it, and the breaker deletes the
  lockfile only while it is still the very file it judged
  (core/lockrelease.py break_stale).

- The COM tier holds the same lock around every Excel save
  (ops/comtier.py), in the calling thread. The COM worker thread never takes
  it: it runs jobs for callers that hold it, and a hold there would wait for
  its own caller, so it is marked with forbid_write_lock_on_this_thread and
  a write_lock there fails at once.
"""

from __future__ import annotations

import hashlib
import json
import os
import socket
import shutil
import sys
import threading
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

from . import lockrelease
from .errors import XlMcpError
from .sandbox import check_path
from .swapwait import aside_name

# ---------------------------------------------------------------- constants

BACKUP_DIR_NAME = ".ks4xl-backups"

#: Slot policy: stable slot file names, at most these per workbook ever.
SLOT_POLICY: tuple[str, ...] = ("prev.xlsx", "anchor.xlsx")
PREV_SLOT = "prev.xlsx"
ANCHOR_SLOT = "anchor.xlsx"

#: Idle gap (seconds) after which the next mutation is a "new session" and
#: the anchor slot rotates to the current pre-mutation state.
ANCHOR_IDLE_SECONDS = 60 * 60

#: Advisory lockfile name inside a workbook's slot folder.
LOCK_FILE_NAME = "write.lock"
#: How long an acquirer waits for a live holder before refusing.
LOCK_WAIT_SECONDS = 10.0
#: Historical age threshold, retained for compatibility only.
#: Elapsed time never evicts a live holder.
LOCK_STALE_SECONDS = 10 * 60

#: Slot folder names longer than this are truncated + hash-suffixed
#: (keeps total path length sane for Windows even without long-path opt-in).
_MAX_FOLDER_NAME = 80

#: Breadcrumb written when the folder name had to be truncated, so orphan
#: detection can still map the folder back to its source workbook name.
_SOURCE_NAME_FILE = "source-name.txt"

_RETRY_DELAYS = (0.05, 0.1, 0.2, 0.4)  # seconds, between retries
_TRANSIENT_WINERRORS = {5, 32, 33}  # access denied / sharing violations


class MutationLockTimeout(XlMcpError):
    """Another process holds the write lock on this file and did not release
    it within the wait window; the mutation was refused, nothing changed."""


def file_stamp(path: str | os.PathLike) -> tuple[int, int] | None:
    """(mtime_ns, size) identity of the file on disk, or None if it is gone.

    A save compares it just before it replaces a workbook: the file tier
    with what its edit read (core.package), the COM tier with what Excel
    opened (com.session.put_in_place). The write lock keeps this
    server's writers apart; the stamp catches a program that does not use
    it."""
    try:
        st = os.stat(path)
    except OSError:
        return None
    return (st.st_mtime_ns, st.st_size)


def changed_on_disk_error(path: str | os.PathLike) -> XlMcpError:
    """The refusal when the workbook changed on disk after the operation read
    it (CONFLICT; nothing was written). One wording for the file tier and the
    COM tier."""
    exc = XlMcpError(
        f"{Path(os.fspath(path)).name} changed on disk after this operation "
        "read it (another process or Excel wrote to it). Refusing rather "
        "than overwriting that change with a stale copy; re-run the "
        "operation against the current file.")
    exc.code = "CONFLICT"
    return exc


# ------------------------------------------------------------ path plumbing


def canonical_key(path: str | os.PathLike) -> str:
    """Stable per-file identity: normcase(realpath). Lock keys use this."""
    return os.path.normcase(os.path.realpath(os.path.abspath(os.fspath(path))))


def in_backup_store(path: str | os.PathLike) -> bool:
    """True when the path is inside a .ks4xl-backups store, judged on the
    path as given AND on its resolved form. The lexical
    test alone let a junction (or any other link) into the store pass, so
    copy_workbook overwrite:true through a junction replaced the prev slot.
    realpath resolves junctions and symlinks on the part of the path that
    exists, which is the part a link can live in."""
    name = BACKUP_DIR_NAME.lower()
    ap = os.path.abspath(os.fspath(path))
    if any(part.lower() == name for part in Path(ap).parts):
        return True
    try:
        real = os.path.realpath(ap)
    except (OSError, ValueError):
        return False
    return any(part.lower() == name for part in Path(real).parts)


def refuse_backup_slot_target(path: str | os.PathLike, what: str = "edit") -> None:
    """No op edits a backup in place (the same protected-output rule as the
    Word editor). A mutating op given a slot file as
    its path used to edit the undo copy itself, so the one thing that could
    put the workbook back was gone."""
    if in_backup_store(path):
        raise XlMcpError(
            f"refusing to {what} {Path(os.fspath(path)).name}: it is a backup inside a "
            f"{BACKUP_DIR_NAME} folder, and editing a backup in place destroys the undo it "
            "holds. Put it back with xl_workbook manage_backups (action='restore'), or copy "
            "it out with xl_workbook copy_workbook and edit the copy.")


#: Per-component filename ceilings. Windows counts CHARACTERS, ext4 and
#: friends count BYTES, and a Korean character is three bytes in UTF-8, so a
#: name that is comfortable on one is unwritable on the other. Every name
#: this server CONSTRUCTS is held inside both limits, on every platform, so a
#: workbook that saves on Windows also saves when the same tree is mounted on
#: Linux (and the CI matrix now runs all three). The user's own filename is
#: not this server's to bound; what is bounded is anything built by appending
#: to it, which is where a legal name became an illegal one.
MAX_NAME_CHARS = 255
MAX_NAME_BYTES = 255

#: Room left for the decorations a caller appends after bounding (a temp
#: suffix, a .bak marker). 24 covers the longest this tree adds,
#: ".20260906_153045.bak".
NAME_HEADROOM = 24


def name_fits(name: str, *, headroom: int = 0) -> bool:
    """True when name is inside BOTH per-component limits."""
    return (len(name) + headroom <= MAX_NAME_CHARS
            and len(name.encode("utf-8", "surrogatepass")) + headroom
            <= MAX_NAME_BYTES)


def _truncate_to_bytes(text: str, limit: int) -> str:
    """Longest prefix of text whose UTF-8 encoding fits in limit bytes,
    never splitting a character."""
    if len(text.encode("utf-8", "surrogatepass")) <= limit:
        return text
    out = text
    while out and len(out.encode("utf-8", "surrogatepass")) > limit:
        out = out[:-1]
    return out


def bound_name(name: str, *, headroom: int = 0,
               max_chars: int | None = None) -> str:
    """Fit a CONSTRUCTED filename inside the per-component limits.

    The extension is kept whole (a truncated ".xlsx" is a different file
    type), and the stem is cut and given an 8-hex digest of the full
    original, so two long names that share a prefix never collapse onto one
    file. A name that already fits comes back untouched, which keeps the
    common case trivially reversible.

    headroom reserves bytes for a suffix the caller appends afterwards.
    max_chars applies an additional, tighter character cap (the slot-folder
    rule uses it).
    """
    cap_chars = min(MAX_NAME_CHARS, max_chars or MAX_NAME_CHARS) - headroom
    cap_bytes = MAX_NAME_BYTES - headroom
    if len(name) <= cap_chars and \
            len(name.encode("utf-8", "surrogatepass")) <= cap_bytes:
        return name
    stem, ext = os.path.splitext(name)
    digest = hashlib.sha1(
        os.path.normcase(name).encode("utf-8", "surrogatepass")
    ).hexdigest()[:8]
    tail = f"-{digest}{ext}"
    stem = stem[: max(1, cap_chars - len(tail))]
    stem = _truncate_to_bytes(
        stem, max(1, cap_bytes - len(tail.encode("utf-8", "surrogatepass"))))
    return f"{stem}{tail}"


def _folder_name_for(doc_name: str) -> str:
    """Slot subfolder name for a workbook file name. The workbook's own name
    when it fits (trivial reverse mapping); truncated + 8-hex-hash when long.

    The 80-character rule is this server's own, chosen so a slot folder plus
    everything under it stays comfortably inside a path limit. bound_name
    adds the byte ceiling on top, which is what a Korean workbook name needs
    on Linux: 80 characters of Korean is 240 bytes. No headroom is reserved,
    because nothing is appended to a slot folder name; the files inside it
    carry their own fixed names."""
    return bound_name(doc_name, max_chars=_MAX_FOLDER_NAME)


#: Prefix of every scratch file this server writes beside a workbook: the
#: save's own payload and the scratch copies its helpers make while they
#: rewrite that payload. The save path's stale-temp sweeper removes aged
#: ones (core.package), so a process killed mid-write leaves nothing that
#: lasts (the helpers used to take tempfile's tmpXXXX names, which the
#: sweeper never matched).
WRITE_TMP_PREFIX = ".ks4xl-write-"


def scratch_file(directory: str | os.PathLike, suffix: str) -> tuple[int, str]:
    """tempfile.mkstemp in `directory` under the swept prefix."""
    import tempfile
    return tempfile.mkstemp(prefix=WRITE_TMP_PREFIX, suffix=suffix,
                            dir=os.fspath(directory))


def backup_folder_error(doc_path: str | os.PathLike,
                        exc: BaseException | None = None) -> XlMcpError:
    """The refusal when the backup folder beside a workbook cannot be used
    It used to surface as raw OS text ("[WinError 183] ...")
    after the save had already written its payload, which stayed behind."""
    doc = Path(doc_path)
    root = doc.resolve().parent / BACKUP_DIR_NAME
    if root.exists() and not root.is_dir():
        why = (f"a file named {BACKUP_DIR_NAME} is in the way, so the folder "
               "cannot be made")
        step = "Move or rename that file, then retry."
    else:
        detail = getattr(exc, "strerror", None) or "it could not be created"
        why = f"that folder could not be created or written ({detail})"
        step = "Check that the folder allows writing, then retry."
    err = XlMcpError(
        f"{doc.name} was not changed. Every edit first keeps a backup in the "
        f"{BACKUP_DIR_NAME} folder beside the workbook, and {why}. {step}")
    err.code = "CONFLICT"
    return err


def check_backup_folder(doc_path: str | os.PathLike) -> None:
    """Refuse before any work when the backup folder's place is taken by a
    file. Other failures surface when the folder is made."""
    root = Path(doc_path).resolve().parent / BACKUP_DIR_NAME
    if root.exists() and not root.is_dir():
        raise backup_folder_error(doc_path)


def backup_root(doc_path: str | os.PathLike) -> Path:
    """The .ks4xl-backups folder next to a workbook (not created)."""
    return Path(doc_path).resolve().parent / BACKUP_DIR_NAME


def slot_dir(doc_path: str | os.PathLike, *, create: bool = False) -> Path:
    """This workbook's slot folder inside .ks4xl-backups (created on demand)."""
    p = Path(doc_path).resolve()
    root = p.parent / BACKUP_DIR_NAME
    folder = _folder_name_for(p.name)
    d = root / folder
    if create:
        d.mkdir(parents=True, exist_ok=True)
        _hide(root)  # dot-prefix everywhere + real hidden bit on Windows
        if folder != p.name:
            crumb = d / _SOURCE_NAME_FILE
            if not crumb.exists():
                try:
                    crumb.write_text(p.name, encoding="utf-8")
                except OSError:
                    pass  # breadcrumb is best-effort
    return d


def source_doc_for(folder: Path) -> Path:
    """Map a slot folder back to its source workbook path (may not exist)."""
    name = folder.name
    crumb = folder / _SOURCE_NAME_FILE
    if crumb.exists():
        try:
            name = crumb.read_text(encoding="utf-8").strip() or folder.name
        except OSError:
            pass
    return folder.parent.parent / name


def _hide(path: Path) -> None:
    """Set FILE_ATTRIBUTE_HIDDEN on Windows. Non-fatal on any failure."""
    if sys.platform != "win32":
        return
    try:
        import ctypes

        target = str(path)
        if len(target) > 240 and not target.startswith("\\\\?\\"):
            target = "\\\\?\\" + target
        FILE_ATTRIBUTE_HIDDEN = 0x2
        attrs = ctypes.windll.kernel32.GetFileAttributesW(target)
        if attrs != -1 and not (attrs & FILE_ATTRIBUTE_HIDDEN):
            ctypes.windll.kernel32.SetFileAttributesW(
                target, attrs | FILE_ATTRIBUTE_HIDDEN
            )
    except Exception:
        pass  # hidden bit is cosmetic; the dot-prefix stands on its own


# ------------------------------------------------------- retry / placement


def _is_transient(exc: OSError) -> bool:
    winerr = getattr(exc, "winerror", None)
    return isinstance(exc, PermissionError) or winerr in _TRANSIENT_WINERRORS


def _with_retry(fn, *args):
    """Run an os-level file op, retrying briefly on AV sharing violations."""
    for delay in _RETRY_DELAYS:
        try:
            return fn(*args)
        except OSError as exc:
            if not _is_transient(exc):
                raise
            time.sleep(delay)
    return fn(*args)  # final attempt, exceptions propagate


def replace_with_retry(src: str | os.PathLike, dst: str | os.PathLike) -> None:
    """os.replace with transient-error backoff (slot files and other files
    this server owns; a user's workbook is put in place with promote)."""
    _with_retry(os.replace, os.fspath(src), os.fspath(dst))


# ----------------------------------------------------------- promote


_REPLACE_FILE = None

#: ReplaceFileW errors.
_ERROR_UNABLE_TO_REMOVE_REPLACED = 1175   # nothing changed (the file is held)
_ERROR_UNABLE_TO_MOVE_REPLACEMENT = 1176  # with a backup name: nothing changed
_ERROR_UNABLE_TO_MOVE_REPLACEMENT_2 = 1177  # the original is at the backup name
#: File systems without ReplaceFile (some network and WebDAV drives) answer
#: one of these. The promote then renames the new version over the workbook
#: and carries the identity over by hand first (_carry_identity).
_REPLACE_UNSUPPORTED = {1, 50, 87, 120}

#: ReplaceFileW flag: go on when the workbook's permissions cannot be written
#: onto the new file. Used only after the call without it was denied.
_REPLACEFILE_IGNORE_ACL_ERRORS = 0x4

#: Attributes the fallback carries over: hidden, system, archive and not
#: content indexed (ReadOnly would stop the rename; the cloud states are the
#: sync engine's).
_CARRIED_ATTRIBUTES = 0x2 | 0x4 | 0x20 | 0x2000


def _named_streams(path: str) -> list[str]:
    """The names of a file's named data streams ("Zone.Identifier" and any
    other), without the main one. Windows only; empty when the file system
    has no streams or they cannot be listed."""
    import ctypes
    from ctypes import wintypes

    class _StreamData(ctypes.Structure):
        _fields_ = [("StreamSize", ctypes.c_longlong),
                    ("cStreamName", ctypes.c_wchar * (260 + 36))]

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.FindFirstStreamW.restype = wintypes.HANDLE
    k32.FindFirstStreamW.argtypes = [wintypes.LPCWSTR, ctypes.c_int,
                                     ctypes.c_void_p, wintypes.DWORD]
    k32.FindNextStreamW.restype = wintypes.BOOL
    k32.FindNextStreamW.argtypes = [wintypes.HANDLE, ctypes.c_void_p]
    k32.FindClose.argtypes = [wintypes.HANDLE]
    data = _StreamData()
    handle = k32.FindFirstStreamW(path, 0, ctypes.byref(data), 0)
    if not handle or handle == ctypes.c_void_p(-1).value:
        return []
    names = []
    try:
        while True:
            raw = data.cStreamName  # ":name:$DATA", or "::$DATA" for the main one
            if raw.startswith(":") and raw.endswith(":$DATA") and raw != "::$DATA":
                names.append(raw[1:-len(":$DATA")])
            if not k32.FindNextStreamW(handle, ctypes.byref(data)):
                break
    finally:
        k32.FindClose(handle)
    return names


def _carry_identity(dst: str, src: str) -> None:
    """Give the new version `src` what ReplaceFileW would have carried over
    from the workbook `dst`, before `src` is renamed over it: every
    named stream (among them the Mark of the Web, the
    Zone.Identifier stream: a workbook from the internet keeps opening in
    Protected View, and Office's block on macros from the internet still
    applies), the creation time, the attributes above and the permissions
    (the DACL, with whether it inherits from the folder).
    The permissions go last, so nothing they forbid is needed afterwards.
    Best effort, step by step: a file system without one of these has
    nothing of it to carry, and the save goes on. Windows only."""
    import ctypes
    from ctypes import wintypes

    from .access import long_path

    ldst, lsrc = long_path(dst), long_path(src)
    try:
        streams = _named_streams(ldst)
    except Exception:  # noqa: BLE001
        streams = ["Zone.Identifier"]
    for stream in streams:
        try:
            with open(f"{ldst}:{stream}", "rb") as fh_in, \
                    open(f"{lsrc}:{stream}", "wb") as fh_out:
                shutil.copyfileobj(fh_in, fh_out, 1 << 20)
        except OSError:
            pass  # gone meanwhile, or no streams on this file system
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    try:
        k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                    ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                                    wintypes.HANDLE]
        k32.CreateFileW.restype = wintypes.HANDLE
        k32.GetFileTime.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_void_p,
                                    ctypes.c_void_p]
        k32.SetFileTime.argtypes = k32.GetFileTime.argtypes
        k32.CloseHandle.argtypes = [wintypes.HANDLE]
        invalid = ctypes.c_void_p(-1).value
        born = wintypes.FILETIME()
        h = k32.CreateFileW(ldst, 0x80, 0x7, None, 3, 0, None)  # READ_ATTRIBUTES
        if h and h != invalid:
            try:
                got = k32.GetFileTime(h, ctypes.byref(born), None, None)
            finally:
                k32.CloseHandle(h)
            if got:
                h = k32.CreateFileW(lsrc, 0x100, 0x7, None, 3, 0, None)  # WRITE_ATTRIBUTES
                if h and h != invalid:
                    try:
                        k32.SetFileTime(h, ctypes.byref(born), None, None)
                    finally:
                        k32.CloseHandle(h)
        k32.GetFileAttributesW.argtypes = [wintypes.LPCWSTR]
        k32.GetFileAttributesW.restype = wintypes.DWORD
        k32.SetFileAttributesW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD]
        a_dst, a_src = k32.GetFileAttributesW(ldst), k32.GetFileAttributesW(lsrc)
        if a_dst != 0xFFFFFFFF and a_src != 0xFFFFFFFF:
            want = (a_src & ~_CARRIED_ATTRIBUTES) | (a_dst & _CARRIED_ATTRIBUTES)
            if want != a_src:
                k32.SetFileAttributesW(lsrc, want)
    except Exception:  # noqa: BLE001 (best effort)
        pass
    try:
        adv = ctypes.WinDLL("advapi32", use_last_error=True)
        adv.GetNamedSecurityInfoW.argtypes = [
            wintypes.LPCWSTR, ctypes.c_int, wintypes.DWORD, ctypes.c_void_p,
            ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p), ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_void_p)]
        adv.GetNamedSecurityInfoW.restype = wintypes.DWORD
        adv.SetNamedSecurityInfoW.argtypes = [
            wintypes.LPWSTR, ctypes.c_int, wintypes.DWORD, ctypes.c_void_p,
            ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
        adv.SetNamedSecurityInfoW.restype = wintypes.DWORD
        adv.GetSecurityDescriptorControl.argtypes = [
            ctypes.c_void_p, ctypes.POINTER(wintypes.WORD), ctypes.POINTER(wintypes.DWORD)]
        k32.LocalFree.argtypes = [ctypes.c_void_p]
        dacl, sd = ctypes.c_void_p(), ctypes.c_void_p()
        # SE_FILE_OBJECT, DACL_SECURITY_INFORMATION
        if adv.GetNamedSecurityInfoW(ldst, 1, 0x4, None, None, ctypes.byref(dacl),
                                     None, ctypes.byref(sd)) == 0:
            try:
                control, rev = wintypes.WORD(), wintypes.DWORD()
                protected = (adv.GetSecurityDescriptorControl(
                    sd, ctypes.byref(control), ctypes.byref(rev))
                    and control.value & 0x1000)  # SE_DACL_PROTECTED
                # PROTECTED_ / UNPROTECTED_DACL_SECURITY_INFORMATION: whether
                # the folder's inheritable entries apply, as on the workbook
                how = 0x4 | (0x80000000 if protected else 0x20000000)
                adv.SetNamedSecurityInfoW(lsrc, 1, how, None, None, dacl, None)
            finally:
                k32.LocalFree(sd)
    except Exception:  # noqa: BLE001 (best effort)
        pass


def _replace_file():
    """(ReplaceFileW, MoveFileExW, last error, error text), loaded once."""
    global _REPLACE_FILE
    if _REPLACE_FILE is None:
        import ctypes
        from ctypes import wintypes
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        fn = k32.ReplaceFileW
        fn.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR,
                       wintypes.DWORD, ctypes.c_void_p, ctypes.c_void_p]
        fn.restype = wintypes.BOOL
        move = k32.MoveFileExW
        move.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD]
        move.restype = wintypes.BOOL
        _REPLACE_FILE = (fn, move, ctypes.get_last_error,
                         lambda e: ctypes.FormatError(e).strip())
    return _REPLACE_FILE


def _dacl_entries(path: str) -> list[tuple[str, ...]] | None:
    """The file's permission entries, each as its SDDL fields without the
    inherited flag, sorted; None when they cannot be read. Two files grant
    the same access exactly when these match: a workbook whose entries were
    its own copies of its folder's (Python's private temp folders make files
    that way) and comes back with the folder's inherited ones has lost
    nothing, and is not warned about."""
    text = _dacl_text(path)
    if text is None:
        return None
    import re
    out = []
    for ace in re.findall(r"\(([^()]*)\)", text.split("S:", 1)[0]):
        fields = ace.split(";")
        if len(fields) >= 2:
            flags = fields[1]
            fields[1] = "".join(flags[i:i + 2] for i in range(0, len(flags), 2)
                                if flags[i:i + 2] != "ID")
        out.append(tuple(fields))
    return sorted(out)


def _dacl_text(path: str) -> str | None:
    """The file's permissions (its DACL) as SDDL text, or None when they
    cannot be read. Inherited entries carry the ID flag."""
    import ctypes
    from ctypes import wintypes

    from .access import long_path

    adv = ctypes.WinDLL("advapi32", use_last_error=True)
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    get = adv.GetFileSecurityW
    get.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, ctypes.c_void_p,
                    wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
    get.restype = wintypes.BOOL
    conv = adv.ConvertSecurityDescriptorToStringSecurityDescriptorW
    conv.argtypes = [ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                     ctypes.POINTER(wintypes.LPWSTR), ctypes.c_void_p]
    conv.restype = wintypes.BOOL
    k32.LocalFree.argtypes = [ctypes.c_void_p]
    dacl_information = 0x4
    target = long_path(path)
    need = wintypes.DWORD(0)
    get(target, dacl_information, None, 0, ctypes.byref(need))
    if not need.value:
        return None
    buf = ctypes.create_string_buffer(need.value)
    if not get(target, dacl_information, buf, need, ctypes.byref(need)):
        return None
    text = wintypes.LPWSTR()
    if not conv(buf, 1, dacl_information, ctypes.byref(text), None):
        return None
    try:
        return text.value
    finally:
        k32.LocalFree(ctypes.cast(text, ctypes.c_void_p))


def promote(src: str | os.PathLike, dst: str | os.PathLike, *,
            ticket: "RotationTicket | None" = None) -> list[str]:
    """Put the new version `src` in place of the workbook `dst`, keeping what
    Windows keeps on the file itself rather than in its bytes.
    Returns the warnings the caller passes on (usually none).

    In a folder where people may write files but not change their
    permissions (OWNER RIGHTS without WRITE_DAC, a common file-share
    hardening) ReplaceFileW with flags 0 is denied (error 5): it cannot
    write the workbook's permissions onto the new file. It is then made once
    more, at once, ignoring that error (REPLACEFILE_IGNORE_ACL_ERRORS), which
    still carries everything else; a workbook whose own permissions did not
    come across is named in a warning.

    os.replace gave the path a NEW file, so one ordinary edit dropped the
    workbook's explicit permissions (a DENY entry and an explicit grant
    vanished, and the file fell back to whatever its folder hands down), its
    Mark of the Web (the Zone.Identifier stream, so Excel stopped opening a
    downloaded file in Protected View and Office's block on macros from the
    internet no longer applied), its attributes (OneDrive's pinned and
    offline states, not content indexed) and its creation time. Excel's own
    save keeps all of them, and so does ReplaceFileW, which this uses.
    MEASURED: explicit grant and deny entries, Zone.Identifier,
    NOT_CONTENT_INDEXED, OFFLINE, PINNED, UNPINNED and the creation time all
    survive; a reader that shares delete does not stop it; one that does not
    share delete fails it with a sharing violation and changes nothing.

    The original goes aside first, beside it, under a name that says what
    it holds (core/swapwait.py aside_name), so that a failure can put it
    back, and is removed after; the backup slots are hard links of the
    original and keep its bytes. The path is missing for the instant
    between the two moves, which every route in this server waits out
    (core/swapwait.py). Each swap is written down first in a small per-user
    journal, so a server stopped between the two moves leaves a record: the
    workbook's next write lock, or the next start of a server, puts it back
    (put_back_if_aside). Nothing removes an aside file by age.

    ``ticket`` is the backup rotation staged for this swap
    (prepare_rotation). It lands on the slots only after the swap went
    through, and is dropped when it did not, so a refused save, restore or
    overwrite never costs the undo of the last change. The
    journal entry records it, so a server stopped after the swap still ends
    with the right undo at the next write lock.

    Where ReplaceFileW is not available (not Windows, a new file, a file
    system without it), os.replace as before; on Windows the workbook's
    identity is first carried over to the new version by hand
    (_carry_identity). Transient sharing violations (an antivirus scan)
    are retried."""
    src_s, dst_s = os.fspath(src), os.fspath(dst)
    entry = aside = None
    try:
        if sys.platform != "win32" or not os.path.exists(dst_s):
            replace_with_retry(src_s, dst_s)
            warnings: list[str] = []
        else:
            fn, move, last_error, error_text = _replace_file()
            folder = os.path.dirname(os.path.abspath(dst_s)) or "."
            aside = os.path.join(folder, aside_name(
                os.path.basename(dst_s), uuid.uuid4().hex[:12]))
            entry = _journal_open(dst_s, aside, src_s, ticket)
            warnings = _swap(fn, move, last_error, error_text,
                             src_s, dst_s, aside)
    except BaseException:
        if ticket is not None:
            ticket.abort()
        # Kept while the old file is still aside (1177 not put back), so
        # the workbook's next write lock puts it back.
        _journal_close(entry, keep=_still_aside(aside))
        raise
    if ticket is not None:
        warnings = [*warnings, *ticket.land()]
    # Kept while the old file is still aside (a copy another program holds),
    # so the workbook's next write lock tidies it.
    _journal_close(entry, keep=_still_aside(aside))
    return warnings


class WorkbookLeftAside(OSError):
    """ReplaceFileW moved the workbook aside, could not move the new version
    in (1177), and the workbook could not be moved back: it is at
    ``aside``, and the next change or the next start of the server puts it
    back (put_back_if_aside)."""

    def __init__(self, workbook: str, aside: str, text: str, winerror: int):
        super().__init__(0, text, workbook, winerror)
        self.aside = aside


def _still_aside(aside: str | None) -> bool:
    from .access import long_path
    return aside is not None and os.path.lexists(long_path(aside))


def _swap(fn, move, last_error, error_text, src_s: str, dst_s: str,
          backup: str) -> list[str]:
    """ReplaceFileW and what its errors mean; see promote."""
    from .access import long_path
    lsrc, ldst, lbak = long_path(src_s), long_path(dst_s), long_path(backup)
    err = 0
    flags = 0
    permissions_before = None
    for delay in _RETRY_DELAYS + (None,):
        if fn(ldst, lsrc, lbak, flags, None, None):
            try:
                _with_retry(os.remove, lbak)
            except OSError:
                pass  # a program still holds the old copy; tidied at the next write lock
            if flags and permissions_before is not None:
                try:
                    kept = _dacl_entries(dst_s) == permissions_before
                except Exception:  # noqa: BLE001
                    kept = True  # unreadable now: nothing to report
                if not kept:
                    return [
                        f"{os.path.basename(dst_s)} was saved, but its own "
                        "permissions could not be kept: this folder does not "
                        "let file owners change permissions, so the workbook "
                        "now has the folder's permissions."]
            return []
        err = last_error()
        if err == 5 and not flags:
            # Denied with flags 0. Where file owners may not change
            # permissions, ReplaceFileW cannot write the merged permissions
            # onto the new file; once more at once, ignoring that error.
            try:
                permissions_before = _dacl_entries(dst_s)
            except Exception:  # noqa: BLE001
                permissions_before = None
            flags = _REPLACEFILE_IGNORE_ACL_ERRORS
            continue
        if err == _ERROR_UNABLE_TO_MOVE_REPLACEMENT_2:
            # The original was renamed to the backup name and the new
            # version could not take its place: put the original back.
            if not move(lbak, ldst, 0x1 | 0x8):  # REPLACE_EXISTING | WRITE_THROUGH
                raise WorkbookLeftAside(
                    dst_s, backup,
                    f"{error_text(err)} The workbook as it was is at {backup}; "
                    "the next change to it through this server, or the next "
                    "start of the server, puts it back, or rename it back to "
                    f"{os.path.basename(dst_s)}.", err)
            break
        if err in _REPLACE_UNSUPPORTED:
            _carry_identity(dst_s, src_s)
            replace_with_retry(src_s, dst_s)
            return []
        if err not in _TRANSIENT_WINERRORS | {
                _ERROR_UNABLE_TO_REMOVE_REPLACED,
                _ERROR_UNABLE_TO_MOVE_REPLACEMENT} or delay is None:
            break
        time.sleep(delay)
    if err in _TRANSIENT_WINERRORS | {_ERROR_UNABLE_TO_REMOVE_REPLACED}:
        # 32 sharing violation, 5 denied (with permission errors already
        # ignored), 1175 the file could not be moved aside: another program
        # holds it, or the folder refuses. The workbook is exactly as it was.
        raise PermissionError(13, error_text(err), dst_s, err)
    raise OSError(0, error_text(err), dst_s, err)


# ---------------------------------------------- a save that did not finish
#
# ReplaceFileW moves the workbook to its aside name and then moves the new
# version in. A server stopped between those two moves (a crash, an app
# closed while it saves, a client that ends its servers) leaves the
# workbook's path empty and the workbook under the aside name, whose words
# say what it holds (MEASURED: 2 of 81 aimed kills with a backup, 10 of 31
# without). Each swap is written down first in a small
# journal in the server's own folder, so the next start of a server finds
# it; the workbook's next write lock puts it back from the files beside it
# even without the journal. Reads never move it: they say where it is
# (core/swapwait.py). Nothing here removes an aside file while the workbook
# is missing, and nothing anywhere removes one by age (the stale-temp sweep
# in core/package.py never matches an aside name).

_JOURNAL_ENV = "KS4XL_SAVE_JOURNAL_DIR"

#: Journal entries of swaps running in this process right now.
_IN_FLIGHT: set[str] = set()
_IN_FLIGHT_GUARD = threading.Lock()

#: MoveFileExW: on disk before it returns. Without REPLACE_EXISTING it never
#: moves the old workbook over anything that is at the workbook's path.
_MOVEFILE_WRITE_THROUGH = 0x8


def _journal_dir() -> Path:
    override = os.environ.get(_JOURNAL_ENV)
    if override:
        return Path(override)
    import tempfile
    base = (os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_STATE_HOME")
            or tempfile.gettempdir())
    return Path(base) / "xlsx-mcp" / "saves-in-progress"


def _journal_open(workbook: str, aside: str, new: str,
                  ticket: "RotationTicket | None" = None) -> str | None:
    """Write a swap down before it starts. Best effort: without the entry
    the workbook is still put back from the files beside it at its next
    write lock. The entry appears whole (written under a temporary name,
    then renamed), so a reader never takes a half-written one for a
    finished swap."""
    try:
        folder = _journal_dir()
        folder.mkdir(parents=True, exist_ok=True)
        name = uuid.uuid4().hex
        entry = str(folder / f"{name}.json")
        staging = folder / f"{name}.writing"
        staging.write_text(json.dumps({
            "workbook": os.path.abspath(workbook),
            "aside": os.path.abspath(aside),
            "new": os.path.abspath(new),
            "staged": ticket.staged() if ticket is not None else [],
            "pid": os.getpid(), "token": _OWNER_TOKEN,
            "pid_created": _OWNER_CREATED,
            "host": _local_host(), "time": time.time(),
        }), encoding="utf-8")
        os.replace(staging, entry)
    except Exception:  # noqa: BLE001
        return None
    with _IN_FLIGHT_GUARD:
        _IN_FLIGHT.add(entry)
    return entry


def _journal_close(entry: str | None, *, keep: bool = False) -> None:
    if entry is None:
        return
    with _IN_FLIGHT_GUARD:
        _IN_FLIGHT.discard(entry)
    if not keep:
        try:
            os.remove(entry)
        except OSError:
            pass


def _journal_entries() -> list[tuple[str, dict]]:
    folder = _journal_dir()
    try:
        names = sorted(n for n in os.listdir(folder) if n.endswith(".json"))
    except OSError:
        return []
    found = []
    for name in names:
        entry = str(folder / name)
        try:
            with open(entry, encoding="utf-8") as fh:
                record = json.loads(fh.read())
        except (OSError, ValueError):
            record = None
        found.append((entry, record if isinstance(record, dict) else {}))
    return found


def _owner_gone(record: dict) -> bool:
    """The server that wrote a journal entry is no longer running: its PID
    is free, or it now names another process (its creation time differs)."""
    pid = record.get("pid", -1)
    if not isinstance(pid, int) or not _pid_alive(pid):
        return True
    if pid == os.getpid():
        return record.get("token") != _OWNER_TOKEN
    born = record.get("pid_created")
    now_born = _process_create_time(pid)
    return born is not None and now_born is not None and born != now_born


def _settled(entry: str, record: dict) -> bool:
    """A journal entry whose swap is not running: its server is gone, or it
    is this server's and not running now."""
    with _IN_FLIGHT_GUARD:
        if entry in _IN_FLIGHT:
            return False
    if _foreign_host(record):
        return False
    return _is_ours(record) or _owner_gone(record)


def protected_new_versions() -> set[str]:
    """The new versions that swaps written down in the journal had not
    finished moving in (normcase absolute paths). While a workbook is aside,
    this is the edit that was being saved; the stale-temp sweep leaves it to
    put_back_if_aside, which removes it only once the workbook is back."""
    out = set()
    for _entry, record in _journal_entries():
        new = record.get("new")
        if isinstance(new, str):
            out.add(os.path.normcase(os.path.abspath(new)))
    return out


def _remove_quietly(path) -> None:
    from .access import long_path
    try:
        os.remove(long_path(os.fspath(path)))
    except OSError:
        pass


def _own_temp_of(workbook: str, new: object) -> bool:
    """Whether a journal entry's ``new`` is a temporary file this server's
    saves of ``workbook`` write, the only thing recovery may remove: in the
    workbook's folder under the write-temp prefix (the file tier's and
    Excel's copy), or in its backup folder named .restore-*.tmp.xlsx (a
    restore). Never the workbook itself, under its name or another one (a
    hard link). A damaged or foreign entry naming any other file is never
    acted on."""
    from .access import long_path
    if not isinstance(new, str) or not new:
        return False
    try:
        wb = os.path.abspath(workbook)
        nw = os.path.abspath(new)
    except (TypeError, ValueError):
        return False
    if os.path.normcase(nw) == os.path.normcase(wb):
        return False
    folder, name = os.path.split(nw)
    in_folder = os.path.normcase(folder) == os.path.normcase(os.path.dirname(wb))
    in_slots = os.path.normcase(folder) == os.path.normcase(str(slot_dir(wb)))
    if not ((in_folder and name.startswith(WRITE_TMP_PREFIX))
            or (in_slots and name.startswith(".restore-") and name.endswith(".tmp.xlsx"))):
        return False
    try:
        if os.path.lexists(long_path(wb)) and os.path.samefile(long_path(nw), long_path(wb)):
            return False
    except (OSError, ValueError):
        pass
    return True


def _new_still_there(workbook: str, record: dict) -> bool:
    """The entry's new version is still beside the workbook: its swap did not
    move it in (it never started, or it stopped between the two moves)."""
    from .access import long_path
    new = record.get("new")
    return _own_temp_of(workbook, new) and os.path.lexists(long_path(new))


def _slot_identity(slot: str | os.PathLike) -> dict | None:
    """Which file a backup slot is right now (file id, size and time), or None
    when there is none. A staged rotation records it, so its landing can tell
    whether another save has rotated the slot since."""
    from .access import long_path
    try:
        st = os.stat(long_path(os.fspath(slot)))
    except OSError:
        return None
    return {"ino": int(st.st_ino), "size": int(st.st_size),
            "mtime_ns": int(st.st_mtime_ns)}


#: A staged move from an entry that recorded no slot identity (written by
#: an earlier version of this server).
_UNRECORDED = object()


def _slot_unchanged(slot: str, recorded: object, staged: str) -> bool:
    """Whether a staged rotation may still land on ``slot``: the slot is the
    very file it was when the rotation was staged. For an entry that did not
    record that, only when the slot is missing or older than the staged
    copy, so a staged copy never lands over a newer version that a save the
    journal did not see (an older build, another journal) put there."""
    from .access import long_path
    if recorded is not _UNRECORDED:
        return _slot_identity(slot) == recorded
    try:
        slot_time = os.stat(long_path(slot)).st_mtime_ns
    except OSError:
        return True
    try:
        return os.stat(long_path(staged)).st_mtime_ns >= slot_time
    except OSError:
        return False


def _slot_holds(workbook: str, candidate: str) -> bool:
    """Whether one of the workbook's backup slots is ``candidate`` itself (a
    hard link) or holds exactly its bytes."""
    import hashlib
    from .access import long_path
    folder = slot_dir(workbook)
    slots = [str(folder / PREV_SLOT), str(folder / ANCHOR_SLOT)]
    for slot in slots:
        try:
            if os.path.samefile(long_path(slot), long_path(candidate)):
                return True
        except (OSError, ValueError):
            continue

    def digest(p):
        h = hashlib.sha256()
        with open(long_path(p), "rb") as fh:
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                h.update(chunk)
        return h.digest()
    try:
        size = os.path.getsize(long_path(candidate))
        want = None
        for slot in slots:
            if os.path.lexists(long_path(slot)) and os.path.getsize(long_path(slot)) == size:
                want = want or digest(candidate)
                if digest(slot) == want:
                    return True
    except OSError:
        return False
    return False


#: What an old version left aside by a save is renamed to when it is no
#: longer the workbook's to put back (its swap finished, or the path holds
#: newer content now): beside the workbook, visible, and never over anything.
KEPT_SUFFIX = " (kept from a save that did not finish)"


def _retire_aside(workbook: str, aside: str) -> str | None:
    """An old version a save left aside that is no longer this workbook's to
    put back. Removed when a backup slot holds it already; otherwise renamed
    to a visible name beside the workbook (KEPT_SUFFIX), never over anything,
    so that no later write lock or create takes it for the workbook. Returns
    the new name, or None when it was removed or could not be moved (a
    program holds it: it stays, and the next write lock tries again)."""
    from .access import long_path
    if not os.path.lexists(long_path(aside)):
        return None
    if _slot_holds(workbook, aside):
        _remove_quietly(aside)
        return None
    folder = os.path.dirname(os.path.abspath(workbook))
    stem, ext = os.path.splitext(os.path.basename(workbook))
    for n in range(1, 100):
        name = f"{stem}{KEPT_SUFFIX}{'' if n == 1 else f' {n}'}{ext}"
        target = os.path.join(folder, name)
        if os.path.lexists(long_path(target)):
            continue
        try:
            os.rename(long_path(aside), long_path(target))  # never replaces a file
        except FileExistsError:
            continue
        except OSError:
            return None
        try:
            _clear_hidden(target)
        except OSError:
            pass
        return target
    return None


def _clear_hidden(path: str) -> None:
    """Make a kept file visible (an aside inherits the workbook's
    attributes, which may include hidden)."""
    if sys.platform != "win32":
        return
    import ctypes
    from .access import long_path
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    attrs = k32.GetFileAttributesW(long_path(path))
    if attrs != 0xFFFFFFFF and attrs & 0x2:
        k32.SetFileAttributesW(long_path(path), attrs & ~0x2)


def _move_back(aside: str, workbook: str) -> bool:
    """Move the old workbook back to its path, never over anything there."""
    if sys.platform != "win32":
        return False  # promote leaves nothing aside off Windows
    from .access import long_path
    try:
        _fn, move, _last_error, _error_text = _replace_file()
    except (OSError, AttributeError):
        return False
    return bool(move(long_path(aside), long_path(workbook),
                     _MOVEFILE_WRITE_THROUGH))


def _staged_moves(workbook: str, record: dict) -> list[tuple[str, str, object]]:
    """The backup rotation a journal entry staged: (temp, slot, the slot's
    identity when it was staged) triples, kept only when both files sit in
    this workbook's own backup folder. An entry from an earlier version
    records pairs; their identity is _UNRECORDED."""
    folder = os.path.normcase(str(slot_dir(workbook)))
    moves = []
    for item in record.get("staged") or []:
        if not (isinstance(item, list) and len(item) in (2, 3)
                and all(isinstance(p, str) for p in item[:2])):
            continue
        tmp, slot = item[0], item[1]
        recorded = item[2] if len(item) == 3 else _UNRECORDED
        if recorded is not _UNRECORDED and not (recorded is None or isinstance(recorded, dict)):
            continue
        if (os.path.normcase(os.path.dirname(tmp)) == folder
                and os.path.basename(tmp).startswith(".slot-")
                and os.path.normcase(os.path.dirname(slot)) == folder
                and os.path.basename(slot) in SLOT_POLICY):
            moves.append((tmp, slot, recorded))
    return moves


def _land_staged(workbook: str, record: dict) -> None:
    """Land a stopped save's staged rotation, each move only while its slot
    is still the file it was when the rotation was staged: when a save the
    journal did not see has rotated the slot since, the staged copy is older
    than what the slot holds now and is dropped (it would replace the only
    copy of that newer version)."""
    from .access import long_path
    for tmp, slot, recorded in _staged_moves(workbook, record):
        if not os.path.lexists(long_path(tmp)):
            continue
        if not _slot_unchanged(slot, recorded, tmp):
            _remove_quietly(tmp)
            continue
        try:
            os.replace(long_path(tmp), long_path(slot))
        except OSError:
            _remove_quietly(tmp)


def _drop_staged(workbook: str, record: dict) -> None:
    for tmp, _slot, _recorded in _staged_moves(workbook, record):
        _remove_quietly(tmp)


def _settle_present(workbook: str, record: dict) -> None:
    """A journal entry for a workbook that is in place: its swap never
    started, it stopped and the path has been given newer content since, or
    it ran to its end and only the clean-up was cut short."""
    from .access import long_path
    new, aside = record.get("new"), record.get("aside")
    aside_there = isinstance(aside, str) and os.path.lexists(long_path(aside))
    if _new_still_there(workbook, record):
        _remove_quietly(new)  # never swapped in: what is at the path stays
        _drop_staged(workbook, record)
        if aside_there:
            # The swap had moved the workbook aside and the path was given
            # other content since (the user's newer file): the old version
            # is no longer this workbook's to put back.
            _retire_aside(workbook, aside)
        return
    _settle_finished(workbook, record)


def _settle_finished(workbook: str, record: dict) -> None:
    """A journal entry whose swap ran to its end (its new version moved in)
    and whose clean-up was cut short: the backups take the version from
    before it, as they would have had the server not been stopped, and the
    old version still aside is no longer the workbook's to put back. It is
    removed when a slot holds it, else kept under a visible name, even when
    the workbook has been renamed or deleted since (nothing recreates it)."""
    from .access import long_path
    _land_staged(workbook, record)
    aside = record.get("aside")
    if isinstance(aside, str) and os.path.lexists(long_path(aside)):
        _retire_aside(workbook, aside)


def _staged_links(workbook: str) -> list[str]:
    from .access import long_path
    folder = slot_dir(workbook)
    try:
        names = os.listdir(long_path(str(folder)))
    except OSError:
        return []
    return [str(folder / n) for n in names
            if n.startswith(".slot-") and n.endswith(".tmp")]


def _legacy_aside_of(workbook: str) -> str | None:
    """An old file that an earlier version's save left aside, when it is
    provably this workbook: the interrupted save had
    staged its backup, and that staged backup in this workbook's own backup
    folder is the very same file. Never a guess."""
    from .access import long_path
    from .swapwait import legacy_asides
    staged = _staged_links(workbook)
    for candidate in legacy_asides(workbook) if staged else []:
        for link in staged:
            try:
                if os.path.samefile(long_path(candidate), long_path(link)):
                    return candidate
            except (OSError, ValueError):
                continue
    return None


def put_back_if_aside(doc_path: str | os.PathLike) -> bool:
    """Put back a workbook that a save which did not finish left aside, and
    tidy up after swaps that did finish. Runs with the workbook's write lock
    held (write_lock calls it at its outermost hold), so no save of this
    workbook is running, here or in another server. True when it moved the
    workbook back.

    Only the workbook's own aside files are ever moved, only onto an empty
    path, and none is removed while the workbook is missing. The new version
    a stopped save was moving in is removed only once the workbook is back:
    the call that made it never got an answer, so the workbook returns to
    the last version a call reported saved."""
    from .access import long_path
    from .swapwait import asides
    workbook = os.path.abspath(os.fspath(doc_path))
    key = canonical_key(workbook)
    ours = [(entry, record) for entry, record in _journal_entries()
            if isinstance(record.get("workbook"), str)
            and _same_workbook(record["workbook"], workbook, key)
            and _settled(entry, record)]
    if os.path.lexists(long_path(workbook)):
        for entry, record in ours:
            _settle_present(workbook, record)
            _journal_close(entry)
        return False
    # The workbook is missing. Only a swap that did NOT finish (its new
    # version is still beside the workbook) has the workbook aside; a
    # finished one's old version is not put back (the user may have renamed
    # or deleted the workbook since), it is settled.
    unfinished = [(e, r) for e, r in ours if _new_still_there(workbook, r)]
    for entry, record in ours:
        if (entry, record) not in unfinished:
            _settle_finished(workbook, record)
            _journal_close(entry)
    legacy = None
    if ours:
        waiting = sorted(
            ((r.get("time") if isinstance(r.get("time"), (int, float)) else 0.0, r["aside"])
             for _e, r in unfinished
             if isinstance(r.get("aside"), str) and os.path.lexists(long_path(r["aside"]))),
            reverse=True)
        if not waiting:
            for entry, _record in unfinished:
                _journal_close(entry)  # its old version was moved by hand
            return False
        chosen = waiting[0][1]  # the save that began last: the newest version
    else:
        # No entry at all (the journal could not be written): by name, as the
        # files beside the workbook say.
        found = asides(workbook)
        legacy = None if found else _legacy_aside_of(workbook)
        if not found and legacy is None:
            return False
        chosen = legacy if legacy is not None else found[0]
    if not _move_back(chosen, workbook):
        return False
    for entry, record in unfinished:
        if _own_temp_of(workbook, record.get("new")):
            _remove_quietly(record["new"])  # the version whose save did not finish
        _drop_staged(workbook, record)  # the workbook is as it was; so are the slots
        aside = record.get("aside")
        if isinstance(aside, str) and os.path.normcase(aside) != os.path.normcase(chosen):
            _retire_aside(workbook, aside)  # an older save's, not the workbook now
        _journal_close(entry)
    if legacy is not None:
        # The earlier version's staged backup links of this very file are
        # spare names of the workbook now that it is back.
        for link in _staged_links(workbook):
            try:
                if os.path.samefile(long_path(link), long_path(workbook)):
                    _remove_quietly(link)
            except (OSError, ValueError):
                continue
    return True


def put_back_before_write(path: str | os.PathLike) -> bool:
    """Before a writing call that names `path`: when the workbook is missing
    and a save that did not finish has it aside, take its write lock, whose
    outermost hold puts it back. A create or copy onto that name then never
    writes over it or strands it. True when the workbook is back. Never
    raises; the call itself answers whatever is wrong with the path."""
    from .access import long_path
    from .swapwait import asides, legacy_asides
    try:
        p = os.path.abspath(os.fspath(path))
        if os.path.lexists(long_path(p)) or not (asides(p) or legacy_asides(p)):
            return False
        with write_lock(p):
            pass
        return os.path.lexists(long_path(p))
    except Exception:  # noqa: BLE001
        return False


#: How long start-up recovery waits for one folder to answer.
UNREACHABLE_WAIT_SECONDS = 1.0

#: An entry whose folder has not been reachable for this long is dropped.
UNREACHABLE_DROP_SECONDS = 30 * 24 * 3600


def _share_of(folder: str) -> str:
    """What a folder's reachability goes with: its network share
    (\\\\host\\share) or its drive."""
    drive, _rest = os.path.splitdrive(os.path.abspath(folder))
    return os.path.normcase(drive)


#: Shares (or drives) that did not answer, and until when (monotonic) they
#: are taken as silent without asking again.
_SILENT: dict[str, float] = {}
SILENT_MEMORY_SECONDS = 300.0


def _share_answers(path: str, wait: float = 0.5) -> bool:
    """Whether the share (or drive) holding ``path`` answers, asked for at
    most ``wait`` seconds and remembered for a while when it does not."""
    share = _share_of(os.path.dirname(os.path.abspath(path)))
    until = _SILENT.get(share)
    if until is not None and until > time.monotonic():
        return False
    root = share + os.sep if share else os.path.dirname(os.path.abspath(path))
    if _folder_state(root, wait) == "silent":
        _SILENT[share] = time.monotonic() + SILENT_MEMORY_SECONDS
        return False
    return True


def _same_workbook(recorded: str, workbook: str, key: str) -> bool:
    """Whether a journal entry's workbook is ``workbook`` (its key given).
    The same spelling decides at once; another spelling is resolved, which
    for a path on another share waits on that share, so it is resolved only
    when the share answers (an entry for a share that is offline is never
    this workbook, whose own share answers)."""
    try:
        if os.path.normcase(os.path.abspath(recorded)) == os.path.normcase(workbook):
            return True
    except (TypeError, ValueError):
        return False
    if _share_of(os.path.dirname(os.path.abspath(recorded))) !=             _share_of(os.path.dirname(workbook)) and not _share_answers(recorded):
        return False
    return canonical_key(recorded) == key


def _folder_state(folder: str, wait: float) -> str:
    """"there", "gone" (it answered: no such folder) or "silent" (no answer
    within ``wait`` seconds). The check runs on its own thread, so a share
    that does not answer costs the wait, not the network's timeout; that
    thread ends by itself later."""
    import threading
    answer: list[bool] = []

    def probe() -> None:
        try:
            answer.append(os.path.isdir(folder))
        except Exception:  # noqa: BLE001
            answer.append(False)

    t = threading.Thread(target=probe, name="ks4xl-recovery-probe", daemon=True)
    t.start()
    t.join(wait)
    if not answer:
        return "silent"
    return "there" if answer[0] else "gone"


def recover_interrupted_saves(*, wait_seconds: float = 0.5,
                              budget_seconds: float = 5.0) -> list[str]:
    """Put back every workbook that a save which did not finish left aside,
    for the journal entries whose server is gone. Run once when a server
    starts (never in a read-only one). A workbook whose lock another server
    holds is left to that server: its next write lock does the same.
    Returns the workbooks put back."""
    from .access import long_path
    restored: list[str] = []
    seen: set[str] = set()
    unreachable: set[str] = set()
    stop = time.monotonic() + budget_seconds
    for entry, record in _journal_entries():
        if time.monotonic() > stop:
            break
        if not _settled(entry, record):
            continue
        workbook = record.get("workbook")
        if not isinstance(workbook, str):
            try:
                if time.time() - os.path.getmtime(entry) > 3600:
                    _journal_close(entry)  # unreadable, and long settled
            except OSError:
                pass
            continue
        folder = os.path.dirname(workbook)
        share = _share_of(folder)
        state = "silent" if share in unreachable else _folder_state(
            folder, min(UNREACHABLE_WAIT_SECONDS, max(0.0, stop - time.monotonic())))
        if state != "there":
            # Not reachable now (an offline share waits out the network's
            # timeout, about 21 s): skipped within the wait, and the rest of
            # a silent share with it, so a start is never held up by it. An
            # entry whose folder stays unreachable for weeks is dropped; its
            # workbook is still put back from the files beside it at its
            # next write lock.
            if state == "silent":
                unreachable.add(share)
                _SILENT[share] = time.monotonic() + SILENT_MEMORY_SECONDS
            when = record.get("time")
            if isinstance(when, (int, float)) and time.time() - when > UNREACHABLE_DROP_SECONDS:
                _journal_close(entry)
            continue
        key = canonical_key(workbook)  # resolves the path: only once it answers
        if key in seen:
            continue
        seen.add(key)
        missing = not os.path.lexists(long_path(workbook))
        try:
            with write_lock(workbook, wait_seconds=wait_seconds):
                pass  # the outermost hold puts the workbook back
        except Exception:  # noqa: BLE001
            continue
        if missing and os.path.lexists(long_path(workbook)):
            restored.append(workbook)
    return restored


def _link_or_copy(src: Path, dst: Path) -> None:
    """Hardlink src as dst; copy2 fallback for non-NTFS / cloud placeholders /
    anything the link call rejects. copy2 preserves mtime, which the anchor
    idle measurement relies on."""
    try:
        _with_retry(os.link, str(src), str(dst))
    except OSError:
        _with_retry(shutil.copy2, str(src), str(dst))


def _place_onto_slot(current: Path, slot_path: Path) -> None:
    """Capture `current`'s present content into a slot without the workbook
    ever leaving its own path: link (or copy) to a temp name in the slot
    folder, then atomically replace the slot."""
    tmp = slot_path.parent / f".slot-{uuid.uuid4().hex}.tmp"
    try:
        _link_or_copy(current, tmp)
        replace_with_retry(tmp, slot_path)
    except BaseException:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        raise


# ------------------------------------------------------------- slot rotation


def rotate_slots(doc_path: str | os.PathLike) -> dict:
    """Rotate backup slots for `doc_path` BEFORE its content is replaced.

    Called by the save path after the new payload is validated and written to
    a temp file, immediately before the final os.replace onto the target.
    Captures the CURRENT (pre-mutation) content:

    - anchor.xlsx: created if absent (first mutation ever / after purge), or
      rotated when the prev slot says the workbook has been idle 60+ minutes
      (new session).
    - prev.xlsx: rotated on every call.

    Never unlinks or renames the workbook itself.
    """
    doc = Path(doc_path).resolve()
    try:
        d = slot_dir(doc, create=True)
    except OSError as exc:
        raise backup_folder_error(doc, exc) from exc
    prev = d / PREV_SLOT
    anchor = d / ANCHOR_SLOT
    rotated = {"prev": True, "anchor": False}

    if not anchor.exists():
        _place_onto_slot(doc, anchor)
        rotated["anchor"] = True
    elif prev.exists():
        try:
            idle = time.time() - prev.stat().st_mtime
        except OSError:
            idle = 0.0
        if idle >= ANCHOR_IDLE_SECONDS:
            _place_onto_slot(doc, anchor)
            rotated["anchor"] = True

    _place_onto_slot(doc, prev)
    return rotated


class RotationTicket:
    """Two-phase slot rotation for the save path.

    rotate_slots() replaces the slots IMMEDIATELY, so a mutation that then
    failed at promote time (os.replace beaten by a lock landing between the
    write-probe and the promote) had already burned the prev slot: prev held
    a copy of the CURRENT content and the restore point for the previous
    successful mutation was gone, even though nothing was written.

    prepare() captures the current content into temp files inside the slot
    folder using the same hardlink-or-copy mechanism; commit() (called only
    AFTER a successful promote) replaces the slots; abort() unlinks the
    temps and leaves every slot exactly as it was. The hardlinked temps keep
    the pre-mutation inode alive across the promote, so commit still stores
    the pre-mutation bytes even though the path now holds the new content.
    """

    def __init__(self) -> None:
        #: (temp, slot, the slot's identity when it was staged)
        self._moves: list[tuple[Path, Path, dict | None]] = []
        self.rotated = {"prev": False, "anchor": False}

    def staged(self) -> list[list]:
        """The staged moves for the save journal: temp, slot, and which file
        the slot was when the rotation was staged, so a landing after a stop
        can tell whether another save has rotated the slot since."""
        return [[str(tmp), str(slot), ident] for tmp, slot, ident in self._moves]

    def land(self) -> list[str]:
        """commit(), after a swap that went through (promote calls it). A
        failure is a warning, because the workbook itself was saved."""
        try:
            self.commit()
        except OSError as exc:
            self.abort()
            return ["the save succeeded but the backup slot rotation failed "
                    f"({type(exc).__name__}: {exc}); the prev slot still holds "
                    "the state before the PREVIOUS mutation, not this one"]
        return []

    def commit(self) -> dict:
        for tmp, slot, _ident in self._moves:
            replace_with_retry(tmp, slot)
            name = "prev" if slot.name == PREV_SLOT else "anchor"
            self.rotated[name] = True
        self._moves = []
        return self.rotated

    def abort(self) -> None:
        for tmp, _slot, _ident in self._moves:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
        self._moves = []


def prepare_rotation(doc_path: str | os.PathLike, *,
                     anchor: bool = True) -> RotationTicket:
    """Stage a slot rotation for `doc_path` WITHOUT touching the slots.
    Same anchor/prev policy as rotate_slots (prev only with anchor=False, as
    a restore does); hand the ticket to promote, which lands it after the
    swap and drops it when the swap fails. commit() lands it directly,
    abort() discards it."""
    doc = Path(doc_path).resolve()
    try:
        d = slot_dir(doc, create=True)
    except OSError as exc:
        raise backup_folder_error(doc, exc) from exc
    prev = d / PREV_SLOT
    anchor_slot = d / ANCHOR_SLOT
    anchor_wanted = anchor
    ticket = RotationTicket()

    def _stage(slot_path: Path) -> None:
        tmp = slot_path.parent / f".slot-{uuid.uuid4().hex}.tmp"
        ident = _slot_identity(slot_path)
        try:
            _link_or_copy(doc, tmp)
        except BaseException:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
            raise
        ticket._moves.append((tmp, slot_path, ident))

    try:
        if not anchor_wanted:
            pass
        elif not anchor_slot.exists():
            _stage(anchor_slot)
        elif prev.exists():
            try:
                idle = time.time() - prev.stat().st_mtime
            except OSError:
                idle = 0.0
            if idle >= ANCHOR_IDLE_SECONDS:
                _stage(anchor_slot)
        _stage(prev)
    except BaseException:
        ticket.abort()
        raise
    return ticket


# ------------------------------------------------------------------ locking

_MUTEXES: dict[str, threading.RLock] = {}
_MUTEX_GUARD = threading.Lock()

#: Open write_lock holds per file in this process, nested ones included.
#: Read and changed only by the thread holding that file's mutex, so zero
#: means no hold in this process is open: a lockfile carrying our token is
#: then a leftover from a release that could not delete it, not a hold.
_DEPTH: dict[str, int] = {}

#: Deletes lockfiles a release could not, while this process lives.
_SWEEPER = lockrelease.Sweeper("ks4xl-write-lock-sweeper")

#: Threads that must never take a write lock, with the reason.
_THREAD_RULES = threading.local()


def forbid_write_lock_on_this_thread(why: str) -> None:
    """Mark the calling thread as one that must never take a write lock.

    The COM worker thread (com/session.py) runs each job for a caller thread
    that already holds the workbook's write lock (ops/comtier.py). The
    lock is re-entrant per THREAD, so a hold taken on the worker would wait
    for its own caller until the COM timeout reclaimed Excel. On a marked
    thread write_lock raises at once instead, so such a mistake fails the
    call (and every test that reaches it) rather than hanging it."""
    _THREAD_RULES.forbidden = str(why)


def _mutex_for(key: str) -> threading.RLock:
    with _MUTEX_GUARD:
        lock = _MUTEXES.get(key)
        if lock is None:
            lock = _MUTEXES[key] = threading.RLock()
        return lock


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if sys.platform == "win32":
        # NEVER os.kill(pid, 0) on Windows - it terminates the process.
        try:
            import ctypes

            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            STILL_ACTIVE = 259
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.OpenProcess(
                PROCESS_QUERY_LIMITED_INFORMATION, False, pid
            )
            if not handle:
                # Access denied means it exists (another user's process).
                return ctypes.get_last_error() == 5 or kernel32.GetLastError() == 5
            try:
                code = ctypes.c_ulong()
                if kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
                    return code.value == STILL_ACTIVE
                return True  # could not query; assume alive (conservative)
            finally:
                kernel32.CloseHandle(handle)
        except Exception:
            return True  # cannot tell; assume alive (never break a live lock)
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except OSError:
        return True


#: A token unique to THIS process instance, minted once at import. A bare PID
#: is not an identity: PIDs are recycled, so a process that inherited a dead
#: writer's number read that writer's leaked lockfile as its OWN (the
#: re-entrancy amnesty below) and never cleaned it up, while a process that
#: merely shared the number with some unrelated python.exe was reported as
#: the holder in the refusal message. The token settles both: re-entrancy is
#: token equality, and a lock carrying our PID but a foreign token is a
#: recycled-PID leftover, which is stale by definition.
_OWNER_TOKEN = f"{os.getpid()}-{uuid.uuid4().hex}"


def _process_create_time(pid: int) -> float | None:
    """A process's creation time (FILETIME ticks), or None when it cannot be
    read. Windows only (GetProcessTimes); elsewhere None, which leaves the
    plain PID liveness test."""
    if pid <= 0 or sys.platform != "win32":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.OpenProcess.restype = wintypes.HANDLE
        k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        k32.CloseHandle.argtypes = [wintypes.HANDLE]
        handle = k32.OpenProcess(0x1000, False, pid)  # QUERY_LIMITED_INFORMATION
        if not handle:
            return None
        try:
            times = [wintypes.FILETIME() for _ in range(4)]
            if not k32.GetProcessTimes(handle, *(ctypes.byref(x) for x in times)):
                return None
            born = times[0]
            return float((born.dwHighDateTime << 32) | born.dwLowDateTime)
        finally:
            k32.CloseHandle(handle)
    except Exception:  # noqa: BLE001
        return None


#: This process's creation time, written into each save journal entry so a
#: later server can tell this process from one that inherited its PID.
_OWNER_CREATED = _process_create_time(os.getpid())

#: A lockfile whose content we cannot parse is only assumed abandoned after
#: this long. Zero grace was the bug that made the lock leak under load: the
#: old acquire created the file with O_CREAT|O_EXCL and wrote the payload as
#: a SECOND syscall, so a concurrent acquirer that read it in between saw an
#: EMPTY file, parsed no pid, declared it stale, deleted a LIVE lock and took
#: its own. Two writers then believed they held it, and each one's release
#: unlinked the other's file. Writing the payload before the file is visible
#: (below) makes the window impossible; this grace covers a torn write from
#: any other source.
_UNREADABLE_GRACE_SECONDS = 30.0


def _read_lock_info(lock_path: Path) -> dict:
    try:
        info = json.loads(lock_path.read_text(encoding="utf-8"))
        if isinstance(info, dict):
            return info
    except (OSError, ValueError):
        pass
    return {}


def _publish_lockfile(lock_path: Path) -> bool:
    """Make a COMPLETE lockfile appear atomically, or report that one is
    already there. The payload is written to a private temp file first and
    linked into place, so the lockfile is never observable in a half-written
    state and its timestamp is minted at the instant it becomes visible (the
    old code computed the payload once BEFORE the retry loop, so a lock
    acquired after a wait was born already aged)."""
    payload = json.dumps({
        "pid": os.getpid(),
        "token": _OWNER_TOKEN,
        "time": time.time(),
        "host": _local_host(),
    })
    tmp = lock_path.parent / f".lock-{uuid.uuid4().hex}.tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(payload)
        try:
            os.link(str(tmp), str(lock_path))
            return True
        except FileExistsError:
            return False
        except OSError:
            # No hardlink support (non-NTFS, some network shares): fall back
            # to the exclusive-create path, still writing the payload before
            # closing the handle.
            try:
                fd = os.open(str(lock_path),
                             os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            except FileExistsError:
                return False
            try:
                os.write(fd, payload.encode("utf-8"))
            finally:
                os.close(fd)
            return True
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass


def _local_host() -> str:
    return os.environ.get("COMPUTERNAME") or socket.gethostname()


def _foreign_host(info: dict) -> bool:
    # Legacy locks without host metadata retain local-machine behavior.
    host = info.get("host")
    return bool(host) and (not isinstance(host, str)
                           or host.casefold() != _local_host().casefold())


def _is_ours(info: dict) -> bool:
    return not _foreign_host(info) and info.get("token") == _OWNER_TOKEN


def _is_stale(info: dict) -> bool:
    """Only local dead/recycled owners can be reclaimed; age is not proof."""
    if _foreign_host(info):
        return False
    pid = info.get("pid", -1)
    return not isinstance(pid, int) or not _pid_alive(pid) or pid == os.getpid()


def _acquire_lockfile(
    lock_path: Path, doc_name: str, *, adopt_own: bool = False,
    wait_seconds: float | None = None,
) -> bool:
    """Create the advisory lockfile. Returns True when this call created it
    (and must therefore remove it); False for a re-entrant same-process hold.
    Raises MutationLockTimeout when a live holder does not release in time.

    ``adopt_own`` is set by the outermost hold in this process. A lockfile
    with our token is then a leftover that an earlier release could not
    delete, not a hold, so this call takes it over (returns True) and the
    hold releases it when it ends."""
    waits = LOCK_WAIT_SECONDS if wait_seconds is None else wait_seconds
    deadline = time.monotonic() + waits
    while True:
        if _publish_lockfile(lock_path):
            return True
        info = _read_lock_info(lock_path)
        if _is_ours(info):
            if adopt_own:
                return True
            # This process instance already holds it; the in-process mutex
            # (held by the caller) is the real serializer and a nested
            # acquisition must not deadlock.
            return False
        pid = info.get("pid", -1)
        stamp = info.get("time", 0.0)
        age = time.time() - stamp if isinstance(stamp, (int, float)) else None
        if not info:
            # Unparseable: give a torn write a moment to settle, then treat a
            # persistently unreadable lock as abandoned.
            try:
                born = lock_path.stat().st_mtime
            except OSError:
                continue                      # it vanished; try again
            if time.time() - born > _UNREADABLE_GRACE_SECONDS:
                _break_lock(lock_path, info)
            elif time.monotonic() > deadline:
                raise MutationLockTimeout(
                    f"{doc_name} is locked by a write.lock this process "
                    "cannot read. Waited "
                    f"{int(waits)}s; retry, or delete "
                    f"{lock_path} if no other kitchensink4xl process is "
                    "running.")
            else:
                time.sleep(0.1)
            continue
        if _is_stale(info):
            _break_lock(lock_path, info)
            continue
        if time.monotonic() > deadline:
            holder = f"PID {pid}"
            if age is not None:
                holder += f", held for {int(age)}s"
            raise MutationLockTimeout(
                f"{doc_name} is being modified by another kitchensink4xl "
                f"process ({holder}). Waited {int(waits)}s; "
                "retry once that operation finishes."
            )
        time.sleep(0.1)


def _break_lock(lock_path: Path, judged: dict) -> None:
    """Delete a lock nobody holds (stale or unreadable), with one winner
    (core/lockrelease.py break_stale): ``judged`` is the record
    the caller read, or {} for a file it could not read, and the file goes
    only while it is still that one. Two waiters that had both read a dead
    holder's lockfile used to both delete "the lockfile"; the slower one
    deleted the live lock the faster one had just published, both then held
    the write lock, and two saves that both passed the stamp check before
    either promoted lost an update (MEASURED: two holders in 1 run of 10
    with a 250 ms pause between the read and the delete). A failure is left
    to the acquire loop, which looks again on its next pass."""
    lockrelease.break_stale(lock_path, judged, _read_lock_info,
                            _UNREADABLE_GRACE_SECONDS)


def _deletable(lock_path: Path) -> bool:
    """The lockfile is there and this process may remove it: it carries our
    token, or it cannot be read at all (the old release's rule)."""
    if not lockrelease.present(lock_path):
        return False
    info = _read_lock_info(lock_path)
    return not info or _is_ours(info)


def _release_lockfile(lock_path: Path) -> bool:
    """Remove the lockfile ONLY while it is still ours, and make sure it went.

    Unconditional unlink is how a leaked lock outlived every writer: once two
    processes believed they held the lock, each release deleted whichever file
    happened to be there, including a live one, and the last writer to publish
    was left with nobody to clean up after it.

    A failed unlink is retried (a waiter reading the file blocks the delete
    for a moment on Windows). Returns True once nothing of ours is left, and
    False when the file is still ours after every retry; write_lock then
    hands it to the sweeper."""
    return lockrelease.release(lock_path, _deletable)


@contextmanager
def write_lock(doc_path: str | os.PathLike, *, wait_seconds: float | None = None):
    """Serialize the full read-modify-validate-save cycle of one workbook.

    In-process mutex (per resolved path) + cross-process advisory lockfile in
    the workbook's slot folder. Hold this around XlsxPackage(...) through
    pkg.save() so every mutation sees the previous one's result and response
    metadata is computed against settled state.

    The outermost hold first puts back a workbook that a save which did not
    finish left aside (put_back_if_aside), so every write starts from
    the workbook in its place. ``wait_seconds`` shortens the wait for another
    holder (the start-up recovery uses it).

    Never on a thread marked with forbid_write_lock_on_this_thread (the COM
    worker): it raises RuntimeError there before waiting for anything.
    """
    why = getattr(_THREAD_RULES, "forbidden", None)
    if why:
        raise RuntimeError(
            f"write_lock({Path(os.fspath(doc_path)).name}) on a thread that "
            f"must never hold one: {why}")
    # Public entry that takes an arbitrary path; slot/backup paths all derive
    # from doc_path, so gating it here covers them too. No-op unless
    # KS4XL_ALLOWED_ROOTS is set.
    check_path(doc_path, "modify workbook")
    key = canonical_key(doc_path)
    mutex = _mutex_for(key)
    mutex.acquire()
    depth = _DEPTH.get(key, 0)
    _DEPTH[key] = depth + 1
    owns_lockfile = False
    lock_path: Path | None = None
    try:
        try:
            d = slot_dir(doc_path, create=True)
            lock_path = d / LOCK_FILE_NAME
        except OSError:
            lock_path = None  # cannot host a lockfile; mutex still serializes
        if lock_path is not None:
            owns_lockfile = _acquire_lockfile(
                lock_path, Path(doc_path).name, adopt_own=depth == 0,
                wait_seconds=wait_seconds,
            )
            if owns_lockfile:
                _SWEEPER.forget(key)
            if depth == 0:
                try:
                    put_back_if_aside(doc_path)
                except Exception:  # noqa: BLE001
                    pass  # the call's own read then says where the workbook is
        yield
    finally:
        try:
            if (
                owns_lockfile and lock_path is not None
                and not _release_lockfile(lock_path)
            ):
                _SWEEPER.add(key, lock_path, mutex, _deletable)
        finally:
            if depth:
                _DEPTH[key] = depth
            else:
                _DEPTH.pop(key, None)
            mutex.release()
