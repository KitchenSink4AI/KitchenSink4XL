"""Opt-in path sandboxing governed by the KS4XL_ALLOWED_ROOTS environment
variable.

Ported VERBATIM from KitchenSink4Word core/sandbox.py; only the env-var
name changed (KS4W_ALLOWED_ROOTS -> KS4XL_ALLOWED_ROOTS) and the error base
(WordMcpError -> XlMcpError). The canonicalization logic is format-agnostic
and battle-tested in two shipped siblings, so it is not re-derived here.

When KS4XL_ALLOWED_ROOTS is unset (or empty), nothing here restricts
anything: check_path returns its input untouched. When it is set to an
os.pathsep-separated list of directories (Windows example:
"C:\\Users\\me\\Documents;D:\\Work"), every filesystem path the server
touches must resolve inside one of those directories or the call refuses
with SandboxViolation before any file is opened.

The check applies to READS as well as writes. A read outside the sandbox
exfiltrates workbook content (Power Query connection strings, pivot-cache
source data, cached external-link values) to the calling agent just as
surely as a write plants content, so workbook opens, image reads, data
imports, and backup listings are all gated.

Canonicalization defeats the classic escapes before containment is tested:

- relative traversal ("..\\") collapses via os.path.abspath
- symlinks and NTFS junctions resolve via os.path.realpath BEFORE the
  containment comparison (realpath also expands 8.3 short names where the
  target exists)
- extended-length prefixes ("\\\\?\\C:\\...", "\\\\?\\UNC\\srv\\share")
  are normalized to their plain spellings
- UNC paths ("\\\\server\\share\\...") are refused outright unless an
  allowed root is itself a UNC path that contains them
- case differences are neutralized with os.path.normcase on both sides
- prefix collisions cannot match: containment compares with a trailing
  separator, so root "C:\\U\\Documents" never admits "C:\\U\\Documents2"
- nonexistent paths (create targets) canonicalize their deepest EXISTING
  ancestor, then rejoin the untraversed tail for the comparison

Roots are parsed lazily and re-parsed whenever the raw environment value
changes, so tests (and long-lived processes) see updates immediately.
"""

from __future__ import annotations

import os

from .errors import XlMcpError

ENV_VAR = "KS4XL_ALLOWED_ROOTS"


class SandboxViolation(XlMcpError):
    """A path resolved outside the allowed roots; nothing was read or
    written. KS4XL_ALLOWED_ROOTS governs which directories are allowed."""


# ------------------------------------------------------------------ parsing

_cache_raw: object = object()  # sentinel: never equal to any env string
_cache_roots: tuple[tuple[str, str], ...] = ()  # (as-given, canonical)


def _allowed_roots() -> tuple[tuple[str, str], ...]:
    """Parsed (display, canonical) roots; cached until the env value changes."""
    global _cache_raw, _cache_roots
    raw = os.environ.get(ENV_VAR)
    if raw != _cache_raw:
        entries: list[tuple[str, str]] = []
        if raw:
            for part in raw.split(os.pathsep):
                part = part.strip().strip('"').strip()
                if part:
                    entries.append((part, _canonicalize(part)))
        _cache_roots = tuple(entries)
        _cache_raw = raw
    return _cache_roots


def active() -> bool:
    """True when sandboxing is in force (env set to a non-empty root list)."""
    return bool(_allowed_roots())


def root_count() -> int:
    """How many roots are allowed. A count rather than the paths themselves:
    get_server_info reports this so an installer can confirm the setting
    arrived, and a directory layout is not the client's business."""
    return len(_allowed_roots())


# --------------------------------------------------------- canonicalization


def _strip_extended_prefix(p: str) -> str:
    r"""\\?\C:\x -> C:\x and \\?\UNC\srv\share -> \\srv\share."""
    if p.startswith("\\\\?\\UNC\\") or p.startswith("//?/UNC/"):
        return "\\\\" + p[8:]
    if p.startswith("\\\\?\\") or p.startswith("//?/"):
        return p[4:]
    return p


def _is_unc(p: str) -> bool:
    q = p.replace("/", "\\")
    return q.startswith("\\\\") and not q.startswith("\\\\?\\")


def _canonicalize(path_str: str) -> str:
    """Absolute + symlink/junction/8.3-resolved form of a path that may not
    exist: the deepest existing ancestor is realpath-resolved, then the
    nonexistent tail (already lexically collapsed by abspath) is rejoined."""
    ap = _strip_extended_prefix(os.path.abspath(path_str))
    if os.path.exists(ap):
        return _strip_extended_prefix(os.path.realpath(ap))
    head = ap
    tail: list[str] = []
    while head and not os.path.exists(head):
        head, last = os.path.split(head)
        if not last:  # reached an unsplittable head (drive/share root)
            break
        tail.append(last)
    if head and os.path.exists(head):
        head = _strip_extended_prefix(os.path.realpath(head))
    else:
        head = head or ap
    for last in reversed(tail):
        head = os.path.join(head, last)
    return head


def _contained(root_canonical: str, candidate_canonical: str) -> bool:
    """Trailing-separator containment on normcased canonical paths, so a
    root never admits a sibling that merely shares its name as a prefix."""
    r = os.path.normcase(root_canonical).rstrip("\\/")
    c = os.path.normcase(candidate_canonical)
    if r.endswith(":"):  # bare drive root like "c:" -> compare as "c:\"
        r += os.sep
        return c == r.rstrip(os.sep) or c.startswith(r)
    if not r:  # POSIX filesystem root "/"
        return c.startswith("/")
    if c == r:
        return True
    return c.startswith(r + "\\") or c.startswith(r + "/")


# ------------------------------------------------------------------- check


#: Reserved Windows device names. A path whose last
#: component is one of these does not address a file: an INPUT path "CON" or
#: "CONIN$" blocked on a console read until the call's deadline and left the
#: thread stuck for the life of the server; an OUTPUT path swallowed the
#: write (core/outguard.py). Windows matches on the stem, case-insensitively,
#: with trailing dots and spaces stripped.
WINDOWS_DEVICE_NAMES = frozenset({
    "CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$", "CLOCK$",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
})


#: Device-namespace prefixes, in either slash spelling.
_DEVICE_PREFIXES = ("\\\\.\\", "\\\\?\\", "//./", "//?/", "\\??\\")


def device_namespace(path: str | os.PathLike) -> bool:
    r"""True for a path in the Windows device namespace (\\.\..., //./...,
    \??\...) or under the \\?\GLOBALROOT object-manager root. The
    \\?\ long-path prefix on a drive or UNC path is NOT the device
    namespace and answers False."""
    text = os.fspath(path).strip().replace("/", "\\")
    if text.startswith("\\\\.\\") or text.startswith("\\??\\"):
        return True
    if text.startswith("\\\\?\\"):
        rest = text[4:].lstrip("\\")
        return rest.split("\\", 1)[0].upper() == "GLOBALROOT"
    return False


def device_name(path: str | os.PathLike) -> str | None:
    r"""The reserved device name a path's last component names, or None.

    The device namespace (\\.\CON, \\?\CONIN$) and a
    trailing colon (CON:, CONIN$:) name the same devices; os.path.basename
    alone returned nothing for the first and kept the colon on the second,
    so both hung until the call deadline."""
    text = os.fspath(path).strip()
    for prefix in _DEVICE_PREFIXES:
        if text.startswith(prefix):
            text = text[len(prefix):]
            break
    text = text.rstrip("\\/").rstrip(": ")
    name = text.replace("/", "\\").rsplit("\\", 1)[-1]
    stem = name.split(".")[0].strip().rstrip(".: ").upper()
    return stem if stem in WINDOWS_DEVICE_NAMES else None


def check_path(path: str | os.PathLike, purpose: str = "access") -> str:
    """Gate one path. Returns the path unchanged when sandboxing is off;
    returns the canonicalized path when it is on and the path is inside an
    allowed root; raises SandboxViolation otherwise. `purpose` names the
    operation in the error text ("open workbook", "export data output")."""
    p = os.fspath(path)
    if isinstance(p, bytes):  # never happens in this codebase; be safe
        p = os.fsdecode(p)
    if "\x00" in p:
        raise SandboxViolation(
            f"refusing to {purpose}: path contains a null byte"
        )
    if os.name == "nt":
        # The device namespace (\\.\ and \??\) and the
        # \\?\GLOBALROOT object-manager root never address a workbook, and a
        # console behind them (\\.\GLOBALROOT\Device\ConDrv\Console) blocked
        # a read until the call deadline. Refused outright. The \\?\ long-path
        # form of an ordinary path (\\?\C:\..., \\?\UNC\...) keeps working.
        if device_namespace(p):
            raise XlMcpError(
                f"refusing to {purpose}: {p} is a Windows device path, not a file. "
                "Pass the path of a real file.")
        # Before any open, on every path this server is handed.
        device = device_name(p)
        if device is not None:
            raise XlMcpError(
                f"refusing to {purpose}: {p} names the reserved Windows device "
                f"{device}, not a file. Pass the path of a real file.")
    roots = _allowed_roots()
    if not roots:
        return p

    if _is_unc(_strip_extended_prefix(p)) and not any(
        _is_unc(canon) for _display, canon in roots
    ):
        raise SandboxViolation(
            f"refusing to {purpose}: {p} is a UNC network path and no "
            f"allowed root is a UNC path. {ENV_VAR} restricts file access "
            f"to: {_roots_display(roots)}."
        )

    canonical = _canonicalize(p)
    for _display, root_canonical in roots:
        if _contained(root_canonical, canonical):
            return canonical

    raise SandboxViolation(
        f"refusing to {purpose}: {p} (resolves to {canonical}) is outside "
        f"the allowed roots. {ENV_VAR} restricts file access to: "
        f"{_roots_display(roots)}. Add the needed directory to "
        f"{ENV_VAR} (or unset it) to allow this path."
    )


def _roots_display(roots: tuple[tuple[str, str], ...]) -> str:
    return "; ".join(display for display, _canon in roots)
