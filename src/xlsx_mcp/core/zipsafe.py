"""Checks on a workbook's ZIP package that need only its directory, and the
file name a picture part is extracted to.

`check_members` runs whenever a package is opened for reading
(xmlsafe.check_package and xmlsafe.load_workbook call it first), before any
part is inflated. It refuses two kinds of package:

- A decompression bomb. A part counts as inflating when its declared size is
  more than INFLATE_RATIO times what the package stores for it; when the
  inflating parts add up past INFLATE_BUDGET, the package is refused.
  Python's zipfile never returns more than a part's declared size, so the
  declared sizes bound what any read of the package can take. Measured on
  made-up parts (zlib at the level zipfile writes): workbook content
  deflates at about 4 to 30 to 1 (random numbers 3.8, a million formatted
  empty rows 28.5), a near-constant pivot cache at 142 to 1 and a
  flat-colour bitmap at 1,028 to 1; spaces or zeros at about 1,030 to 1.
  A read here costs up to about 6.5 times a part's size (a part of many
  short shared strings: 32 MB took 307 MB and 101 s for one set_cell), so
  128 MB of such parts may take about 0.9 GB, where the 1 GB part this
  stops would ask for about 6 GB. A large part that deflates like real
  content, or a large stored member, is not counted: the file is that big.
- A part named twice, or twice with different capital letters (package part
  names compare without regard to case, ECMA-376 Part 2). zipfile reads the
  last copy, an edit used to save both copies back, and other programs may
  pick either copy, so the same file could show two different workbooks.

`media_file_name` is the rule for manage_image extract: a picture part's file
name is the last segment of its part name, by either separator, and a
segment that cannot be a plain file name in the output folder on every
system this server runs on is refused, never repaired."""

from __future__ import annotations

import re
import zipfile

#: A part counts as inflating above this many declared bytes per stored byte.
INFLATE_RATIO = 100
#: The most the inflating parts of one package may declare in all.
INFLATE_BUDGET = 128 * 1024 * 1024


def _size(n: int) -> str:
    if n < 1024 * 1024:
        return f"{max(1, round(n / 1024)):,} KB"
    return f"{round(n / (1024 * 1024)):,} MB"


def _refusal(text: str):
    from .errors import UnsafePackage

    return UnsafePackage(text)


def inflates(info: zipfile.ZipInfo) -> bool:
    """True when the part declares more than INFLATE_RATIO times what the
    package stores for it."""
    return not info.is_dir() and info.file_size > INFLATE_RATIO * info.compress_size


def bomb_refusal(doc: str, counted: list[zipfile.ZipInfo], total: int):
    big = max(counted, key=lambda i: i.file_size)
    more = (f", and its highly compressed parts would take {_size(total)} in all"
            if len(counted) > 1 else "")
    return _refusal(
        f"{doc}: its part {big.filename} would expand from {_size(big.compress_size)} "
        f"in the file to {_size(big.file_size)} when read{more}. A package like that "
        "is built to exhaust memory, and no workbook needs one, so the file was not "
        "read. Nothing was changed. If you expected this file to open, ask whoever "
        "sent it for a new copy.")


def duplicate_refusal(doc: str, first: str, second: str):
    if first == second:
        what = f"lists the part {first} twice"
    else:
        what = (f"holds {first} and {second}, which name the same part (package part "
                "names ignore capital letters)")
    return _refusal(
        f"{doc}: its ZIP package {what}. The two copies can differ, and a program may "
        "read either one, so the file was not read. Nothing was changed. If Excel opens "
        "the file, save it again from Excel and try again; otherwise ask whoever sent "
        "it for a new copy.")


def check_members(package, *, doc: str) -> None:
    """Refuse a package that inflates past the budget or names a part twice,
    from its directory alone: nothing is inflated. `package` is a ZipFile or
    its infolist()."""
    infos = package.infolist() if hasattr(package, "infolist") else list(package)
    counted = [i for i in infos if inflates(i)]
    total = sum(i.file_size for i in counted)
    if total > INFLATE_BUDGET:
        raise bomb_refusal(doc, counted, total)
    seen: dict[str, str] = {}
    for info in infos:
        if info.is_dir():
            continue
        key = info.filename.lower()
        if key in seen:
            raise duplicate_refusal(doc, seen[key], info.filename)
        seen[key] = info.filename


# --------------------------------------------------------- picture files

#: Names Windows opens as a device, with or without an extension.
_DEVICES = frozenset(
    {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$", "CLOCK$"}
    | {f"{d}{n}" for d in ("COM", "LPT") for n in "123456789¹²³"})
_FORBIDDEN = re.compile(r'[<>"|?*\x00-\x1f]')


def media_file_name(member: str) -> tuple[str | None, str | None]:
    """(file name, None): the plain file name a picture part `member` is
    extracted to, the last segment of its name by either separator. (None,
    why) when that segment cannot be one."""
    base = re.split(r"[\\/]", member)[-1]
    if base in ("", ".", ".."):
        return None, "it names a folder, not a file"
    if ":" in base:
        return None, "it holds a colon, which Windows reads as a drive or a hidden stream"
    if _FORBIDDEN.search(base):
        return None, "it holds characters a Windows file name cannot"
    if base[-1] in ". ":
        return None, "it ends in a dot or a space, which Windows drops from a file name"
    if base.partition(".")[0].rstrip(" ").upper() in _DEVICES:
        return None, "Windows reserves that name for a device"
    if len(base.encode("utf-8")) > 255:
        return None, "it is longer than a file name can be"
    return base, None


def _shown(member: str) -> str:
    text = repr(member)
    return text if len(text) <= 120 else text[:100] + "...'"


def media_name_refusal(doc: str, member: str, why: str):
    return _refusal(
        f"{doc}: its picture part {_shown(member)} cannot be saved as a file in the "
        f"output folder ({why}), so no picture was extracted. Nothing was changed. "
        "If Excel opens the workbook, save it again from Excel and extract again.")


__all__ = ["INFLATE_RATIO", "INFLATE_BUDGET", "inflates", "check_members", "bomb_refusal",
           "duplicate_refusal", "media_file_name", "media_name_refusal"]
