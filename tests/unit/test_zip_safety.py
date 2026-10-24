"""A workbook package is checked from its ZIP directory before any part is
inflated: a decompression bomb is refused, and so is a part named twice. And
extracting pictures writes only plain files inside the folder asked for.

Before this:
- a part that declares 1 GB and stores about 1 MB (it deflates at about
  1,000 to 1) was read in full, at about 3 to 6.5 times its size in memory:
  measured at 32 MB, a part of many one-letter shared strings took 307 MB
  and 101 s for one set_cell, so a 1 GB part asks for about 6 GB;
- a part named twice was read from its last copy, and an edit saved both
  copies back; a part named twice with different capital letters made an
  edit fail its own verification;
- manage_image extract joined the folder with each picture's own name, so a
  name such as "Q:made_up.png" (another drive), "made_up.png:hidden" (a
  hidden stream), a device name or a name ending in a dot wrote outside the
  folder, into a hidden stream or under another name, and two pictures with
  the same name wrote one file and then failed.

Every workbook here is made up in code (hostile_xlsx.py). The bomb tests spy
on zipfile: opening a bomb part fails the test, so not even the code before
this change can inflate one here. Picture writes are spied on too: a write
outside the test's own folder is refused and recorded. Excel cannot start in
this file."""

from __future__ import annotations

import hashlib
import json
import os
import time
import tracemalloc
import zipfile
from pathlib import Path

import pytest

import hostile_xlsx as hx

MiB = 1 << 20
#: The limits under test, written out so that changing them is a visible
#: decision (zipsafe.py has the measurements behind them).
RATIO = 100
BUDGET = 128 * MiB

BOMB_WORDS = "would expand"
TWICE_WORDS = "twice"
SAME_PART_WORDS = "name the same part"
NAME_WORDS = "no picture was extracted"

SHEET = "xl/worksheets/sheet1.xml"


# ------------------------------------------------------------------ fixtures

class _BombRead(Exception):
    """The spy's answer when anything opens a bomb part for reading."""


class _Spy:
    def __init__(self):
        self.forbidden: set[str] = set()
        self.opened: list[str] = []

    def bomb_reads(self) -> list[str]:
        return [n for n in self.opened if n in self.forbidden]


@pytest.fixture
def spy(monkeypatch):
    """Records every member any code opens for reading (zipfile's read()
    goes through open() too). Opening one in `spy.forbidden` raises."""
    real_open = zipfile.ZipFile.open
    s = _Spy()

    def open_(zf, name, mode="r", pwd=None, *, force_zip64=False):
        member = getattr(name, "filename", name)
        if mode == "r":
            s.opened.append(member)
            if member in s.forbidden:
                raise _BombRead(member)
        return real_open(zf, name, mode, pwd, force_zip64=force_zip64)

    monkeypatch.setattr(zipfile.ZipFile, "open", open_)
    return s


@pytest.fixture
def no_excel(monkeypatch):
    """Excel cannot start: the COM tier's spawn fails the test instead."""
    from xlsx_mcp.com import instances

    def refuse(*_a, **_k):
        raise AssertionError("Excel must not start in this test")

    monkeypatch.setattr(instances.ExcelInstanceManager, "_new_app", refuse)
    try:
        import win32com.client as w32
    except ImportError:
        return
    monkeypatch.setattr(w32, "DispatchEx", refuse)
    monkeypatch.setattr(w32, "Dispatch", refuse)


def _inside(path: str, folder: Path) -> bool:
    p = os.path.normcase(os.path.abspath(path))
    return p.startswith(os.path.normcase(str(folder)) + os.sep)


_DEVICES = {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"} | {
    f"{d}{n}" for d in ("COM", "LPT") for n in "123456789"}


def _device(path: str) -> bool:
    """A name Windows may open as a device rather than a file."""
    base = path.replace("\\", "/").rsplit("/", 1)[-1]
    return base.partition(".")[0].rstrip(" ").upper() in _DEVICES


@pytest.fixture
def writes(monkeypatch, tmp_path):
    """Every file manage_image opens for writing. One outside the test's
    folder, or one Windows could open as a device, is refused (nothing
    reaches it) and recorded in `outside`."""
    from xlsx_mcp.ops import objects

    real_open = open
    log = {"inside": [], "outside": []}

    def open_(file, mode="r", *args, **kwargs):
        if any(c in mode for c in "wxa+"):
            target = os.fspath(file)
            if not _inside(target, tmp_path) or _device(target):
                log["outside"].append(target)
                raise PermissionError(f"the test's write spy refuses {target}")
            log["inside"].append(target)
        return real_open(file, mode, *args, **kwargs)

    monkeypatch.setattr(objects, "open", open_, raising=False)
    return log


def _digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _tree(folder: Path) -> set[str]:
    return {str(p.relative_to(folder)) for p in folder.rglob("*")}


def _entries():
    from xlsx_mcp.core import hazard, package, xmlsafe
    from xlsx_mcp.ops import gridio

    return {
        "check_file": lambda p: xmlsafe.check_file(p),
        "scan_path": lambda p: hazard.scan_path(p),
        "read_model": lambda p: gridio.open_wb(p),
        "edit_model": lambda p: package.WorkbookPackage.open(p).workbook,
        "loader": lambda p: xmlsafe.load_workbook(p),
    }


ENTRIES = ("check_file", "scan_path", "read_model", "edit_model", "loader")


def _assert_plain(msg: str, *words: str):
    for w in words:
        assert w in msg, msg
    assert "Nothing was changed" in msg, msg
    for raw in ("Traceback", "BadZipFile", "Errno", "zlib", "_BombRead"):
        assert raw not in msg, msg


#: Ops that read no workbook part: they make, copy or list files, or answer
#: without a file (the update ops take no path at all).
_NO_PART_READ = {"create_workbook", "copy_workbook", "manage_backups", "get_server_info",
                 "get_workflows", "com_status", "update_now", "update_notices"}


def _examples() -> dict[str, tuple[str, dict]]:
    import family_harness as fh
    from xlsx_mcp import registry

    h = fh.start()
    out = {}
    for name in sorted(registry.public_ops()):
        if name in _NO_PART_READ:
            continue
        pl, _err, _ = h.call("xl_help", {"op": name})
        out[name] = (pl["tool"], pl["example"]["args"])
    return out


def _every_op_refuses(f: Path, tmp_path: Path, words: tuple[str, ...]) -> dict:
    import family_harness as fh

    h = fh.start()
    wrong = {}
    for op, (tool, example) in _examples().items():
        args = dict(example, path=str(f))
        if "output" in args:
            args["output"] = str(tmp_path / ("out_" + op + Path(args["output"]).suffix))
        payload, is_error, _ = h.call(tool, {"op": op, "args": args})
        err = payload.get("error") if isinstance(payload, dict) else None
        if not (is_error and err and err.get("code") == "UNSUPPORTED_CONTENT"
                and all(w in err.get("message", "") for w in words)):
            wrong[op] = payload
    return wrong


# =========================================================== 5a. the bombs

#: Where a bomb part sits: three parts every read or edit inflates, a member
#: no reader parses (an edit copies it into the saved file), and a picture.
PLACEMENTS = {
    "sheet": SHEET,
    "shared_strings": "xl/sharedStrings.xml",
    "workbook": "xl/workbook.xml",
    "unreferenced_member": "xl/made_up_zeros.bin",
    "picture": "xl/media/made_up_zeros.png",
}


@pytest.fixture(scope="module")
def bombs(tmp_path_factory) -> dict[str, Path]:
    """One made-up workbook per placement, its bomb part declared at 8 MiB
    past the budget (each file about 0.15 to 0.4 MB)."""
    folder = tmp_path_factory.mktemp("bombs")
    return {k: hx.make_bomb_xlsx(folder / f"{k}.xlsx", {name: BUDGET + 8 * MiB})
            for k, name in PLACEMENTS.items()}


def test_the_made_up_bombs_are_bombs(bombs):
    for key, f in bombs.items():
        with zipfile.ZipFile(f) as z:
            info = z.getinfo(PLACEMENTS[key])
        assert info.file_size > BUDGET and info.file_size > 300 * info.compress_size, key
        assert f.stat().st_size < 1 * MiB, key


@pytest.mark.parametrize("entry", ENTRIES)
@pytest.mark.parametrize("placement", sorted(PLACEMENTS))
def test_a_bomb_is_refused_from_the_zip_directory_before_any_part_is_inflated(
        placement, entry, bombs, spy):
    from xlsx_mcp.core.errors import UnsafePackage

    f = bombs[placement]
    spy.forbidden = {PLACEMENTS[placement]}
    t0 = time.perf_counter()
    with pytest.raises(UnsafePackage) as err:
        _entries()[entry](str(f))
    assert time.perf_counter() - t0 < 2.0
    _assert_plain(str(err.value), BOMB_WORDS, PLACEMENTS[placement])
    assert spy.bomb_reads() == []


@pytest.mark.parametrize("placement", ["sheet", "unreferenced_member", "picture"])
def test_the_tools_refuse_a_bomb_with_unsupported_content(placement, bombs, spy, tmp_path,
                                                          no_excel):
    import family_harness as fh

    f = tmp_path / "bomb.xlsx"
    f.write_bytes(bombs[placement].read_bytes())
    before = _digest(f)
    spy.forbidden = {PLACEMENTS[placement]}
    out = tmp_path / "pictures"
    out.mkdir()
    h = fh.start()
    for tool, op, args in [
            ("xl_read", "read_range", {"location": "A1:B2", "values": "both"}),
            ("xl_read", "diagnose_workbook", {}),
            ("xl_edit", "set_cell", {"location": "B1", "value": 43}),
            ("xl_format", "manage_image", {"action": "list"}),
            ("xl_format", "manage_image", {"action": "extract", "out_dir": str(out)})]:
        t0 = time.perf_counter()
        payload, is_error, _ = h.call(tool, {"op": op, "args": {"path": str(f), **args}})
        assert time.perf_counter() - t0 < 5.0, op
        assert is_error and payload["error"]["code"] == "UNSUPPORTED_CONTENT", (op, payload)
        _assert_plain(payload["error"]["message"], BOMB_WORDS, PLACEMENTS[placement])
    assert _digest(f) == before
    assert list(out.iterdir()) == []
    assert spy.bomb_reads() == []


@pytest.mark.parametrize("placement", ["sheet", "unreferenced_member"])
def test_every_op_that_reads_a_workbook_refuses_a_bomb_before_reading_it(
        placement, bombs, spy, tmp_path, no_excel):
    """The COM tier included: refused before Excel would start."""
    f = tmp_path / "bomb.xlsx"
    f.write_bytes(bombs[placement].read_bytes())
    before = _digest(f)
    spy.forbidden = {PLACEMENTS[placement]}
    assert _every_op_refuses(f, tmp_path, (BOMB_WORDS,)) == {}
    assert _digest(f) == before
    assert spy.bomb_reads() == []


@pytest.mark.parametrize("member", [SHEET, "xl/made_up_zeros.bin"])
def test_a_one_gib_bomb_is_refused_in_well_under_a_second_with_little_memory(
        member, tmp_path, spy, no_excel):
    """Real size: 1 GiB declared in a file of about 1 MB."""
    import family_harness as fh

    f = hx.make_bomb_xlsx(tmp_path / "gib.xlsx", {member: 1 << 30})
    assert f.stat().st_size < 2 * MiB
    spy.forbidden = {member}
    h = fh.start()
    calls = [("xl_read", "read_range", {"location": "A1:B2"}),
             ("xl_edit", "set_cell", {"location": "B1", "value": 43})]
    # The first call of a fresh process imports the server's modules, which
    # on a slow runner takes seconds and is not what the bound measures. Make
    # each call once untimed, then measure the refusal itself.
    for tool, op, args in calls:
        h.call(tool, {"op": op, "args": {"path": str(f), **args}})
    for tool, op, args in calls:
        tracemalloc.start()
        t0 = time.perf_counter()
        payload, is_error, _ = h.call(tool, {"op": op, "args": {"path": str(f), **args}})
        took = time.perf_counter() - t0
        peak = tracemalloc.get_traced_memory()[1]
        tracemalloc.stop()
        assert is_error and payload["error"]["code"] == "UNSUPPORTED_CONTENT", (op, payload)
        _assert_plain(payload["error"]["message"], BOMB_WORDS, member, "1,024 MB")
        assert took < 1.0, (op, took)
        assert peak < 64 * MiB, (op, peak)
    assert spy.bomb_reads() == []


def test_many_small_bombs_add_up_to_a_refusal(tmp_path, spy):
    """Nine parts of 16 MiB each, every one alone under the budget."""
    from xlsx_mcp.core import xmlsafe
    from xlsx_mcp.core.errors import UnsafePackage

    names = [f"xl/made_up_{i}.bin" for i in range(9)]
    f = hx.make_bomb_xlsx(tmp_path / "many.xlsx", {n: 16 * MiB for n in names})
    spy.forbidden = set(names)
    with pytest.raises(UnsafePackage) as err:
        xmlsafe.check_file(str(f))
    _assert_plain(str(err.value), BOMB_WORDS, "144 MB in all")
    assert spy.bomb_reads() == []


def test_compressible_parts_under_the_budget_are_read(tmp_path):
    """Seven parts of 16 MiB that deflate at 1,000 to 1 (112 MiB in all), and a
    sheet of 24 MiB of spaces, are under the budget: the package opens."""
    import family_harness as fh
    from xlsx_mcp.core import xmlsafe

    f = hx.make_bomb_xlsx(tmp_path / "under.xlsx",
                          {**{f"xl/made_up_{i}.bin": 16 * MiB for i in range(7)}})
    xmlsafe.check_file(str(f))
    g = hx.make_bomb_xlsx(tmp_path / "sheet.xlsx", {SHEET: 24 * MiB})
    payload, is_error, _ = fh.start().call("xl_read", {"op": "read_range", "args": {
        "path": str(g), "location": "A1:B2"}})
    assert not is_error and payload["values"] == [["Made-up label", 42],
                                                  ["Made-up inline text", 7.5]], payload


def _info(name: str, file_size: int, compress_size: int) -> zipfile.ZipInfo:
    info = zipfile.ZipInfo(name)
    info.file_size, info.compress_size = file_size, compress_size
    return info


def test_the_limits_are_the_measured_ones():
    from xlsx_mcp.core import zipsafe

    assert zipsafe.INFLATE_RATIO == RATIO
    assert zipsafe.INFLATE_BUDGET == BUDGET


@pytest.mark.parametrize("infos, refused", [
    # A part at exactly 100 to 1 does not count; one byte more does.
    ([_info("xl/a.bin", 200 * MiB, 2 * MiB)], False),
    ([_info("xl/a.bin", 200 * MiB + 1, 2 * MiB)], True),
    # The budget is a total: at it is fine, past it is not.
    ([_info("xl/a.bin", BUDGET // 2, 1), _info("xl/b.bin", BUDGET // 2, 1)], False),
    ([_info("xl/a.bin", BUDGET // 2, 1), _info("xl/b.bin", BUDGET // 2 + 1, 1)], True),
    # Real workbook content deflates at 4 to 30 to 1: a very large part at
    # that ratio is not a bomb, nor is a large stored (uncompressed) member.
    ([_info(SHEET, 4 << 30, (4 << 30) // 30)], False),
    ([_info("xl/media/made_up.mp4", 3 << 30, 3 << 30)], False),
    # A part that declares a size and stores nothing counts.
    ([_info("xl/a.bin", BUDGET + 1, 0)], True),
    # Folder entries are not parts.
    ([_info("xl/", BUDGET * 4, 0)], False),
])
def test_the_bomb_rule(infos, refused):
    from xlsx_mcp.core import zipsafe
    from xlsx_mcp.core.errors import UnsafePackage

    if refused:
        with pytest.raises(UnsafePackage):
            zipsafe.check_members(infos, doc="made-up.xlsx")
    else:
        zipsafe.check_members(infos, doc="made-up.xlsx")


# ======================================================= 5c. names repeated

def _second(text: str = "Made-up second copy") -> str:
    return hx.base_parts()[SHEET].replace("Made-up inline text", text)


#: case -> (members after the made-up workbook's own, or a full member list,
#: and the part names the refusal must name).
REPEATS = {
    "sheet_twice_second_differs": ([(SHEET, _second())], [SHEET]),
    "sheet_twice_second_first": (None, [SHEET]),
    "sheet_capital_twin": ([("xl/worksheets/Sheet1.xml", _second())],
                           ["xl/worksheets/Sheet1.xml", SHEET]),
    "content_types_twice": ([("[Content_Types].xml", hx.base_parts()["[Content_Types].xml"])],
                            ["[Content_Types].xml"]),
    "workbook_rels_twice": ([("xl/_rels/workbook.xml.rels",
                              hx.base_parts()["xl/_rels/workbook.xml.rels"])],
                            ["xl/_rels/workbook.xml.rels"]),
    "picture_twice": ([("xl/media/made_up.png", b"made-up one"),
                       ("xl/media/made_up.png", b"made-up two")], ["xl/media/made_up.png"]),
}


def _repeat_file(tmp_path: Path, case: str) -> Path:
    extra, _names = REPEATS[case]
    f = tmp_path / f"{case}.xlsx"
    if extra is None:  # the second copy comes first in the package
        return hx.make_package(f, [(SHEET, _second())] + hx.base_members())
    return hx.make_package(f, hx.base_members() + extra)


@pytest.mark.parametrize("entry", ENTRIES)
@pytest.mark.parametrize("case", sorted(REPEATS))
def test_a_part_named_twice_is_refused(case, entry, tmp_path):
    from xlsx_mcp.core.errors import UnsafePackage

    f = _repeat_file(tmp_path, case)
    t0 = time.perf_counter()
    with pytest.raises(UnsafePackage) as err:
        _entries()[entry](str(f))
    assert time.perf_counter() - t0 < 2.0
    names = REPEATS[case][1]
    words = (TWICE_WORDS,) if len(names) == 1 else (SAME_PART_WORDS,)
    _assert_plain(str(err.value), *words, *names)


@pytest.mark.parametrize("case", sorted(REPEATS))
def test_the_tools_refuse_a_part_named_twice(case, tmp_path, no_excel):
    import family_harness as fh

    f = _repeat_file(tmp_path, case)
    before = _digest(f)
    out = tmp_path / "pictures"
    out.mkdir()
    h = fh.start()
    for tool, op, args in [
            ("xl_read", "read_range", {"location": "A1:B2", "values": "both"}),
            ("xl_read", "diagnose_workbook", {}),
            ("xl_edit", "set_cell", {"location": "B1", "value": 43}),
            ("xl_format", "manage_image", {"action": "extract", "out_dir": str(out)})]:
        payload, is_error, _ = h.call(tool, {"op": op, "args": {"path": str(f), **args}})
        assert is_error and payload["error"]["code"] == "UNSUPPORTED_CONTENT", (op, payload)
        _assert_plain(payload["error"]["message"], REPEATS[case][1][0])
        assert "Made-up second copy" not in json.dumps(payload)
    assert _digest(f) == before
    assert list(out.iterdir()) == []


def test_every_op_that_reads_a_workbook_refuses_a_part_named_twice(tmp_path, no_excel):
    f = _repeat_file(tmp_path, "sheet_twice_second_differs")
    before = _digest(f)
    assert _every_op_refuses(f, tmp_path, (TWICE_WORDS, SHEET)) == {}
    assert _digest(f) == before


def test_folder_entries_are_not_parts_named_twice(tmp_path):
    """Some zip tools add an entry per folder; they name no part."""
    import family_harness as fh

    f = hx.make_package(tmp_path / "folders.xlsx",
                        [("xl/", b""), ("xl/worksheets/", b""), ("XL/", b"")] + hx.base_members())
    payload, is_error, _ = fh.start().call("xl_read", {"op": "read_range", "args": {
        "path": str(f), "location": "A1:B2"}})
    assert not is_error, payload


# ============================================ 5b. extracting the pictures

def _other_drive(tmp_path: Path) -> str:
    here = os.path.splitdrive(str(tmp_path))[0][:1].upper()
    return "R" if here == "Q" else "Q"


#: Picture names that cannot be a plain file name in the output folder.
#: Each sits after a normal picture whose name sorts first, so a partial
#: extract would show.
UNSAFE = {
    "other_drive": "{drive}:made_up_escape.png",
    "hidden_stream": "made_up.png:hidden",
    "parent": "..",
    "self": ".",
    "device_nul": "NUL.png",
    "device_con": "con",
    "device_com1": "COM1.made_up",
    "device_lpt9": "lpt9",
    "trailing_dot": "made_up.png.",
    "trailing_space": "made_up.png ",
    "angle_brackets": "made_up<1>.png",
    "quote": 'made_up".png',
    "pipe": "made_up|.png",
    "question": "made_up?.png",
    "star": "made_up*.png",
    "control_char": "made\x01up.png",
    "too_long": "m" * 300 + ".png",
}


@pytest.mark.parametrize("case", sorted(UNSAFE))
def test_extract_refuses_a_picture_name_that_is_not_a_plain_file_name(
        case, tmp_path, writes):
    import family_harness as fh

    name = UNSAFE[case].format(drive=_other_drive(tmp_path))
    member = "xl/media/" + name
    book = hx.make_package(tmp_path / "book.xlsx", hx.base_members() + [
        ("xl/media/!made_up_first.png", b"made-up first picture"),
        (member, b"made-up picture")])
    out = tmp_path / "pictures"
    out.mkdir()
    before = _tree(tmp_path)
    payload, is_error, _ = fh.start().call("xl_format", {"op": "manage_image", "args": {
        "path": str(book), "action": "extract", "out_dir": str(out)}})
    assert writes["outside"] == []
    assert is_error and payload["error"]["code"] == "UNSUPPORTED_CONTENT", payload
    _assert_plain(payload["error"]["message"], NAME_WORDS)
    assert list(out.iterdir()) == [], "a refusal writes nothing"
    assert _tree(tmp_path) == before


#: Hostile names whose last segment is a plain name: the picture lands in the
#: output folder under that name, and nowhere else.
HARMLESS = {
    "dot_dot_slash": "xl/media/../../made_up_escape.png",
    "dot_dot_backslash": "xl/media/..\\..\\made_up_escape.png",
    "drive_after_prefix": "xl/media/C:/made_up_escape.png",
    "drive_backslash_after_prefix": "xl/media/C:\\made_up_escape.png",
    "double_slash": "xl/media//made_up_escape.png",
    "deeper_folder": "xl/media/sub/made_up_escape.png",
}


@pytest.mark.parametrize("case", sorted(HARMLESS))
def test_extract_keeps_a_traversal_name_inside_the_folder(case, tmp_path, writes):
    from xlsx_mcp.ops import objects

    book = hx.make_package(tmp_path / "book.xlsx", hx.base_members() + [
        (HARMLESS[case], b"made-up picture")])
    out = tmp_path / "pictures"
    out.mkdir()
    before = _tree(tmp_path)
    got = objects.manage_image(str(book), "extract", out_dir=str(out))
    assert writes["outside"] == []
    assert [Path(p).name for p in got["extracted"]] == ["made_up_escape.png"]
    assert (out / "made_up_escape.png").read_bytes() == b"made-up picture"
    assert _tree(tmp_path) - before == {"pictures" + os.sep + "made_up_escape.png"}


def test_extract_gives_a_second_picture_with_the_same_name_a_name_of_its_own(tmp_path, writes):
    from xlsx_mcp.ops import objects

    book = hx.make_package(tmp_path / "book.xlsx", hx.base_members() + [
        ("xl/media/a/made_up.png", b"made-up one"), ("xl/media/b/made_up.png", b"made-up two"),
        ("xl/media/c/MADE_UP.png", b"made-up three")])
    out = tmp_path / "pictures"
    out.mkdir()
    got = objects.manage_image(str(book), "extract", out_dir=str(out))
    assert writes["outside"] == []
    assert got["count"] == 3
    assert sorted(p.name for p in out.iterdir()) == ["MADE_UP (3).png", "made_up (2).png",
                                                     "made_up.png"]
    assert (out / "made_up.png").read_bytes() == b"made-up one"
    assert (out / "made_up (2).png").read_bytes() == b"made-up two"
    assert got["renamed"] == {"xl/media/b/made_up.png": "made_up (2).png",
                              "xl/media/c/MADE_UP.png": "MADE_UP (3).png"}


def test_extract_skips_folder_entries(tmp_path, writes):
    from xlsx_mcp.ops import objects

    book = hx.make_package(tmp_path / "book.xlsx", hx.base_members() + [
        ("xl/media/", b""), ("xl/media/sub/", b""), ("xl/media/image1.png", b"made-up picture")])
    out = tmp_path / "pictures"
    out.mkdir()
    got = objects.manage_image(str(book), "extract", out_dir=str(out))
    assert [Path(p).name for p in got["extracted"]] == ["image1.png"]
    assert "renamed" not in got


def test_extract_writes_a_normal_picture_as_before(tmp_path, writes):
    from xlsx_mcp.ops import objects

    book = hx.make_package(tmp_path / "book.xlsx", hx.base_members() + [
        ("xl/media/image1.png", b"made-up picture one"),
        ("xl/media/image2.jpeg", b"made-up picture two")])
    out = tmp_path / "pictures"
    out.mkdir()
    got = objects.manage_image(str(book), "extract", out_dir=str(out))
    assert got == {"extracted": [str(out / "image1.png"), str(out / "image2.jpeg")], "count": 2}
    assert (out / "image2.jpeg").read_bytes() == b"made-up picture two"
    assert writes["outside"] == []


@pytest.mark.parametrize("member, plain", [
    ("xl/media/image1.png", "image1.png"),
    ("xl/media/sub\\image1.png", "image1.png"),
    ("xl/media/console.png", "console.png"),
    ("xl/media/com10.png", "com10.png"),
    ("xl/media/made up (1).png", "made up (1).png"),
    ("xl/media/\u00e9t\u00e9.png", "\u00e9t\u00e9.png"),
    ("xl/media/" + "m" * 251 + ".png", "m" * 251 + ".png"),
])
def test_the_file_name_rule_keeps_ordinary_names(member, plain):
    from xlsx_mcp.core import zipsafe

    assert zipsafe.media_file_name(member) == (plain, None)


def test_no_unsafe_case_is_an_ordinary_name():
    """The rule is the same on every system: a file extracted here may be
    copied to Windows later."""
    from xlsx_mcp.core import zipsafe

    for case, name in UNSAFE.items():
        got, why = zipsafe.media_file_name("xl/media/" + name.format(drive="Q"))
        assert got is None and why, case
