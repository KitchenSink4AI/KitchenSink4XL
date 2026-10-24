"""ops/annotations.py: comments and hyperlinks.

manage_comment writes LEGACY comments (the yellow sticky notes, openpyxl's
Comment model), which round-trip cleanly through a file-based save. THREADED
comments (the modern reply/resolve conversations) live in xl/threadedComments/,
a part openpyxl does not model and the hazard scan flags as a drop risk, so this
tool does not author threaded comments: a workbook that already holds them is
handled by the safety core (preserved on a loss-safe route or refused loudly),
never silently converted, and the reply and resolve actions those threads
support are deferred with a clear pointer rather than faked on the legacy model.

manage_hyperlink writes real cell hyperlinks (openpyxl's Hyperlink model) and,
on list, also surfaces HYPERLINK() formula links so an audit sees both kinds.
Mutations route through WorkbookPackage for the backup and verify.
"""

from __future__ import annotations

from typing import Any

from openpyxl.utils import get_column_letter

from ..core import authorship as _authorship
from ..core import limits as _limits
from ..core.errors import TargetNotFound, XlMcpError
from ..core.package import WorkbookPackage
from . import gridio

COMMENT_ACTIONS = ("add", "edit", "delete", "list")
HYPERLINK_ACTIONS = ("add", "remove", "list")
#: A note added without an author is signed by core.authorship: the user's
#: own name by default, "AI Assistant via KS4XL" with KS4XL_AUTHOR=assistant
#: or when the call passes author="assistant". An author that is an email
#: address is never written; the same module picks the name that signs
#: instead. Excel's legacy notes carry no initials field.


_TC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _resolve(base_dir: str, target: str) -> str:
    import posixpath
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def threaded_comments(path: str) -> dict[tuple[str, str], dict]:
    """The modern threaded comments, read from the
    package (openpyxl reads only Excel's legacy placeholder note, "[Threaded
    comment] Your version of Excel allows you to read this...", authored
    "tc={GUID}"). (sheet, cell) -> {author, text, date, replies}. Read-only
    and best effort: a part that does not parse reads as no threads."""
    import zipfile

    from ..core import xmlsafe as _xmlsafe

    def parse(part: str):
        return _xmlsafe.fromstring(zf.read(part), doc=path, part=part)

    out: dict[tuple[str, str], dict] = {}
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            if not any(n.startswith("xl/threadedComments/") for n in names):
                return out
            wb = parse("xl/workbook.xml")
            wrels = parse("xl/_rels/workbook.xml.rels")
            targets = {r.get("Id"): _resolve("xl", r.get("Target", ""))
                       for r in wrels.findall(f"{{{_REL_NS}}}Relationship")}
            people: dict[str, str] = {}
            for r in wrels.findall(f"{{{_REL_NS}}}Relationship"):
                if r.get("Type", "").endswith("/person"):
                    part = _resolve("xl", r.get("Target", ""))
                    if part in names:
                        for person in parse(part).findall(
                                f"{{{_TC_NS}}}person"):
                            people[person.get("id")] = person.get("displayName") or ""
            sheets = wb.find(f"{{{_MAIN_NS}}}sheets")
            for sh in (sheets if sheets is not None else []):
                title = sh.get("name")
                part = targets.get(sh.get(f"{{{_DOC_REL_NS}}}id"))
                if not part:
                    continue
                d, f = part.rsplit("/", 1)
                rels_part = f"{d}/_rels/{f}.rels"
                if rels_part not in names:
                    continue
                for r in parse(rels_part).findall(f"{{{_REL_NS}}}Relationship"):
                    if not r.get("Type", "").endswith("/threadedComment"):
                        continue
                    tc_part = _resolve(d, r.get("Target", ""))
                    if tc_part not in names:
                        continue
                    threads: dict[str, dict] = {}
                    for tc in parse(tc_part).findall(
                            f"{{{_TC_NS}}}threadedComment"):
                        text_el = tc.find(f"{{{_TC_NS}}}text")
                        item = {"author": people.get(tc.get("personId"), "") or None,
                                "text": text_el.text if text_el is not None else "",
                                "date": tc.get("dT")}
                        parent = tc.get("parentId")
                        if parent and parent in threads:
                            threads[parent]["replies"].append(item)
                        else:
                            threads[tc.get("id")] = {**item, "ref": tc.get("ref"),
                                                     "replies": []}
                    for th in threads.values():
                        ref = th.pop("ref")
                        out[(title, ref)] = th
    except Exception:  # noqa: BLE001 (a reader never fails on this part)
        return out
    return out


def _single(pkg, location, sheet):
    grid = pkg.resolve(location, default_sheet=sheet)
    if not grid.is_single:
        raise XlMcpError(
            f"this action needs a single cell; {grid.a1} is a range")
    return grid


def manage_comment(path: str, action: str, location: Any = None,
                   text: str | None = None, author: str | None = None,
                   sheet: str | None = None, allow_loss: bool = False,
                   backup: bool = True,
                   verify_com: bool | None = None) -> dict:
    """Add / edit / delete / list legacy cell comments. Backup + verify."""
    if action not in COMMENT_ACTIONS:
        raise XlMcpError(
            f"action must be one of {COMMENT_ACTIONS}, got {action!r}")
    if action in ("reply", "resolve"):  # defensive; not in the closed set
        raise XlMcpError(
            "reply and resolve are threaded-comment actions; this tool writes "
            "legacy notes only")

    if action == "list":
        wb = gridio.open_wb(path, data_only=False)
        try:
            threads = threaded_comments(path)
            out = []
            for ws in (wb.worksheets if sheet is None else [wb[sheet]]):
                cells = getattr(ws, "_cells", None) or {}
                seen = set()
                for (r, c), cell in cells.items():
                    cm = getattr(cell, "comment", None)
                    addr = f"{get_column_letter(c)}{r}"
                    th = threads.get((ws.title, addr))
                    if th is not None:
                        # The thread itself, not Excel's placeholder.
                        out.append({"sheet": ws.title, "cell": addr, "threaded": True, **th})
                        seen.add(addr)
                    elif cm is not None:
                        out.append({"sheet": ws.title, "cell": addr,
                                    "text": cm.text, "author": cm.author})
                for (title, addr), th in threads.items():
                    if title == ws.title and addr not in seen:
                        out.append({"sheet": title, "cell": addr, "threaded": True, **th})
            return {"comments": out, "count": len(out)}
        finally:
            wb.close()

    from openpyxl.comments import Comment
    pkg = WorkbookPackage.open(path)
    grid = _single(pkg, location, sheet)
    ws = pkg.workbook[grid.sheet]
    cell = ws.cell(grid.min_row, grid.min_col)

    note = None  # the sentence an email address sent as the author earns
    if action == "add":
        if not text:
            raise XlMcpError("add needs text")
        if cell.comment is not None:
            raise XlMcpError(
                f"{grid.a1} already has a comment; use action 'edit'")
        _limits.check_comment_text(text)
        signed, source, note = _authorship.signer_with_note(author)
        cell.comment = Comment(text, signed)
        detail = {"added": grid.a1, "author": signed, "author_source": source}
    elif action == "edit":
        if cell.comment is None:
            raise TargetNotFound(f"{grid.a1} has no comment to edit")
        new_text = text if text is not None else cell.comment.text
        _limits.check_comment_text(new_text)
        if author:
            signed, source, note = _authorship.signer_with_note(author)
        else:
            signed, source = cell.comment.author, "kept"
        cell.comment = Comment(new_text, signed)
        detail = {"edited": grid.a1, "author": signed, "author_source": source}
    else:  # delete
        if cell.comment is None:
            raise TargetNotFound(f"{grid.a1} has no comment to delete")
        cell.comment = None
        # Removing the sheet's last comment drops its comment part and VML
        # anchor; declare the deliberate removal so the default-fail inventory
        # verify does not read it as silent loss.
        pkg.expect_removal("xl/comments", "xl/drawings/commentsdrawing")
        detail = {"deleted": grid.a1}

    pkg._changed["comment"] = {"sheet": ws.title, **detail}
    result = pkg.save(allow_loss=allow_loss, backup=backup,
                      verify_com=verify_com)
    if note:
        result["warnings"] = [note, *result.get("warnings", [])]
    return result


def manage_hyperlink(path: str, action: str, location: Any = None,
                     target: str | None = None, display: str | None = None,
                     tooltip: str | None = None, sheet: str | None = None,
                     allow_loss: bool = False, backup: bool = True,
                     verify_com: bool | None = None) -> dict:
    """Add / remove / list cell hyperlinks (and audit HYPERLINK formulas).
    Backup + verify on write."""
    if action not in HYPERLINK_ACTIONS:
        raise XlMcpError(
            f"action must be one of {HYPERLINK_ACTIONS}, got {action!r}")

    if action == "list":
        wb = gridio.open_wb(path, data_only=False)
        try:
            out = []
            for ws in (wb.worksheets if sheet is None else [wb[sheet]]):
                cells = getattr(ws, "_cells", None) or {}
                for (r, c), cell in cells.items():
                    addr = f"{get_column_letter(c)}{r}"
                    hl = getattr(cell, "hyperlink", None)
                    if hl is not None:
                        out.append({"sheet": ws.title, "cell": addr,
                                    "kind": "hyperlink",
                                    "target": hl.target or hl.location,
                                    "tooltip": hl.tooltip})
                    v = cell.value
                    if isinstance(v, str) and v.upper().startswith(
                            "=HYPERLINK("):
                        out.append({"sheet": ws.title, "cell": addr,
                                    "kind": "formula", "target": v})
            return {"hyperlinks": out, "count": len(out)}
        finally:
            wb.close()

    pkg = WorkbookPackage.open(path)
    grid = _single(pkg, location, sheet)
    ws = pkg.workbook[grid.sheet]
    cell = ws.cell(grid.min_row, grid.min_col)

    hl_warnings: list[str] = []
    if action == "add":
        if not target:
            raise XlMcpError(
                "add needs target (a URL, or an in-workbook 'Sheet!A1' ref)")
        # A target Excel cannot parse makes the workbook unopenable, and an
        # unsafe scheme is the injection posture the server already takes on
        # cell text (core.limits).
        hl_warnings = _limits.check_hyperlink_target(target)
        if tooltip is not None:
            _limits.check_text_storable(tooltip, what="the hyperlink tooltip")
        from openpyxl.worksheet.hyperlink import Hyperlink
        internal = "!" in target and "://" not in target
        hl = Hyperlink(ref=grid.a1,
                       target=None if internal else target,
                       location=target if internal else None,
                       tooltip=tooltip)
        cell.hyperlink = hl
        if display is not None:
            cell.value = display
            pkg.note_region_write(ws.title, grid.min_row, grid.min_col,
                                  grid.min_row, grid.min_col)
        elif cell.value is None:
            cell.value = target
            pkg.note_region_write(ws.title, grid.min_row, grid.min_col,
                                  grid.min_row, grid.min_col)
        pkg._changed["hyperlink"] = {"sheet": ws.title, "added": grid.a1,
                                     "target": target}
    else:  # remove
        if cell.hyperlink is None:
            raise TargetNotFound(f"{grid.a1} has no hyperlink to remove")
        cell.hyperlink = None
        pkg._changed["hyperlink"] = {"sheet": ws.title, "removed": grid.a1}

    result = pkg.save(allow_loss=allow_loss, backup=backup,
                      verify_com=verify_com)
    if hl_warnings:
        result["warnings"] = list(result.get("warnings", [])) + hl_warnings
    return result


__all__ = [
    "manage_comment", "manage_hyperlink",
    "COMMENT_ACTIONS", "HYPERLINK_ACTIONS",
]
